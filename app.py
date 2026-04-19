from flask import Flask, render_template, request, jsonify
import pandas as pd
import os, json
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

ALLOWED = {'.xlsx', '.xls', '.csv'}
LATEST_FILE = os.path.join('uploads', '.latest')

def get_latest_path():
    if os.path.exists(LATEST_FILE):
        p = open(LATEST_FILE).read().strip()
        if os.path.exists(p):
            return p
    return None

def load_df(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)
    df.columns = df.columns.str.strip()
    return df

def detect_columns(df):
    cols = {c.lower(): c for c in df.columns}
    def find(candidates):
        for c in candidates:
            if c in cols: return cols[c]
        return None

    return {
        'date':     find(['order date','date','orderdate','order_date','transaction date']),
        'sales':    find(['sales','revenue','amount','total','sales amount','price']),
        'profit':   find(['profit','margin','net profit','profit amount']),
        'region':   find(['region','area','zone','territory']),
        'category': find(['category','product category','cat','type','product type']),
        'segment':  find(['segment','customer segment','customer type','cust segment']),
        'ship':     find(['ship mode','shipping mode','shipment mode','ship method','shipping']),
        'customer': find(['customer name','customer','client','client name']),
        'state':    find(['state','province','state/province']),
        'product':  find(['product id','product','product name','item','sku']),
        'order_id': find(['order id','order_id','orderid','order number']),
        'quantity': find(['quantity','qty','units','count']),
    }

def safe_val(v):
    if pd.isna(v): return 0
    return round(float(v), 2)

def process(df):
    m = detect_columns(df)
    out = {'columns': m, 'rows': len(df), 'col_names': list(df.columns)}

    # --- Sales column is mandatory ---
    if not m['sales']:
        return {'error': 'Could not find a Sales/Revenue column in your dataset.'}

    df['_sales'] = pd.to_numeric(df[m['sales']], errors='coerce').fillna(0)

    # --- Date parsing ---
    if m['date']:
        df['_date'] = pd.to_datetime(df[m['date']], errors='coerce', dayfirst=False)
        df['_year']    = df['_date'].dt.year
        df['_month']   = df['_date'].dt.month
        df['_quarter'] = df['_date'].dt.quarter
        years = sorted(df['_year'].dropna().unique().astype(int).tolist())
        out['years'] = years

        # Monthly by year
        monthly = {}
        for yr in years:
            sub = df[df['_year'] == yr]
            monthly[str(yr)] = sub.groupby('_month')['_sales'].sum().reindex(range(1,13), fill_value=0).round(2).tolist()
        out['monthly'] = monthly

        # Quarterly by year
        quarterly = {}
        for yr in years:
            sub = df[df['_year'] == yr]
            quarterly[str(yr)] = sub.groupby('_quarter')['_sales'].sum().reindex([1,2,3,4], fill_value=0).round(2).tolist()
        out['quarterly'] = quarterly

        # YoY comparison
        if len(years) >= 2:
            y1, y2 = years[-2], years[-1]
            r1 = df[df['_year']==y1]['_sales'].sum()
            r2 = df[df['_year']==y2]['_sales'].sum()
            out['yoy'] = round((r2-r1)/r1*100, 1) if r1 else 0
        else:
            out['yoy'] = None
    else:
        out['years'] = []
        out['monthly'] = {}
        out['quarterly'] = {}

    # --- KPIs ---
    out['total_revenue'] = round(df['_sales'].sum(), 2)
    out['total_quantity'] = int(df[m['quantity']].sum()) if m['quantity'] else None

    if m['profit']:
        df['_profit'] = pd.to_numeric(df[m['profit']], errors='coerce').fillna(0)
        out['total_profit'] = round(df['_profit'].sum(), 2)
        out['profit_margin'] = round(out['total_profit']/out['total_revenue']*100, 1) if out['total_revenue'] else 0
        out['profit_by_cat'] = df.groupby(df[m['category']])['_profit'].sum().round(2).to_dict() if m['category'] else {}
    else:
        out['total_profit'] = None
        out['profit_margin'] = None
        out['profit_by_cat'] = {}

    if m['order_id']:
        out['total_orders'] = int(df[m['order_id']].nunique())
        out['aov'] = round(out['total_revenue'] / out['total_orders'], 2) if out['total_orders'] else 0
    else:
        out['total_orders'] = len(df)
        out['aov'] = round(out['total_revenue'] / len(df), 2)

    # --- Breakdowns ---
    def top_dict(col, n=None):
        if not col: return {}
        g = df.groupby(df[col])['_sales'].sum().round(2)
        if n: g = g.nlargest(n)
        return g.to_dict()

    out['by_region']   = top_dict(m['region'])
    out['by_category'] = top_dict(m['category'])
    out['by_segment']  = top_dict(m['segment'])
    out['by_ship']     = top_dict(m['ship'])
    out['by_state']    = top_dict(m['state'], 7)
    out['by_customer'] = top_dict(m['customer'], 7)

    # --- Discount analysis ---
    disc_col = next((c for c in df.columns if 'discount' in c.lower()), None)
    if disc_col:
        df['_disc'] = pd.to_numeric(df[disc_col], errors='coerce').fillna(0)
        out['avg_discount'] = round(df['_disc'].mean() * 100, 1)
    else:
        out['avg_discount'] = None

    return out

@app.route('/')
def index():
    path = get_latest_path()
    data = None
    filename = None
    if path:
        try:
            df = load_df(path)
            data = process(df)
            filename = os.path.basename(path)
        except Exception as e:
            data = {'error': str(e)}
    return render_template('index.html', data=json.dumps(data), filename=filename)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    f = request.files['file']
    if not f.filename:
        return jsonify({'error': 'No file selected'}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED:
        return jsonify({'error': f'Unsupported file type: {ext}. Use .xlsx, .xls, or .csv'}), 400
    fname = secure_filename(f.filename)
    path = os.path.join(app.config['UPLOAD_FOLDER'], fname)
    f.save(path)
    open(LATEST_FILE, 'w').write(path)
    try:
        df = load_df(path)
        data = process(df)
        return jsonify({'success': True, 'filename': fname, 'data': data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/data')
def get_data():
    path = get_latest_path()
    if not path:
        return jsonify({'error': 'No dataset uploaded yet'})
    try:
        df = load_df(path)
        return jsonify(process(df))
    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    print("\n✅  Sales Dashboard running at → http://127.0.0.1:5000\n")
    app.run(debug=True, port=5000)
