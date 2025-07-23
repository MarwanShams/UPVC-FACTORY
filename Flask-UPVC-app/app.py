from flask import Flask, request, render_template
import pandas as pd

app = Flask(__name__)

EXCEL_FILE = 'Storage UPVC Factory.xlsx'
SHEET_NAME = 'Remaining'

def load_excel():
    """Reads and cleans the Excel file, returning a DataFrame."""
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        df = df.fillna("")
        # ✅ Clean 'Part. No.': remove .0 if float, strip spaces
        df["Part. No."] = df["Part. No."].apply(
            lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x)
        ).str.strip()
        return df
    except Exception as e:
        print("❌ Error reading Excel:", e)
        return None

@app.route('/')
def home():
    return """
    <h3>Welcome to the Product Lookup System</h3>
    <p>Use this format: <code>/product?id=YourPartNo</code></p>
    <p>Example: <a href="/product?id=202610">/product?id=202610</a></p>
    """

@app.route('/test')
def test():
    return "✅ Flask is running!"

@app.route('/product')
def product():
    product_id = request.args.get("id")
    if not product_id:
        return "<h3>❌ No product ID provided. Use ?id=YourPartNo</h3>"

    df = load_excel()
    if df is None:
        return "<h3>❌ Failed to load Excel file</h3>"

    if "Part. No." not in df.columns:
        return "<h3>❌ 'Part. No.' column not found in Excel</h3>"

    # ✅ Also clean the input ID
    product_id = str(product_id).strip()

    # ✅ Find the matching product
    product_row = df[df["Part. No."] == product_id]

    if product_row.empty:
        return f"<h3>❌ Product with Part. No. {product_id} not found</h3>"

    product_data = product_row.iloc[0].to_dict()
    return render_template("product.html", product=product_data)

if __name__ == '__main__':
    app.run(debug=True)
