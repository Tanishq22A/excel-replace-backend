from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
import os

app = Flask(__name__)
CORS(app)

# For simple project (single file in memory)
df_store = None


# -------------------------
# Upload Excel
# -------------------------
@app.route("/upload", methods=["POST"])
def upload():

    global df_store

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    try:
        df_store = pd.read_excel(file)
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    return jsonify({
        "columns": df_store.columns.tolist()
    })


# -------------------------
# Replace values
# -------------------------
@app.route("/replace", methods=["POST"])
def replace():

    global df_store

    if df_store is None:
        return jsonify({"error": "No file uploaded"}), 400

    data = request.get_json()

    mode = data.get("mode")
    find_val = str(data.get("find", "")).strip()
    replace_val = data.get("replace", "")
    column = data.get("column")

    if find_val == "":
        return jsonify({"error": "Find value is empty"}), 400

    # ---------------- find & replace (entire sheet)
    if mode == "find":

        def replace_cell(x):
            if pd.isna(x):
                return x
            if str(x).strip().lower() == find_val.lower():
                return replace_val
            return x

        df_store = df_store.applymap(replace_cell)

    # ---------------- row wise (column based)
    elif mode == "row":

        if column not in df_store.columns:
            return jsonify({"error": "Invalid column"}), 400

        def replace_cell_col(x):
            if pd.isna(x):
                return x
            if str(x).strip().lower() == find_val.lower():
                return replace_val
            return x

        df_store[column] = df_store[column].apply(replace_cell_col)

    else:
        return jsonify({"error": "Invalid mode"}), 400

    return jsonify({"status": "ok"})


# -------------------------
# Download Excel
# -------------------------
@app.route("/download/excel", methods=["GET"])
def download_excel():

    global df_store

    if df_store is None:
        return jsonify({"error": "No file"}), 400

    buf = BytesIO()
    df_store.to_excel(buf, index=False)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name="updated.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# -------------------------
# Download PDF
# -------------------------
@app.route("/download/pdf", methods=["GET"])
def download_pdf():

    global df_store

    if df_store is None:
        return jsonify({"error": "No file"}), 400

    buf = BytesIO()

    data = [df_store.columns.tolist()] + df_store.astype(str).values.tolist()

    doc = SimpleDocTemplate(buf, pagesize=A4)

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTSIZE", (0, 0), (-1, -1), 7)
    ]))

    doc.build([table])
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name="updated.pdf",
        mimetype="application/pdf"
    )


# -------------------------
# Run (Render compatible)
# -------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
