from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors

app = Flask(__name__)
CORS(app)

df_store = None


@app.route("/upload", methods=["POST"])
def upload():

    global df_store

    if "file" not in request.files:
        return jsonify({"error": "No file"}), 400

    file = request.files["file"]
    df_store = pd.read_excel(file)

    return jsonify({
        "columns": df_store.columns.tolist()
    })


@app.route("/replace", methods=["POST"])
def replace():

    global df_store

    if df_store is None:
        return jsonify({"error": "No file uploaded"}), 400

    data = request.json

    mode = data["mode"]
    find_val = data["find"]
    replace_val = data["replace"]
    column = data.get("column")

    if mode == "find":
        df_store = df_store.applymap(
            lambda x: replace_val if str(x) == find_val else x
        )

    elif mode == "row":
        if column not in df_store.columns:
            return jsonify({"error": "Invalid column"}), 400

        df_store[column] = df_store[column].apply(
            lambda x: replace_val if str(x) == find_val else x
        )

    return jsonify({"status": "ok"})


@app.route("/download/excel")
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
        download_name="updated.xlsx"
    )


@app.route("/download/pdf")
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
        download_name="updated.pdf"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
