from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import zipfile
import io
import os

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})


@app.route("/")
def home():
    return "Excel PID-POC Splitter API Running"


@app.route("/upload", methods=["POST"])
def upload():

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    try:

        xls = pd.ExcelFile(file)

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:

            for sheet in xls.sheet_names:

                df = xls.parse(sheet)

                columns = [c.lower() for c in df.columns]

                if "pid" not in columns or "poc" not in columns:
                    continue

                pid_col = df.columns[columns.index("pid")]
                poc_col = df.columns[columns.index("poc")]

                grouped = df.groupby([pid_col, poc_col])

                for (pid, poc), data in grouped:

                    filename = f"{sheet}_{pid}_{poc}.xlsx"

                    excel_buffer = io.BytesIO()

                    data.to_excel(excel_buffer, index=False)

                    z.writestr(filename, excel_buffer.getvalue())

        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name="split_files.zip"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
