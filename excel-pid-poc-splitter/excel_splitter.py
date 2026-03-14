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

                cols = [c.lower() for c in df.columns]

                if "pid" not in cols or "poc" not in cols:
                    continue

                pid_col = df.columns[cols.index("pid")]
                poc_col = df.columns[cols.index("poc")]

                for (pid, poc), data in df.groupby([pid_col, poc_col]):

                    filename = f"{sheet}_{pid}_{poc}.xlsx"

                    excel_buffer = io.BytesIO()

                    data.to_excel(
                        excel_buffer,
                        index=False,
                        engine="xlsxwriter"
                    )

                    z.writestr(filename, excel_buffer.getvalue())

                    del data

                del df

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
