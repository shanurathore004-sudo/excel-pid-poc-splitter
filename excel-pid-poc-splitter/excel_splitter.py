from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd
import zipfile
import os
import uuid

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")
def home():
    return "Excel PID-POC Splitter API Running"


@app.route("/upload", methods=["POST"])
def upload():

    if "file" not in request.files:
        return {"error": "No file uploaded"}

    file = request.files["file"]

    unique_id = str(uuid.uuid4())
    filepath = os.path.join(UPLOAD_FOLDER, unique_id + "_" + file.filename)

    file.save(filepath)

    xls = pd.ExcelFile(filepath)

    output_files = []

    for sheet in xls.sheet_names:

        df = pd.read_excel(filepath, sheet_name=sheet)

        columns = [c.lower() for c in df.columns]

        if "pid" not in columns or "poc" not in columns:
            continue

        pid_col = df.columns[columns.index("pid")]
        poc_col = df.columns[columns.index("poc")]

        grouped = df.groupby([pid_col, poc_col])

        for (pid, poc), data in grouped:

            safe_pid = str(pid).replace("/", "_")
            safe_poc = str(poc).replace("/", "_")

            filename = f"{sheet}_{safe_pid}_{safe_poc}.xlsx"
            filepath_out = os.path.join(OUTPUT_FOLDER, filename)

            data.to_excel(filepath_out, index=False)

            output_files.append(filepath_out)

    if not output_files:
        return {"error": "No PID + POC columns found"}

    zip_name = os.path.join(OUTPUT_FOLDER, unique_id + ".zip")

    with zipfile.ZipFile(zip_name, "w") as z:
        for f in output_files:
            z.write(f, os.path.basename(f))

    return send_file(zip_name, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)