from flask import Flask, request, render_template, send_file
import pandas as pd
import os
import zipfile
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():

    file = request.files["file"]

    if file.filename == "":
        return "No file selected"

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)

    file.save(filepath)

    data = {}
    pairs = set()

    # read workbook
    xls = pd.ExcelFile(filepath)

    for sheet in xls.sheet_names:

        df = pd.read_excel(filepath, sheet_name=sheet)

        df.columns = df.columns.astype(str).str.strip().str.upper()

        if "PID" not in df.columns or "POC" not in df.columns:
            continue

        pid = df["PID"].fillna("")
        poc = df["POC"].fillna("")

        pairs.update(zip(pid, poc))
        data[sheet] = df

    if not pairs:
        return "No PID + POC pairs found."

    base_name = os.path.splitext(filename)[0]
    result_folder = os.path.join(OUTPUT_FOLDER, base_name)

    os.makedirs(result_folder, exist_ok=True)

    for pid, poc in pairs:

        if pid == "" and poc == "":
            continue

        output_file = f"Data_{pid}_{poc}.xlsx"
        out_path = os.path.join(result_folder, output_file)

        with pd.ExcelWriter(out_path) as writer:

            for sheet, df in data.items():

                filtered = df[(df["PID"] == pid) & (df["POC"] == poc)]

                if not filtered.empty:
                    filtered.to_excel(writer, sheet_name=sheet, index=False)

    zip_path = result_folder + ".zip"

    with zipfile.ZipFile(zip_path, "w") as z:
        for f in os.listdir(result_folder):
            z.write(os.path.join(result_folder, f), f)

    return send_file(zip_path, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)