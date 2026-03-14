from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
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

        df = pd.read_excel(file)

        columns = [c.lower() for c in df.columns]

        if "pid" not in columns or "poc" not in columns:
            return jsonify({"error": "PID or POC column missing"}), 400

        pid_col = df.columns[columns.index("pid")]
        poc_col = df.columns[columns.index("poc")]

        grouped = df.groupby([pid_col, poc_col])

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            for (pid, poc), data in grouped:

                sheet_name = f"{pid}_{poc}"

                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]

                data.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="split_files.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
