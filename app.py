from __future__ import annotations

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from flask import Flask, jsonify, render_template, request, send_file

from generate_report import TEMPLATE_NAME, build_report

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / TEMPLATE_NAME

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/api/generate")
@app.post("/generate")
def generate():
    uploaded_file = request.files.get("file")
    if uploaded_file is None or not uploaded_file.filename:
        return jsonify({"error": "Select an XLSX file."}), 400

    input_name = Path(uploaded_file.filename).name
    if Path(input_name).suffix.lower() != ".xlsx":
        return jsonify({"error": "Only .xlsx files are supported."}), 400

    if not TEMPLATE_PATH.is_file():
        return jsonify({"error": f"Template not found: {TEMPLATE_PATH.name}"}), 500

    try:
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            xlsx_path = temp_path / input_name
            uploaded_file.save(xlsx_path)

            output_path = build_report(xlsx_path, TEMPLATE_PATH)
            download_name = output_path.name
            output_bytes = output_path.read_bytes()
            response = send_file(
                BytesIO(output_bytes),
                as_attachment=True,
                download_name=download_name,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            response.headers["X-Output-Filename"] = download_name
            return response
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
