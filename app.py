"""
app.py
======
Ban WEB (Flask) cua local_server.py — deploy len dich vu host Python mien
phi (vd Render.com), ket noi voi GitHub de tu dong deploy moi lan cap
nhat code. Dung LAI 100% logic da test trong report_engine.py va
json_to_excel.py — chi doi lop giao tiep HTTP tu http.server sang Flask
cho tuong thich voi cac dich vu hosting pho bien.

CHAY THU O MAY (truoc khi deploy):
    pip install flask openpyxl pillow
    python app.py
    (mo http://localhost:8765 de kiem tra)

DEPLOY LEN RENDER.COM (mien phi):
    1) Day toan bo thu muc nay (ca 20260731-DR1.xlsx / DR2.xlsx) len GitHub.
    2) Vao https://render.com -> New -> Web Service -> chon repo GitHub nay.
    3) Build Command:  pip install -r requirements.txt
       Start Command:  gunicorn app:app
    4) Deploy xong se co URL dang https://ten-app.onrender.com
    5) Sua LOCAL_SERVER_BASE trong daily_report_app.html thanh URL do.
"""

import os
import tempfile
from pathlib import Path

from flask import Flask, request, jsonify, send_file, Response

from json_to_excel import REPORT_FIELDS, PLAN_FIELDS, build_data_dict, _decode_photos, read_legend
from report_engine import generate_daily_report, generate_daily_plan

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_REPORT = BASE_DIR / "20260731-DR1.xlsx"
TEMPLATE_PLAN = BASE_DIR / "20260731-DR2.xlsx"

app = Flask(__name__)


def _cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.after_request
def add_cors(resp):
    return _cors(resp)


@app.route("/legend/report", methods=["GET"])
def legend_report():
    if not TEMPLATE_REPORT.exists():
        return jsonify({"error": f"Không tìm thấy file mẫu '{TEMPLATE_REPORT.name}'"}), 500
    return jsonify(read_legend(str(TEMPLATE_REPORT), REPORT_FIELDS))


@app.route("/legend/plan", methods=["GET"])
def legend_plan():
    if not TEMPLATE_PLAN.exists():
        return jsonify({"error": f"Không tìm thấy file mẫu '{TEMPLATE_PLAN.name}'"}), 500
    return jsonify(read_legend(str(TEMPLATE_PLAN), PLAN_FIELDS))


@app.route("/convert/report", methods=["POST", "OPTIONS"])
def convert_report_route():
    if request.method == "OPTIONS":
        return _cors(Response(status=204))
    if not TEMPLATE_REPORT.exists():
        return jsonify({"error": f"Không tìm thấy file mẫu '{TEMPLATE_REPORT.name}'"}), 500
    payload = request.get_json(force=True)
    try:
        fields, input_map, applied, skipped = build_data_dict(
            payload, REPORT_FIELDS, template_path=str(TEMPLATE_REPORT))
        data = {REPORT_FIELDS["sheet"]: fields}
        with tempfile.TemporaryDirectory() as tmp:
            photos = _decode_photos(payload, tmp)
            out_path = os.path.join(tmp, "output.xlsx")
            generate_daily_report(str(TEMPLATE_REPORT), out_path, data, photos, input_map=input_map)
            content = Path(out_path).read_bytes()
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    resp = Response(content,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["Content-Disposition"] = 'attachment; filename="daily_report_output.xlsx"'
    resp.headers["X-Applied-Fields"] = str(len(applied))
    resp.headers["X-Skipped-Fields"] = str(len(skipped))
    return resp


@app.route("/convert/plan", methods=["POST", "OPTIONS"])
def convert_plan_route():
    if request.method == "OPTIONS":
        return _cors(Response(status=204))
    if not TEMPLATE_PLAN.exists():
        return jsonify({"error": f"Không tìm thấy file mẫu '{TEMPLATE_PLAN.name}'"}), 500
    payload = request.get_json(force=True)
    if payload.get("photos"):
        return jsonify({"error": "Daily Plan không được đính kèm ảnh."}), 400
    try:
        fields, input_map, applied, skipped = build_data_dict(
            payload, PLAN_FIELDS, template_path=str(TEMPLATE_PLAN))
        data = {PLAN_FIELDS["sheet"]: fields}
        with tempfile.TemporaryDirectory() as tmp:
            out_path = os.path.join(tmp, "output.xlsx")
            generate_daily_plan(str(TEMPLATE_PLAN), out_path, data, input_map=input_map)
            content = Path(out_path).read_bytes()
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    resp = Response(content,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["Content-Disposition"] = 'attachment; filename="daily_plan_output.xlsx"'
    resp.headers["X-Applied-Fields"] = str(len(applied))
    resp.headers["X-Skipped-Fields"] = str(len(skipped))
    return resp


@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "template_report_found": TEMPLATE_REPORT.exists(),
        "template_plan_found": TEMPLATE_PLAN.exists(),
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8765))
    print(f"📁 Thư mục đang dùng: {BASE_DIR}")
    print(f"✔ Template report: {'tìm thấy' if TEMPLATE_REPORT.exists() else 'THIẾU'}")
    print(f"✔ Template plan: {'tìm thấy' if TEMPLATE_PLAN.exists() else 'THIẾU'}")
    app.run(host="0.0.0.0", port=port, debug=False)
