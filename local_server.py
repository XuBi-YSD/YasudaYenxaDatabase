"""
local_server.py
================
May chu local toi thieu (chi dung thu vien co san cua Python, KHONG can
cai Flask/gi them ngoai openpyxl+pillow da co) de app HTML goi thang toi
va nhan ve file .xlsx that — bo qua hoan toan buoc tai JSON + chay lenh
tay.

CACH DUNG:
  1) Dat file nay CUNG THU MUC voi report_engine.py, json_to_excel.py,
     va 2 file mau .xlsx goc (doi ten dung nhu bien TEMPLATE_REPORT /
     TEMPLATE_PLAN ben duoi, hoac sua lai duong dan cho khop).
  2) Mo Command Prompt tai thu muc do, chay:
         python local_server.py
  3) Giu cua so do dang mo (dung tat), quay lai daily_report_app.html
     dang mo trong trinh duyet, bam nut "Xuat Excel truc tiep (local
     server)". File .xlsx se tu tai ve ngay, khong can qua JSON nua.
"""

import http.server
import json
import os
import socketserver
import sys
import tempfile
from pathlib import Path
from urllib.parse import urlparse

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
from json_to_excel import REPORT_FIELDS, PLAN_FIELDS, build_data_dict, _decode_photos, read_legend, read_options
from report_engine import generate_daily_report, generate_daily_plan

PORT = 8765

# QUAN TRONG: duong dan LUON tinh theo thu muc chua local_server.py
# (BASE_DIR), KHONG phai theo noi ban dang dung lenh `python`. Truoc day
# dung duong dan tuong doi don gian ("20260731-DR1.xlsx") nen neu terminal
# dang dung o mot thu muc khac (vd mo San lai, mo tu shortcut...), script
# se bao "khong tim thay file" DU file thuc su nam dung cho — day chinh
# la loi ban vua gap. Doi ten file cho khop 2 file mau ban dang dung,
# neu khac ten mac dinh:
TEMPLATE_REPORT = BASE_DIR / "20260731-DR1.xlsx"
TEMPLATE_PLAN = BASE_DIR / "20260731-DR2.xlsx"


class Handler(http.server.BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def _error(self, status, message):
        self.send_response(status)
        self._cors()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps({"error": message}, ensure_ascii=False).encode("utf-8"))

    def do_GET(self):
        path = urlparse(self.path).path
        try:
            if path == "/legend/report":
                if not TEMPLATE_REPORT.exists():
                    return self._error(500, f"Không tìm thấy file mẫu '{TEMPLATE_REPORT.name}'")
                data = read_legend(TEMPLATE_REPORT, REPORT_FIELDS)
            elif path == "/legend/plan":
                if not TEMPLATE_PLAN.exists():
                    return self._error(500, f"Không tìm thấy file mẫu '{TEMPLATE_PLAN.name}'")
                data = read_legend(TEMPLATE_PLAN, PLAN_FIELDS)
            elif path == "/options/report":
                if not TEMPLATE_REPORT.exists():
                    return self._error(500, f"Không tìm thấy file mẫu '{TEMPLATE_REPORT.name}'")
                data = read_options(TEMPLATE_REPORT, REPORT_FIELDS)
            elif path == "/options/plan":
                if not TEMPLATE_PLAN.exists():
                    return self._error(500, f"Không tìm thấy file mẫu '{TEMPLATE_PLAN.name}'")
                data = read_options(TEMPLATE_PLAN, PLAN_FIELDS)
            else:
                return self._error(404, f"Không có endpoint {path}")
        except Exception as e:
            return self._error(500, str(e))

        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self._cors()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        path = urlparse(self.path).path
        if path not in ("/convert/report", "/convert/plan"):
            return self._error(404, f"Không có endpoint {path}")

        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            payload = json.loads(body)
        except Exception as e:
            return self._error(400, f"JSON không hợp lệ: {e}")

        try:
            with tempfile.TemporaryDirectory() as tmp:
                out_path = os.path.join(tmp, "output.xlsx")
                if path == "/convert/report":
                    if not Path(TEMPLATE_REPORT).exists():
                        return self._error(500, f"Không tìm thấy file mẫu '{TEMPLATE_REPORT}' cùng thư mục với local_server.py")
                    fields, input_map, applied, skipped = build_data_dict(payload, REPORT_FIELDS, template_path=TEMPLATE_REPORT)
                    data = {REPORT_FIELDS["sheet"]: fields}
                    photos = _decode_photos(payload, tmp)
                    generate_daily_report(TEMPLATE_REPORT, out_path, data, photos, input_map=input_map)
                    filename = "daily_report_output.xlsx"
                else:
                    if not Path(TEMPLATE_PLAN).exists():
                        return self._error(500, f"Không tìm thấy file mẫu '{TEMPLATE_PLAN}' cùng thư mục với local_server.py")
                    fields, input_map, applied, skipped = build_data_dict(payload, PLAN_FIELDS, template_path=TEMPLATE_PLAN)
                    data = {PLAN_FIELDS["sheet"]: fields}
                    generate_daily_plan(TEMPLATE_PLAN, out_path, data, input_map=input_map)
                    filename = "daily_plan_output.xlsx"

                content = Path(out_path).read_bytes()
        except Exception as e:
            return self._error(500, str(e))

        self.send_response(200)
        self._cors()
        self.send_header("Content-Type",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("X-Applied-Fields", str(len(applied)))
        self.send_header("X-Skipped-Fields", str(len(skipped)))
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)
        print(f"[{filename}] đã ghi {len(applied)} trường, bỏ qua {len(skipped)} trường.")

    def log_message(self, fmt, *args):
        print("  ", fmt % args)


if __name__ == "__main__":
    print(f"📁 Thư mục đang dùng (nơi local_server.py nằm): {BASE_DIR}")
    missing = [t for t in (TEMPLATE_REPORT, TEMPLATE_PLAN) if not t.exists()]
    if missing:
        print("[CẢNH BÁO] Không thấy các file mẫu sau trong thư mục trên:")
        for m in missing:
            print("   -", m.name)
        print("           Copy 2 file .xlsx mẫu vào đúng thư mục in ở trên,")
        print("           hoặc sửa TEMPLATE_REPORT / TEMPLATE_PLAN cho đúng tên file bạn có.\n")
    else:
        print("✔ Tìm thấy đủ 2 file mẫu.\n")

    with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
        print(f"✔ Server local đang chạy: http://localhost:{PORT}")
        print("  Giữ cửa sổ này mở. Quay lại trình duyệt, mở daily_report_app.html,")
        print("  bấm nút 'Xuất Excel trực tiếp (local server)'.")
        print("  Nhấn Ctrl+C để dừng.\n")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nĐã dừng server.")
