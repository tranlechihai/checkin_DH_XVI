from flask import Flask, request, jsonify, send_from_directory,send_file
from flask_cors import CORS
import json
import os
import io
from datetime import datetime
from collections import Counter
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook


app = Flask(__name__, static_folder='static')
CORS(app)  # Cho phép các yêu cầu từ frontend

ALL_GENDER = ["Nam", "Nữ"]
ALL_PARTY_POS = [
    "Ủy viên Ban Chấp hành Đảng bộ",
    "Cấp ủy các chi bộ trực thuộc",
    "Không"
]
ALL_GOV_POS = [
    "Lãnh đạo trường",
    "Lãnh đạo phòng, khoa và tương đương",
    "Không"
]
ALL_PRO_LEVEL = [
    "Sinh viên",
    "Đại học",
    "Thạc sĩ",
    "Tiến sĩ",
    "Phó Giáo sư, Tiến sĩ",
    "Giáo sư, Tiến sĩ"
]
ALL_POLITIC_LEVEL = [
    "Cao cấp",
    "Trung cấp",
    "Đang học Trung cấp",
    "Chưa qua đào tạo"
]
ALL_AGE_GROUPS = [
    "Dưới 40 tuổi",
    "Từ 40 đến 50 tuổi",
    "Từ 51 đến 60 tuổi",
    "Trên 60 tuổi"
]


# Tạo mảng lưu trữ điểm danh tạm thời
DATA_FILE = os.path.join('static', 'dai_bieu_updated.json')
SUMMARY_FILE = os.path.join('static', 'wrapup.html')

@app.route('/')
def serve_index():
    return send_from_directory('static', 'index.html')

@app.route('/display')
def serve_display():
    return send_from_directory('static', 'display.html')

@app.route("/wrapup", methods=["GET","POST"])
def generate_summary():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            people = json.load(f)
            
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

    checked = [p for p in people if p.get("check_up")]
    # return jsonify(checked)
    def count_by_field(data, field, all_possible_values=None):
        counter = Counter(p.get(field, "Không rõ") for p in data)
        if all_possible_values:
            return {val: counter.get(val, 0) for val in all_possible_values}
        else:
            return dict(counter)

    def calc_age(date_str):
        birth_date = datetime.strptime(date_str, "%d/%m/%Y")

        return relativedelta(datetime.now(), birth_date)

    def age_group(age):
        years = age.years
        if years < 40:
            return "Dưới 40 tuổi"
        elif 40 <= years <= 50:
            return "Từ 40 đến 50 tuổi"
        elif 51 <= years <= 60:
            return "Từ 51 đến 60 tuổi"
        else:
            return "Trên 60 tuổi"

    gender_count = count_by_field(checked, "gioi", ALL_GENDER)
    party_pos_count = count_by_field(checked, "Chuc_vu_dang", ALL_PARTY_POS)
    gov_pos_count = count_by_field(checked, "Chuc_vu_chinh_quyen", ALL_GOV_POS)
    professional_level_count = count_by_field(checked, "trinh_do_chuyen_mon", ALL_PRO_LEVEL)
    political_theory_count = count_by_field(checked, "trinh_do_ly_luan_chinh_tri", ALL_POLITIC_LEVEL)

    ages = [(p["ho_va_ten"], calc_age(p["ngay_sinh"])) for p in checked if "ngay_sinh" in p]
    party_ages = [(p["ho_va_ten"], calc_age(p["ngay_du_bi"])) for p in checked if "ngay_du_bi" in p]
    age_groups = count_by_field(
    [{"group": age_group(calc_age(p["ngay_sinh"]))} for p in checked if "ngay_sinh" in p],
    "group",
    ALL_AGE_GROUPS)
    # Tuổi đời
    age_sorted = sorted(ages, key=lambda x: (x[1].years, x[1].months, x[1].days), reverse=True)
    if age_sorted:
        oldest_person = age_sorted[0]
        youngest_person = age_sorted[-1]
    else:
        oldest_person = youngest_person = ("Không có dữ liệu", relativedelta())

    # Tuổi Đảng
    party_ages_sorted = sorted(party_ages, key=lambda x: (x[1].years, x[1].months, x[1].days), reverse=True)
    if party_ages_sorted:
        max_party_age = party_ages_sorted[0]
        min_party_age = party_ages_sorted[-1]
    else:
        max_party_age = min_party_age = ("Không có dữ liệu", relativedelta())

    def format_table_html(counts, total):
        rows = ""
        for k, v in counts.items():
            percentage = f"{(v / total * 100):.1f}%"
            rows += f"<tr><td>{k}</td><td style='text-align:center'>{v}</td><td style='text-align:center'>{percentage}</td></tr>"
        return f"""
        <table border="1" cellspacing="0" cellpadding="6" style="width:100%; border-collapse: collapse; margin-bottom: 20px;">
            <thead style="background-color:#dbe9f7;">
                <tr>
                    <th>Nhóm</th>
                    <th style="text-align:center">Số lượng</th>
                    <th style="text-align:center">Tỉ lệ</th>
                </tr>
            </thead>
            <tbody>
                {rows}
            </tbody>
        </table>
        """


    
    checked_count = len(checked)
    total_count = len(people)
    ratio_present = f"{checked_count / total_count:.1%}"



    html = f"""
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <title>Thống kê Đại hội XVI</title>
        <link rel="icon" type="image/png" href="./static/images/logo_hcmue.ico">
        <style>
            body {{
                font-family: Arial, sans-serif;
                max-width: 900px;
                margin: auto;
                padding: 20px;
                background-color: #f9f9f9;
            }}
            h1 {{
                text-align: center;
                color: #2a4d9b;
            }}
            h2 {{
                margin-top: 40px;
                color: #1a3e6e;
                border-bottom: 2px solid #ccc;
                padding-bottom: 5px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 10px;
            }}
            th, td {{
                border: 1px solid #ccc;
                padding: 8px 12px;
                text-align: left;
            }}
            thead {{
                background-color: #e6f0ff;
            }}
            .highlight {{
                font-size: 20px;
                font-weight: bold;
                color: #b22222;
                margin-top: 10px;
            }}
            p {{
                margin-top: 10px;
            }}
        </style>
    </head>
    <body>
        <h1>THỐNG KÊ PHIÊN 1<br>ĐẠI HỘI ĐẠI BIỂU XVI</br></h1>

        <div style="text-align: right; margin-bottom: 10px;">
        <form action="/api/export_wrapup" method="get">
            <button type="submit" style="padding: 8px 12px; background-color: #2a4d9b; color: white; border: none; border-radius: 4px;">
            Xuất phiếu thống kê
            </button>
        </form>
        </div>

        <h2 style="color: #b22222; font-size: 20px;">
            Tổng số đại biểu đã điểm danh: {checked_count} / {total_count} ({ratio_present})
        </h2>

        <h2>Giới tính</h2>
        <ul>
            {format_table_html(gender_count, len(checked))}
        </ul>

        <h2>Chức vụ Đảng</h2>
        <ul>
            {format_table_html(party_pos_count, len(checked))}
        </ul>

        <h2>Chức vụ Chính quyền</h2>
        <ul>
            {format_table_html(gov_pos_count, len(checked))}
        </ul>

        <h2>Trình độ chuyên môn</h2>
        <ul>
            {format_table_html(professional_level_count, len(checked))}
        </ul>

        <h2>Trình độ lý luận chính trị</h2>
        <ul>
            {format_table_html(political_theory_count, len(checked))}
        </ul>

        <h2>Thống kê theo nhóm tuổi</h2>
        <ul>
            {format_table_html(age_groups, len(checked))}
        </ul>

        <h2>Tuổi đời</h2>
        <p>Đồng chí lớn tuổi nhất: <strong>{oldest_person[0]}</strong> ({oldest_person[1].years} năm {oldest_person[1].months} tháng {oldest_person[1].days} ngày)</p>
        <p>Đồng chí trẻ tuổi nhất: <strong>{youngest_person[0]}</strong> ({youngest_person[1].years} năm {youngest_person[1].months} tháng {youngest_person[1].days} ngày)</p>

        <h2>Tuổi Đảng</h2>
        <p>Đồng chí có tuổi Đảng cao nhất: <strong>{max_party_age[0]}</strong> ({max_party_age[1].years} năm {max_party_age[1].months} tháng {max_party_age[1].days} ngày)</p>
        <p>Đồng chí có tuổi Đảng thấp nhất: <strong>{min_party_age[0]}</strong> ({min_party_age[1].years} năm {min_party_age[1].months} tháng {min_party_age[1].days} ngày)</p>
    </body>
    </html>
    """
    try:
        with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
            f.write(html)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    return send_from_directory('static', 'wrapup.html')


@app.route("/api/update_check", methods=["POST"])
def update_check():
    data = request.json  # Nhận JSON từ client
    ma_dai_bieu = data.get("ma_dai_bieu")
    new_check_up = data.get("check_up")

    # Đọc dữ liệu hiện có
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            danh_sach = json.load(f)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

    # Cập nhật người tương ứng
    updated = False
    for person in danh_sach:
        if person["ma_dai_bieu"] == ma_dai_bieu:
            person["check_up"] = new_check_up
            updated = True
            break

    if not updated:
        return jsonify({"success": False, "error": "Không tìm thấy đại biểu"}), 404

    # Ghi lại file
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(danh_sach, f, ensure_ascii=False, indent=2)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

    return jsonify({"success": True})
@app.route('/api/reset_check', methods=['POST'])
def reset_check():
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)

        for p in data:
            p['check_up'] = False  # Reset về false

        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return jsonify(success=True)

    except Exception as e:
        return jsonify(success=False, error=str(e)), 500


@app.route("/api/export_vang", methods=["GET"])
def export_vang():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            people = json.load(f)

        vang_list = [p for p in people if not p.get("check_up")]

        wb = Workbook()
        ws = wb.active
        ws.title = "Đại biểu vắng"

        # Ghi tiêu đề
        headers = ["Mã", "Họ và tên", "Chi bộ", "Giới", "Ngày sinh", "Chức vụ Đảng", "Chức vụ Chính quyền","Trình độ chuyên môn","Trình độ lý luận chính trị","Check-up"]
        ws.append(headers)

        # Ghi dữ liệu
        for p in vang_list:
            ws.append([
                p.get("ma_dai_bieu", ""),
                p.get("ho_va_ten", ""),
                p.get("chi_bo", ""),
                p.get("gioi", ""),
                p.get("ngay_sinh", ""),
                p.get("Chuc_vu_dang", ""),
                p.get("Chuc_vu_chinh_quyen", ""),
                p.get("trinh_do_chuyen_mon", ""),
                p.get("trinh_do_ly_luan_chinh_tri", ""),
                

            ])

        # Trả file về client
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="DHXVI_danh_sach_vang_mat.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/export_comat", methods=["GET"])
def export_comat():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            people = json.load(f)

        vang_list = [p for p in people if p.get("check_up")]

        wb = Workbook()
        ws = wb.active
        ws.title = "Đại biểu có mặt"

        # Ghi tiêu đề
        headers = ["Mã", "Họ và tên", "Chi bộ", "Giới", "Ngày sinh", "Chức vụ Đảng", "Chức vụ Chính quyền","Trình độ chuyên môn","Trình độ lý luận chính trị","Check-up"]
        ws.append(headers)

        # Ghi dữ liệu
        for p in vang_list:
            ws.append([
                p.get("ma_dai_bieu", ""),
                p.get("ho_va_ten", ""),
                p.get("chi_bo", ""),
                p.get("gioi", ""),
                p.get("ngay_sinh", ""),
                p.get("Chuc_vu_dang", ""),
                p.get("Chuc_vu_chinh_quyen", ""),
                p.get("trinh_do_chuyen_mon", ""),
                p.get("trinh_do_ly_luan_chinh_tri", ""),
                

            ])

        # Trả file về client
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="DHXVI_danh_sach_co_mat.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/export_wrapup", methods=["GET"])
def export_wrapup_file():
    try:
        wrapup_path = os.path.join("static", "wrapup.html")
        return send_file(
            wrapup_path,
            as_attachment=True,
            download_name="DHXVI_phieu_thong_ke.html",
            mimetype="text/html"
        )
    except Exception as e:
        return jsonify(success=False, error=str(e)), 500


if __name__ == '__main__':
    app.run(debug=True)
