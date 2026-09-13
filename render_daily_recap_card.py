# -*- coding: utf-8 -*-
"""
Module Render Bảng Tổng Hợp Chốt Phạt Điểm Danh Daily Dạng Ảnh Sang Xịn
Sử dụng Playwright để chụp ảnh thẻ Dashboard HTML cực kỳ trực quan, dễ view trên điện thoại và máy tính.
"""

import os
import sys
from datetime import datetime, date
import json
from playwright.sync_api import sync_playwright
from PIL import Image
import requests

from diem_danh_bot import get_vn_today, get_vn_now, load_config, get_db, get_m5_required_ams
from sync_attendance_sheets import get_admin_excused_ams, SPREADSHEET_ID

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_HTML = os.path.join(BASE_DIR, "recap_daily_attendance.html")
OUTPUT_PNG = os.path.join(BASE_DIR, "recap_daily_attendance.png")

def build_daily_recap_data(target_date: date = None):
    if target_date is None:
        target_date = get_vn_today()
        
    date_str = target_date.strftime("%Y-%m-%d")
    date_display = target_date.strftime("%d/%m/%Y")

    # Tự động đồng bộ từ Google Sheet về SQLite để dữ liệu luôn chính xác 100%
    try:
        from sync_attendance_sheets import restore_db_from_sheet
        restore_db_from_sheet(target_date)
    except Exception as e:
        print(f"⚠️ restore_db_from_sheet error: {e}")

    config = load_config()
    active_ams = [am for am in config["ams"] if am.get("is_active", True)]
    
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT am_id, milestone_id, status, submitted_at, late_minutes, penalty_amount FROM attendance_records WHERE date = ?", (date_str,))
        records = {}
        for row in cur.fetchall():
            records[(row["am_id"], row["milestone_id"])] = dict(row)
            
        cur.execute("SELECT am_id, am_name, milestone_id, reason, excuse_type FROM excuses WHERE date = ?", (date_str,))
        excuses_map = {}
        for row in cur.fetchall():
            aid = row["am_id"]
            if aid not in excuses_map:
                excuses_map[aid] = {"am_name": row["am_name"], "milestones": set(), "exemptions": set(), "reasons": []}
            excuses_map[aid]["milestones"].add(row["milestone_id"])
            if row["excuse_type"] == "EXEMPTION":
                excuses_map[aid]["exemptions"].add(row["milestone_id"])
            if row["reason"] and row["reason"] not in excuses_map[aid]["reasons"]:
                excuses_map[aid]["reasons"].append(row["reason"])

    excused_emp_ids = set()
    try:
        excused_emp_ids = get_admin_excused_ams(target_date)
    except Exception:
        pass

    m5_required = get_m5_required_ams()

    rows = []
    total_region_fine = 0
    clean_am_count = 0
    violation_am_count = 0

    for am in active_ams:
        am_id = am["id"]
        emp_id = str(am.get("employee_id", ""))
        am_name = am["display_name"]
        
        is_admin_excused = emp_id in excused_emp_ids
        am_fine = 0
        violations = []
        m_statuses = {} # m_id -> dict(text, badge_class)

        for m_id in (1, 2, 3, 4):
            rec = records.get((am_id, m_id))
            has_excuse = am_id in excuses_map and (m_id in excuses_map[am_id]["milestones"] or None in excuses_map[am_id]["milestones"])
            is_exempt = am_id in excuses_map and (m_id in excuses_map[am_id]["exemptions"] or None in excuses_map[am_id]["exemptions"])

            if is_admin_excused:
                m_statuses[m_id] = {"text": "🏖️ Nghỉ phép", "class": "badge-gray"}
            elif is_exempt:
                m_statuses[m_id] = {"text": "🛡️ Miễn nộp", "class": "badge-blue"}
            elif not rec or rec["status"] == "NOT_SUBMITTED":
                am_fine += config["fines"]["not_submitted"]
                if has_excuse:
                    m_statuses[m_id] = {"text": "⏳ Xin trễ (chưa nộp)", "class": "badge-yellow"}
                    violations.append(f"M{m_id} chưa nộp (100k)")
                else:
                    m_statuses[m_id] = {"text": "❌ Chưa nộp", "class": "badge-red"}
                    violations.append(f"M{m_id} chưa nộp (100k)")
            elif rec["status"] == "ON_TIME":
                sub_time = rec["submitted_at"].split()[1][:5] if rec.get("submitted_at") else "Đúng hạn"
                m_statuses[m_id] = {"text": f"✅ {sub_time}", "class": "badge-green"}
            elif rec["status"] == "LATE":
                sub_time = rec["submitted_at"].split()[1][:5] if rec.get("submitted_at") else ""
                late_m = rec.get("late_minutes", 0)
                if has_excuse or rec.get("penalty_amount", 0) == 0:
                    m_statuses[m_id] = {"text": f"⚠️ Trễ {late_m}p (0đ)", "class": "badge-orange"}
                else:
                    am_fine += config["fines"]["late"]
                    m_statuses[m_id] = {"text": f"⚠️ Trễ {late_m}p (50k)", "class": "badge-red-soft"}
                    violations.append(f"M{m_id} trễ {late_m}p (50k)")
            elif rec["status"] == "INVALID":
                am_fine += config["fines"]["invalid"]
                m_statuses[m_id] = {"text": "🚫 Sai mẫu", "class": "badge-red"}
                violations.append(f"M{m_id} sai mẫu (200k)")

        # Mốc 5
        if am_id in m5_required:
            rec_m5 = records.get((am_id, 5))
            has_excuse_5 = am_id in excuses_map and (5 in excuses_map[am_id]["milestones"] or None in excuses_map[am_id]["milestones"])
            is_exempt_5 = am_id in excuses_map and (5 in excuses_map[am_id]["exemptions"] or None in excuses_map[am_id]["exemptions"])

            if is_admin_excused:
                m_statuses[5] = {"text": "🏖️ Nghỉ phép", "class": "badge-gray"}
            elif is_exempt_5:
                m_statuses[5] = {"text": "🛡️ Miễn nộp", "class": "badge-blue"}
            elif not rec_m5 or rec_m5["status"] == "NOT_SUBMITTED":
                am_fine += config["fines"]["not_submitted"]
                m_statuses[5] = {"text": "❌ Chưa nộp", "class": "badge-red"}
                violations.append("M5 chưa nộp (100k)")
            elif rec_m5["status"] == "ON_TIME":
                sub_time = rec_m5["submitted_at"].split()[1][:5] if rec_m5.get("submitted_at") else "Đúng hạn"
                m_statuses[5] = {"text": f"✅ {sub_time}", "class": "badge-green"}
            elif rec_m5["status"] == "LATE":
                late_m = rec_m5.get("late_minutes", 0)
                if has_excuse_5 or rec_m5.get("penalty_amount", 0) == 0:
                    m_statuses[5] = {"text": f"⚠️ Trễ {late_m}p (0đ)", "class": "badge-orange"}
                else:
                    am_fine += config["fines"]["late"]
                    m_statuses[5] = {"text": f"⚠️ Trễ {late_m}p (50k)", "class": "badge-red-soft"}
                    violations.append(f"M5 trễ {late_m}p (50k)")
        else:
            m_statuses[5] = {"text": "—", "class": "badge-none"}

        if is_admin_excused:
            am_fine = 0
            violations = ["🏖️ Admin miễn phạt / Nghỉ phép"]

        total_region_fine += am_fine
        if am_fine > 0:
            violation_am_count += 1
        else:
            clean_am_count += 1

        excuse_desc = ""
        if am_id in excuses_map:
            reasons = excuses_map[am_id]["reasons"]
            if reasons:
                excuse_desc = f"📝 {', '.join(reasons)}"

        rows.append({
            "am_name": am_name,
            "am_short": am.get("display_name", am_name),
            "m1": m_statuses.get(1),
            "m2": m_statuses.get(2),
            "m3": m_statuses.get(3),
            "m4": m_statuses.get(4),
            "m5": m_statuses.get(5),
            "fine": am_fine,
            "violations": ", ".join(violations) if violations else "🌟 Hoàn thành chuẩn 100%",
            "excuse_desc": excuse_desc
        })

    # Sắp xếp phạt nhiều nhất lên đầu
    rows.sort(key=lambda x: x["fine"], reverse=True)

    return {
        "date_display": date_display,
        "total_fine": total_region_fine,
        "total_ams": len(active_ams),
        "clean_ams": clean_am_count,
        "violation_ams": violation_am_count,
        "rows": rows,
        "generated_at": get_vn_now().strftime("%H:%M - %d/%m/%Y")
    }

def generate_daily_recap_html(data: dict) -> str:
    rows_html = []
    for idx, r in enumerate(data["rows"], 1):
        fine_class = "fine-zero" if r["fine"] == 0 else "fine-positive"
        fine_text = f"{r['fine']:,.0f}đ" if r["fine"] > 0 else "0đ"
        tr_class = "tr-clean" if r["fine"] == 0 else ("tr-warning" if idx % 2 == 0 else "tr-warning-alt")

        note_html = ""
        if r["fine"] > 0:
            note_html = f'<div class="violation-text">🚫 {r["violations"]}</div>'
        else:
            note_html = f'<div class="clean-text">🌟 {r["violations"]}</div>'
            
        if r["excuse_desc"]:
            note_html += f'<div class="excuse-text">{r["excuse_desc"]}</div>'

        rows_html.append(f"""
        <tr class="{tr_class}">
          <td class="center font-bold">{idx}</td>
          <td class="am-name font-bold">{r['am_name']}</td>
          <td class="center"><span class="{r['m1']['class']}">{r['m1']['text']}</span></td>
          <td class="center"><span class="{r['m2']['class']}">{r['m2']['text']}</span></td>
          <td class="center"><span class="{r['m3']['class']}">{r['m3']['text']}</span></td>
          <td class="center"><span class="{r['m4']['class']}">{r['m4']['text']}</span></td>
          <td class="center"><span class="{r['m5']['class']}">{r['m5']['text']}</span></td>
          <td class="right font-black {fine_class}">{fine_text}</td>
        </tr>
        """)

    html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }}
  body {{ margin: 0; padding: 20px; background: #0f172a; display: flex; justify-content: center; }}
  .card-container {{
    width: 1060px;
    background: #ffffff;
    border-radius: 18px;
    overflow: hidden;
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.45);
  }}

  /* Header */
  .header {{
    background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
    color: white;
    padding: 24px 32px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 3px solid #3b82f6;
  }}
  .brand-title {{ font-size: 24px; font-weight: 900; letter-spacing: -0.5px; display: flex; align-items: center; gap: 10px; }}
  .brand-subtitle {{ font-size: 13px; color: #94a3b8; font-weight: 600; margin-top: 4px; }}
  .date-badge {{
    background: rgba(59, 130, 246, 0.18);
    border: 1px solid #3b82f6;
    color: #60a5fa;
    padding: 8px 18px;
    border-radius: 10px;
    font-size: 14px;
    font-weight: 800;
  }}

  /* Metrics Summary Bar */
  .metrics-bar {{
    display: flex;
    background: #f8fafc;
    border-bottom: 1px solid #e2e8f0;
    padding: 18px 32px;
    gap: 20px;
  }}
  .metric-card {{
    flex: 1;
    background: #ffffff;
    border-radius: 12px;
    padding: 14px 18px;
    border: 1px solid #e2e8f0;
    display: flex;
    flex-direction: column;
    gap: 4px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.03);
  }}
  .metric-label {{ font-size: 12px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; }}
  .metric-value {{ font-size: 24px; font-weight: 900; color: #0f172a; }}
  .text-danger {{ color: #dc2626 !important; }}
  .text-success {{ color: #16a34a !important; }}
  .text-primary {{ color: #2563eb !important; }}

  /* Table */
  .table-wrapper {{ padding: 24px 32px 30px; }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13.5px;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    overflow: hidden;
  }}
  th {{
    background: #1e293b;
    color: #ffffff;
    font-weight: 800;
    font-size: 12.5px;
    text-transform: uppercase;
    letter-spacing: 0.4px;
    padding: 12px 10px;
    border-right: 1px solid rgba(255,255,255,0.1);
  }}
  th:last-child {{ border-right: none; }}
  td {{
    padding: 10px 10px;
    border-bottom: 1px solid #e2e8f0;
    border-right: 1px solid #f1f5f9;
    color: #1e293b;
    vertical-align: middle;
  }}
  td:last-child {{ border-right: none; }}

  /* Row styles */
  .tr-warning {{ background-color: #ffffff; }}
  .tr-warning-alt {{ background-color: #fcfdfe; }}
  .tr-clean {{ background-color: #f0fdf4; }}

  .center {{ text-align: center; }}
  .right {{ text-align: right; }}
  .font-bold {{ font-weight: 700; }}
  .font-black {{ font-weight: 900; }}

  .am-name {{ font-size: 14.5px; color: #0f172a; white-space: nowrap; }}
  .fine-positive {{ color: #dc2626; font-size: 14.5px; }}
  .fine-zero {{ color: #16a34a; font-size: 14px; }}

  /* Badges */
  .badge-green {{
    background: #dcfce7;
    color: #15803d;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 12px;
    white-space: nowrap;
    border: 1px solid #bbf7d0;
  }}
  .badge-red {{
    background: #fee2e2;
    color: #b91c1c;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 12px;
    white-space: nowrap;
    border: 1px solid #fecaca;
  }}
  .badge-red-soft {{
    background: #fff1f2;
    color: #e11d48;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 12px;
    white-space: nowrap;
    border: 1px solid #ffe4e6;
  }}
  .badge-orange {{
    background: #ffedd5;
    color: #c2410c;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 12px;
    white-space: nowrap;
    border: 1px solid #fed7aa;
  }}
  .badge-yellow {{
    background: #fef9c3;
    color: #a16207;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 12px;
    white-space: nowrap;
    border: 1px solid #fef08a;
  }}
  .badge-blue {{
    background: #e0f2fe;
    color: #0369a1;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 12px;
    white-space: nowrap;
    border: 1px solid #bae6fd;
  }}
  .badge-gray {{
    background: #f1f5f9;
    color: #64748b;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 12px;
    white-space: nowrap;
    border: 1px solid #e2e8f0;
  }}
  .badge-none {{ color: #94a3b8; font-weight: 600; }}

  .violation-text {{ font-size: 12px; color: #b91c1c; font-weight: 600; line-height: 1.35; }}
  .clean-text {{ font-size: 12px; color: #16a34a; font-weight: 700; }}
  .excuse-text {{ font-size: 11.5px; color: #0284c7; font-style: italic; margin-top: 3px; }}

  /* Footer */
  .footer {{
    background: #f8fafc;
    padding: 16px 32px;
    border-top: 1px solid #e2e8f0;
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 12px;
    color: #64748b;
  }}
  .footer-link {{ color: #2563eb; font-weight: 700; text-decoration: none; }}
</style>
</head>
<body>
<div class="card-container">
  <div class="header">
    <div class="brand-title">📊 BẢNG TỔNG HỢP ĐIỂM DANH & CHỐT PHẠT VÙNG NTB</div>
    <div class="date-badge">📅 NGÀY {data['date_display']}</div>
  </div>

  <div class="metrics-bar">
    <div class="metric-card">
      <div class="metric-label">Tổng Tiền Phạt Vùng</div>
      <div class="metric-value text-danger">{data['total_fine']:,.0f} VNĐ</div>
    </div>
    <div class="metric-card">
      <div class="metric-label">Số AM Vi Phạm</div>
      <div class="metric-value text-danger">{data['violation_ams']} / {data['total_ams']} AM</div>
    </div>
    <div class="metric-card">
      <div class="metric-label">AM Chuẩn 100% (0đ)</div>
      <div class="metric-value text-success">{data['clean_ams']} AM</div>
    </div>
    <div class="metric-card">
      <div class="metric-label">Tỷ Lệ Tuân Thủ</div>
      <div class="metric-value text-primary">{((data['clean_ams']/data['total_ams'])*100):.1f}%</div>
    </div>
  </div>

  <div class="table-wrapper">
    <table>
      <thead>
        <tr>
          <th style="width: 45px;">STT</th>
          <th style="width: 200px;">AM Quản Lý</th>
          <th style="width: 110px;">M1: Đầu Ngày<br><span style="font-size: 10px; font-weight: normal; opacity: 0.8;">(08:00)</span></th>
          <th style="width: 110px;">M2: TTS Ca 1<br><span style="font-size: 10px; font-weight: normal; opacity: 0.8;">(11:00)</span></th>
          <th style="width: 110px;">M3: TTS Ca 2<br><span style="font-size: 10px; font-weight: normal; opacity: 0.8;">(16:00)</span></th>
          <th style="width: 110px;">M4: LTC TTS<br><span style="font-size: 10px; font-weight: normal; opacity: 0.8;">(20:00)</span></th>
          <th style="width: 110px;">M5: GTC &lt;50%<br><span style="font-size: 10px; font-weight: normal; opacity: 0.8;">(10:00)</span></th>
          <th style="width: 120px;">Tổng Phạt</th>
        </tr>
      </thead>
      <tbody>
        {"".join(rows_html)}
      </tbody>
    </table>
  </div>
</div>
</body>
</html>
"""
    return html

def render_html_to_image(html_content: str, output_image_path: str = OUTPUT_PNG):
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_content)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1120, "height": 900})
        page.goto(f"file:///{os.path.abspath(OUTPUT_HTML)}")
        page.wait_for_timeout(300)
        card_el = page.locator(".card-container").first
        if card_el.count() > 0:
            card_el.screenshot(path=output_image_path)
        else:
            page.screenshot(path=output_image_path, full_page=True)
        browser.close()
        
    return output_image_path

def upload_and_send_image_gtalk(image_path: str, caption: str, channel_id: str = None):
    config = load_config()
    cfg_gtalk = config.get("gtalk", {})
    oa_token = cfg_gtalk.get("token")
    target_channel = channel_id or cfg_gtalk.get("channel_id_group_a")
    
    print(f"📡 Đang upload ảnh {image_path} sang GTalk (Group {target_channel})...")
    try:
        img = Image.open(image_path)
        width, height = img.size
        file_size = os.path.getsize(image_path)
    except Exception as e:
        print(f"❌ Lỗi đọc ảnh: {e}")
        return False, str(e)

    # 1. Initiate Upload
    initiate_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
    payload_init = {
        "ChannelId": str(target_channel),
        "FileName": os.path.basename(image_path),
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": f'{{"width": {width}, "height": {height}}}',
        "oaToken": oa_token
    }
    headers = {"Content-Type": "application/json"}
    try:
        res_init = requests.post(initiate_url, json=payload_init, headers=headers, timeout=20, verify=False)
        if res_init.status_code != 200:
            return False, f"initiate-upload HTTP {res_init.status_code}: {res_init.text}"
        res_data = res_init.json()
        presigned_url = res_data["data"]["PresignedURL"]
        upload_id = res_data["data"]["UploadId"]

        # 2. Upload S3
        with open(image_path, "rb") as f:
            res_put = requests.put(presigned_url, data=f, headers={"Content-Type": "image/png"}, timeout=60, verify=False)
            if res_put.status_code != 200:
                return False, f"S3 PUT error HTTP {res_put.status_code}"

        # 3. Complete Upload
        complete_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
        res_comp = requests.post(complete_url, json={"oaToken": oa_token, "UploadId": upload_id}, headers=headers, timeout=20, verify=False)
        file_id = res_comp.json()["data"]["Id"]

        # 4. Send Message with Image Attachment
        send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
        client_msg_id = str(int(datetime.now().timestamp() * 1000))
        payload_send = {
            "channelId": str(target_channel),
            "clientMsgId": client_msg_id,
            "content": {
                "parseMode": "HTML",
                "attachment": {
                    "caption": caption,
                    "items": [{"image": {"fileId": file_id, "width": width, "height": height}}]
                }
            },
            "oaToken": oa_token
        }
        res_send = requests.post(send_url, json=payload_send, headers=headers, timeout=20, verify=False)
        if res_send.status_code == 200 and res_send.json().get("errorCode") == "success":
            return True, "OK"
        return False, f"send-message error: {res_send.text}"
    except Exception as e:
        return False, str(e)

def send_daily_attendance_recap_image(target_date: date = None, channel_id: str = None):
    print("1. Đang trích xuất dữ liệu điểm danh...")
    data = build_daily_recap_data(target_date)
    
    print("2. Đang render HTML Card...")
    html = generate_daily_recap_html(data)
    
    print("3. Đang chụp ảnh chất lượng cao...")
    img_path = render_html_to_image(html, OUTPUT_PNG)
    print(f"✅ Đã tạo ảnh: {img_path}")
    
    caption = f"📊 <b>BẢNG TỔNG HỢP ĐIỂM DANH & CHỐT PHẠT BÁO CÁO NGÀY {data['date_display']} - VÙNG NTB</b>\n💰 <b>Tổng tiền phạt:</b> <code>{data['total_fine']:,.0f} VNĐ</code> | <b>Số AM vi phạm:</b> <code>{data['violation_ams']}/{data['total_ams']}</code>\n🔗 Chi tiết: <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid=0'>Mở Google Sheet</a>"
    
    print("4. Đang gửi ảnh sang GTalk...")
    ok, err = upload_and_send_image_gtalk(img_path, caption, channel_id)
    if ok:
        print("🎉 GỬI ẢNH BẢNG TỔNG HỢP DAILY LÊN GTALK THÀNH CÔNG!")
    else:
        print(f"❌ Gửi ảnh thất bại: {err}")
    return ok, err

if __name__ == "__main__":
    send_daily_attendance_recap_image()
