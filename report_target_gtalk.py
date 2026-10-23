# -*- coding: utf-8 -*-
"""
Script: report_target_gtalk.py
Mô tả:
  1. Đọc dữ liệu từ Google Sheet (14cN4YAf01NqKtREtmhBngRP58OCdgby-PjVoixVlpZo).
  2. Lấy tất cả các bưu cục của danh sách SELECTED_AMS (hoặc tất cả bưu cục nếu SELECTED_AMS rỗng).
  3. Lấy Volume hôm nay (tab hàng về) và GTC hiện tại (tab GTC hiện tại).
  4. Linh hoạt xác định các mốc Target dựa trên % GTC hiện tại (pct_now):
     - Dưới 55%: [40%, 45%, 50%, 55%, 60%]
     - Đạt 55% - dưới 65%: Nâng lên [50%, 55%, 60%, 65%, 70%]
     - Đạt 65% - dưới 75%: Nâng lên [60%, 65%, 70%, 75%, 80%]
     - Đạt 75% - dưới 85%: Nâng lên [70%, 75%, 80%, 85%, 90%]
     - Trên 85%: Nâng lên [80%, 85%, 90%, 95%, 100%]
  5. Render hình ảnh bảng/card chi tiết theo tông màu CAM sang trọng.
  6. Tải ảnh lên S3 và gửi thông báo qua GTalk cho từng AM phụ trách.
"""

import sys, os
sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)

import json
import time
import requests
import urllib3
import pytz
from datetime import datetime
import pandas as pd
from PIL import Image
import gspread
from google.oauth2.credentials import Credentials as UserCredentials
from google.oauth2.service_account import Credentials as SACredentials
from playwright.sync_api import sync_playwright
import unicodedata

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============ CẤU HÌNH ============
SPREADSHEET_ID = "14cN4YAf01NqKtREtmhBngRP58OCdgby-PjVoixVlpZo"
GTALK_OA_TOKEN = "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"

# 🎯 DANH SÁCH AM CẦN GỬI BÁO CÁO
# - Muốn gửi cho ai thì điền tên người đó vào đây. Ví dụ: ["Trần Thị Nhung", "Phan Đình Duy"]
# - Nếu muốn GỬI TẤT CẢ AM như bình thường thì chỉ cần để rỗng: SELECTED_AMS = []
SELECTED_AMS = [
    "Nguyễn Thanh Long",
    "Trầm Hữu Tiến",
    "Huỳnh Thị Kim Chi",
    "Lê Văn Trường",
    "Lê Minh Lợi",
    "Trương Quang Linh",
    "Trần Thị Nhung",
    "Nguyễn Hoàng Phi",
    "Hồng Bích Nga",
    "Nguyễn Lê Nguyên Vũ",
    "Phan Đình Duy",
    "Nguyễn Đỗ Minh Nghĩa"
]

# Mapping AM sang Chat/Channel ID
AM_CHANNEL_MAP = {
    "Nguyễn Ngọc Khánh": "2077277510775197696",
    "Nguyễn Duy Long": "2077277718988836864",
    "Lê Thanh Nhựt": "2077277754418147328",
    "Trần Văn Phước": "2077277797832024064",
    "Trần Thị Nhung": "2077277827057745920",
    "Huỳnh Thị Kim Chi": "2077277857186131968",
    "Phan Đình Duy": "2077278383281876992",
    "Phạm Bá Thành Công": "2077277907735883776",
    "Thái Thị Thanh Thư": "2077277934947827712",
    "Nguyễn Thanh Long": "2077277974325506048",
    "Nguyễn Hoàng Phi": "2077278021170323456",
    "Trầm Hữu Tiến": "2077278046487142400",
    "Nguyễn Lê Nguyên Vũ": "2077278095459835904",
    "Lê Văn Trường": "2077278127814696960",
    "Hồng Bích Nga": "2077278157729837056",
    "Lê Minh Đại": "2077278182818799616",
    "Phan Thị Ngọc Diễm": "2079827073949868032",
    "Lê Hồng Minh Tâm": "2079827054540226560",
    "Cao Thị Thanh Thủy": "2083241927281995776",
    "Huỳnh Thúc Duân": "2089391843188129792",
    "Lê Hồng Phong": "2089391843188129792",
    "Nguyễn Thị Tuyết Thơ": "2089391817020141568",
    "Trương Quang Linh": "2094079475020627968",
    "Lê Minh Lợi": "2094079507615027200",
    "Nguyễn Đỗ Minh Nghĩa": "2100062122691026944",
    "Phan Nguyễn Yến Nhi": "2105595062412402688",
}
DEFAULT_CHANNEL_ID = "2077277510775197696"

def get_channel_id(am_name):
    if not am_name:
        return DEFAULT_CHANNEL_ID
    norm_am = unicodedata.normalize('NFC', str(am_name)).strip().lower()
    # 1. So khớp chính xác sau khi chuẩn hóa Unicode NFC
    for k, v in AM_CHANNEL_MAP.items():
        if unicodedata.normalize('NFC', k).strip().lower() == norm_am:
            return v
    # 2. So khớp chuỗi con (phòng khi có khoảng trắng dư hoặc viết tắt)
    for k, v in AM_CHANNEL_MAP.items():
        k_norm = unicodedata.normalize('NFC', k).strip().lower()
        if k_norm in norm_am or norm_am in k_norm:
            return v
    print(f"⚠️ [CẢNH BÁO] Không tìm thấy Channel ID cho AM '{am_name}', dùng default: {DEFAULT_CHANNEL_ID}", flush=True)
    return DEFAULT_CHANNEL_ID

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def get_gspread_client():
    auth_user_candidates = [
        os.path.join(SCRIPT_DIR, 'authorized_user.json'),
        r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
        'authorized_user.json'
    ]
    for f in auth_user_candidates:
        if os.path.exists(f):
            try:
                creds = UserCredentials.from_authorized_user_file(f, scopes=SCOPES)
                return gspread.authorize(creds)
            except Exception:
                pass
    raise PermissionError("Không thể xác thực Google Sheets API.")

def get_dynamic_milestones(pct_now):
    """
    Tự động nâng mốc mục tiêu Target dựa trên % GTC hiện tại của bưu cục:
    - Nếu pct_now < 55%: Dùng bộ mốc chuẩn [40, 45, 50, 55, 60]
    - Nếu pct_now đạt 55% - 65%: Nâng lên [50, 55, 60, 65, 70]
    - Nếu pct_now đạt 65% - 75%: Nâng lên [60, 65, 70, 75, 80]
    - Nếu pct_now đạt 75% - 85%: Nâng lên [70, 75, 80, 85, 90]
    - Nếu pct_now >= 85%: Nâng lên [80, 85, 90, 95, 100]
    Đảm bảo luôn hiển thị 5 mốc: vừa có mốc đã hoàn thành, vừa có mốc tiếp theo cần đẩy.
    """
    if pct_now < 55.0:
        return [40, 45, 50, 55, 60]
    elif pct_now < 65.0:
        return [50, 55, 60, 65, 70]
    elif pct_now < 75.0:
        return [60, 65, 70, 75, 80]
    elif pct_now < 85.0:
        return [70, 75, 80, 85, 90]
    else:
        return [80, 85, 90, 95, 100]

def load_hub_data():
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    # 1. Map AM & Tỉnh từ tab cocau
    ws_cocau = sh.worksheet('cocau')
    cocau_rows = ws_cocau.get_all_values()
    hub_info = {}
    for r in cocau_rows[1:]:
        if len(r) >= 4:
            bc_id = r[0].strip()
            bc_name = unicodedata.normalize('NFC', r[1]).strip()
            tinh = unicodedata.normalize('NFC', r[2]).strip()
            am = unicodedata.normalize('NFC', r[3]).strip()
            if bc_name:
                hub_info[bc_name] = {'id': bc_id, 'tinh': tinh, 'am': am}

    # 2. Đọc % GTC ngày N-1 từ tab Data NTB
    ws_data = sh.worksheet('Data NTB')
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    df_data['Volume_num'] = pd.to_numeric(df_data['Volume'].str.replace(',', '').str.replace('.', ''), errors='coerce').fillna(0)
    df_data['GTC_num'] = pd.to_numeric(df_data['Sản Lượng Giao Thành Công'].str.replace(',', '').str.replace('.', ''), errors='coerce').fillna(0)
    
    sorted_dates = sorted([d for d in df_data['Time'].unique() if d], reverse=True)
    latest_date_n1 = sorted_dates[0] if sorted_dates else "N-1"
    df_n1 = df_data[df_data['Time'] == latest_date_n1]
    agg_n1 = df_n1.groupby('Chi tiết')[['Volume_num', 'GTC_num']].sum().reset_index()
    agg_n1['pct_gtc_n1'] = (agg_n1['GTC_num'] / agg_n1['Volume_num'] * 100).round(2)
    
    gtc_n1_map = {}
    for _, row in agg_n1.iterrows():
        gtc_n1_map[row['Chi tiết']] = {
            'vol_n1': int(row['Volume_num']),
            'gtc_n1': int(row['GTC_num']),
            'pct_gtc_n1': row['pct_gtc_n1']
        }

    # 3. Lấy Volume hôm nay từ tab hàng về
    ws_hv = sh.worksheet('hàng về')
    hv_rows = ws_hv.get_all_values()
    vol_today_map = {}
    for r in hv_rows[1:]:
        if len(r) >= 2:
            bc_id = r[0].strip()
            vol_val = r[1].strip().replace(',', '').replace('.', '')
            if vol_val.isdigit():
                vol_today_map[bc_id] = int(vol_val)

    # 4. Lấy GTC hiện tại từ tab GTC hiện tại
    ws_gtc = sh.worksheet('GTC hiện tại')
    gtc_rows = ws_gtc.get_all_values()
    gtc_now_map = {}
    for r in gtc_rows[1:]:
        if len(r) >= 2:
            bc_name = unicodedata.normalize('NFC', r[0]).strip()
            gtc_val = r[1].strip().replace(',', '').replace('.', '')
            if gtc_val.isdigit():
                gtc_now_map[bc_name] = int(gtc_val)

    # 5. Lấy danh sách bưu cục theo AM và tính toán target linh hoạt
    affected_hubs = []
    for bc_name, info in hub_info.items():
        hub_am = info['am']

        # Lọc theo danh sách SELECTED_AMS nếu có chỉ định
        if SELECTED_AMS:
            norm_hub_am = unicodedata.normalize('NFC', str(hub_am)).strip().lower()
            is_matched = any(
                unicodedata.normalize('NFC', str(target)).strip().lower() in norm_hub_am or
                norm_hub_am in unicodedata.normalize('NFC', str(target)).strip().lower()
                for target in SELECTED_AMS
            )
            if not is_matched:
                continue

        n1_data = gtc_n1_map.get(bc_name)
        if not n1_data:
            for k, v in gtc_n1_map.items():
                if k in bc_name or bc_name in k:
                    n1_data = v
                    break
        
        pct_gtc_n1 = n1_data['pct_gtc_n1'] if n1_data else 0.0

        bc_id = info['id']
        vol_today = vol_today_map.get(bc_id, 0)
        gtc_now = gtc_now_map.get(bc_name, 0)
        
        # Fallback đối soát tên phòng khi tab GTC viết tắt hoặc không khớp chính xác 100%
        if gtc_now == 0:
            norm_bc = unicodedata.normalize('NFC', bc_name).strip().lower()
            for k_g, v_g in gtc_now_map.items():
                norm_k = unicodedata.normalize('NFC', k_g).strip().lower()
                if norm_k in norm_bc or norm_bc in norm_k:
                    gtc_now = v_g
                    break

        # Fallback an toàn: nếu tab hàng về chưa có dữ liệu (vol_today < gtc_now)
        # thì tối thiểu vol_today phải bằng gtc_now để số liệu hợp lý
        if vol_today < gtc_now:
            vol_today = gtc_now

        pct_now = round(gtc_now / vol_today * 100, 2) if vol_today > 0 else 0.0

        # Xác định 5 mốc target linh hoạt phù hợp với tiến độ hiện tại của bưu cục
        milestones = get_dynamic_milestones(pct_now)

        targets = {}
        for t_pct in milestones:
            t_vol = int(round(vol_today * (t_pct / 100.0)))
            gap = t_vol - gtc_now
            targets[t_pct] = {
                'target_vol': t_vol,
                'gap': gap,
                'status': 'ACHIEVED' if gap <= 0 else 'MISSING'
            }

        affected_hubs.append({
            'bc_name': bc_name,
            'bc_id': bc_id,
            'tinh': info['tinh'],
            'am': hub_am,
            'date_n1': latest_date_n1,
            'pct_gtc_n1': pct_gtc_n1,
            'vol_today': vol_today,
            'gtc_now': gtc_now,
            'pct_now': pct_now,
            'milestones': milestones,
            'targets': targets
        })

    affected_hubs.sort(key=lambda x: (x['am'], x['pct_gtc_n1']))
    return affected_hubs, latest_date_n1

def render_hub_card_html(hub, now_str):
    targets = hub['targets']
    milestones = hub.get('milestones', list(targets.keys()))
    
    rows_html = ""
    for pct in milestones:
        t_data = targets[pct]
        t_vol = t_data['target_vol']
        gap = t_data['gap']
        
        if gap <= 0:
            gap_badge = f'<span class="badge-success">✅ Đã Đạt (Vượt {abs(gap):,} đơn)</span>'
            row_class = "row-success"
        else:
            gap_badge = f'<span class="badge-danger">🔴 THIẾU {gap:,} ĐƠN</span>'
            row_class = "row-missing"

        rows_html += f"""
        <tr class="{row_class}">
          <td class="col-target"><span class="target-chip">{pct}%</span></td>
          <td class="col-num"><b>{t_vol:,}</b> đơn</td>
          <td class="col-num">{hub['gtc_now']:,} đơn</td>
          <td class="col-gap">{gap_badge}</td>
        </tr>
        """

    # Ghi chú trạng thái nâng target nếu có
    target_note = ""
    if min(milestones) > 40:
        target_note = f'<span class="badge-elevated">🔥 Đã nâng target lên mốc {max(milestones)}%</span>'

    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; }}
    body {{ background: #1e1e24; padding: 24px; display: flex; justify-content: center; align-items: center; }}
    .card {{
      width: 760px;
      background: #ffffff;
      border-radius: 20px;
      overflow: hidden;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
      border: 1px solid #fed7aa;
    }}
    .header {{
      background: linear-gradient(135deg, #c2410c 0%, #ea580c 45%, #f97316 100%);
      color: white;
      padding: 24px 28px;
      position: relative;
    }}
    .header-top {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }}
    .header-tags {{
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    .badge-hub {{
      background: rgba(255, 255, 255, 0.22);
      backdrop-filter: blur(8px);
      padding: 4px 14px;
      border-radius: 20px;
      font-size: 13px;
      font-weight: 700;
      letter-spacing: 0.5px;
      border: 1px solid rgba(255,255,255,0.3);
    }}
    .badge-elevated {{
      background: #fef08a;
      color: #854d0e;
      padding: 4px 12px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 800;
      box-shadow: 0 2px 4px rgba(0,0,0,0.15);
    }}
    .time-badge {{
      font-size: 12px;
      opacity: 0.95;
      background: rgba(0, 0, 0, 0.25);
      padding: 4px 10px;
      border-radius: 12px;
    }}
    .hub-title {{
      font-size: 27px;
      font-weight: 800;
      letter-spacing: -0.5px;
      margin-bottom: 5px;
    }}
    .hub-subtitle {{
      font-size: 14px;
      opacity: 0.95;
      display: flex;
      gap: 16px;
    }}
    .metrics-grid {{
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 12px;
      padding: 20px 28px;
      background: #fffaf5;
      border-bottom: 1px solid #ffedd5;
    }}
    .metric-box {{
      background: white;
      padding: 14px;
      border-radius: 12px;
      border: 1px solid #fed7aa;
      text-align: center;
      box-shadow: 0 2px 4px rgba(234, 88, 12, 0.04);
    }}
    .metric-box.danger {{
      background: #fef2f2;
      border-color: #fecaca;
    }}
    .metric-box.primary {{
      background: #fff7ed;
      border-color: #fdba74;
    }}
    .metric-label {{
      font-size: 11px;
      font-weight: 700;
      color: #7c2d12;
      text-transform: uppercase;
      margin-bottom: 4px;
    }}
    .metric-value {{
      font-size: 20px;
      font-weight: 800;
      color: #1e293b;
    }}
    .metric-value.danger {{ color: #dc2626; }}
    .metric-value.primary {{ color: #ea580c; }}

    .table-container {{
      padding: 20px 28px 24px;
    }}
    .table-header-title {{
      font-size: 14px;
      font-weight: 800;
      color: #9a3412;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin-bottom: 12px;
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    table {{
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      border-radius: 12px;
      overflow: hidden;
      border: 1px solid #fed7aa;
    }}
    th {{
      background: #7c2d12;
      color: white;
      font-size: 12px;
      font-weight: 700;
      text-transform: uppercase;
      padding: 12px 14px;
      text-align: center;
    }}
    td {{
      padding: 14px 16px;
      text-align: center;
      font-size: 14px;
      border-bottom: 1px solid #ffedd5;
    }}
    tr:last-child td {{ border-bottom: none; }}
    .col-target {{ font-weight: 800; }}
    .target-chip {{
      background: #ea580c;
      color: white;
      padding: 4px 12px;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 800;
    }}
    .col-num {{ font-size: 15px; color: #1e293b; }}
    .col-gap {{ text-align: center; }}
    .badge-danger {{
      background: #fee2e2;
      color: #b91c1c;
      padding: 6px 14px;
      border-radius: 20px;
      font-weight: 800;
      font-size: 13px;
      display: inline-block;
      border: 1px solid #fca5a5;
    }}
    .badge-success {{
      background: #dcfce7;
      color: #15803d;
      padding: 6px 14px;
      border-radius: 20px;
      font-weight: 800;
      font-size: 13px;
      display: inline-block;
      border: 1px solid #86efac;
    }}
    .footer {{
      background: #fff7ed;
      padding: 12px 28px;
      font-size: 11px;
      color: #9a3412;
      display: flex;
      justify-content: space-between;
      border-top: 1px solid #fed7aa;
    }}
  </style>
</head>
<body>
  <div class="card">
    <div class="header">
      <div class="header-top">
        <div class="header-tags">
          <span class="badge-hub">THEO DÕI TARGET GTC</span>
          {target_note}
        </div>
        <span class="time-badge">Cập nhật: {now_str}</span>
      </div>
      <div class="hub-title">{hub['bc_name']}</div>
      <div class="hub-subtitle">
        <span>📍 <b>Tỉnh:</b> {hub['tinh']}</span>
        <span>👤 <b>AM:</b> {hub['am']}</span>
      </div>
    </div>

    <div class="metrics-grid">
      <div class="metric-box {'danger' if hub['pct_gtc_n1'] < 55.0 else ''}">
        <div class="metric-label">% GTC Hôm Qua (N-1)</div>
        <div class="metric-value {'danger' if hub['pct_gtc_n1'] < 55.0 else ''}">{hub['pct_gtc_n1']}%</div>
      </div>
      <div class="metric-box primary">
        <div class="metric-label">Tổng Hàng Hôm Nay</div>
        <div class="metric-value primary">{hub['vol_today']:,}</div>
      </div>
      <div class="metric-box">
        <div class="metric-label">Đã Giao Hiện Tại</div>
        <div class="metric-value">{hub['gtc_now']:,}</div>
      </div>
      <div class="metric-box">
        <div class="metric-label">% GTC Hiện Tại</div>
        <div class="metric-value">{hub['pct_now']}%</div>
      </div>
    </div>

    <div class="table-container">
      <div class="table-header-title">
        🎯 MỤC TIÊU SỐ ĐƠN CẦN ĐẠT & CÒN THIẾU HÔM NAY
      </div>
      <table>
        <thead>
          <tr>
            <th>Mốc % Target</th>
            <th>Target Tổng Cần Đạt</th>
            <th>Đã Giao Thực Tế</th>
            <th>SỐ ĐƠN CÒN THIẾU CẦN ĐẨY NGAY</th>
          </tr>
        </thead>
        <tbody>
          {rows_html}
        </tbody>
      </table>
    </div>

    <div class="footer">
      <span>Auto Report System · Vận Hành NTB</span>
      <span>Công thức: Số đơn còn thiếu = Target Tổng - Đã giao hiện tại</span>
    </div>
  </div>
</body>
</html>"""
    return html

def render_image(html_content, image_path):
    temp_html = image_path.replace(".png", ".html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 820, "height": 700})
        page.goto(f"file:///{os.path.abspath(temp_html)}")
        page.wait_for_timeout(300)
        card_el = page.locator(".card").first
        if card_el.count() > 0:
            card_el.screenshot(path=image_path)
        else:
            page.screenshot(path=image_path, full_page=True)
        browser.close()

    try:
        if os.path.exists(temp_html):
            os.remove(temp_html)
    except Exception:
        pass

def send_gtalk_photo(image_path, caption, channel_id):
    oa_token = GTALK_OA_TOKEN
    img = Image.open(image_path)
    width, height = img.size
    file_size = os.path.getsize(image_path)

    # 1. Initiate Upload
    initiate_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
    payload_init = {
        "ChannelId": channel_id,
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
            print(f"❌ Initiate error {res_init.status_code}: {res_init.text}")
            return False
        res_data = res_init.json()
        if res_data.get("errorCode") != "success":
            print(f"❌ Initiate API error: {res_data.get('error')}")
            return False
        presigned_url = res_data["data"]["PresignedURL"]
        upload_id = res_data["data"]["UploadId"]

        # 2. Upload S3
        with open(image_path, "rb") as f:
            res_put = requests.put(presigned_url, data=f, headers={"Content-Type": "image/png"}, timeout=60, verify=False)
            if res_put.status_code != 200:
                print(f"❌ S3 PUT error {res_put.status_code}")
                return False

        # 3. Complete Upload
        complete_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
        res_comp = requests.post(complete_url, json={"oaToken": oa_token, "UploadId": upload_id}, headers=headers, timeout=20, verify=False)
        file_id = res_comp.json()["data"]["Id"]

        # 4. Send Message
        send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
        client_msg_id = str(int(time.time() * 1000))
        payload_send = {
            "channelId": channel_id,
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
            print(f"🎉 Đã gửi thành công ảnh bưu cục sang GTalk (Channel: {channel_id})")
            return True
        else:
            print(f"❌ Send message error: {res_send.text}")
    except Exception as e:
        print(f"❌ GTalk error: {e}")
    return False

def format_gap_text(gap_val):
    if gap_val <= 0:
        return "✅ Đã đạt"
    return f"🔴 <b>Thiếu {gap_val:,} đơn</b>"

def main():
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    now_str = datetime.now(tz).strftime("%d/%m/%Y %H:%M")
    
    if SELECTED_AMS:
        print(f"[{now_str}] 🚀 Đang lấy tất cả bưu cục của {len(SELECTED_AMS)} AM được chọn: {', '.join(SELECTED_AMS)}...")
    else:
        print(f"[{now_str}] 🚀 Đang lấy tất cả bưu cục của toàn bộ AM...")

    hubs, date_n1 = load_hub_data()
    print(f"✅ Tìm thấy {len(hubs)} bưu cục cần gửi báo cáo (Ngày N-1: {date_n1}).")

    img_dir = os.path.join(SCRIPT_DIR, "output_target_cards")
    os.makedirs(img_dir, exist_ok=True)

    success_count = 0

    for idx, hub in enumerate(hubs):
        bc_clean_name = hub['bc_name'].replace("(", "").replace(")", "").replace(" ", "_")
        img_path = os.path.join(img_dir, f"target_{bc_clean_name}.png")

        milestones_txt = ", ".join([f"{p}%" for p in hub['milestones']])
        print(f"\n[{idx+1}/{len(hubs)}] 🎨 Đang tạo ảnh màu CAM cho: {hub['bc_name']} (AM: {hub['am']} | GTC N-1: {hub['pct_gtc_n1']}% | Hiện tại: {hub['pct_now']}% | Mốc: {milestones_txt})")
        html = render_hub_card_html(hub, now_str)
        render_image(html, img_path)

        # Lấy channel ID của AM (đã chuẩn hóa Unicode)
        channel_id = get_channel_id(hub['am'])
        
        # Caption chi tiết linh hoạt theo các mốc target
        milestone_lines = []
        for pct in hub['milestones']:
            gap_txt = format_gap_text(hub['targets'][pct]['gap'])
            milestone_lines.append(f"• Mốc {pct}%: {gap_txt}")
        milestone_str = "\n".join(milestone_lines)

        caption = (
            f"🎯 <b>TARGET GTC - {hub['bc_name']}</b>\n"
            f"👤 <b>AM Phụ Trách:</b> {hub['am']} | 📍 <b>Tỉnh:</b> {hub['tinh']}\n"
            f"📉 <b>% GTC Ngày N-1 ({date_n1}):</b> <code>{hub['pct_gtc_n1']}%</code>\n\n"
            f"📦 <b>TIẾN ĐỘ HÔM NAY:</b>\n"
            f"• Tổng hàng cần giao: <b>{hub['vol_today']:,} đơn</b>\n"
            f"• Đã giao thực tế: <b>{hub['gtc_now']:,} đơn ({hub['pct_now']}%)</b>\n\n"
            f"🚀 <b>SỐ ĐƠN CÒN THIẾU ĐỂ ĐẠT MỤC TIÊU:</b>\n"
            f"{milestone_str}\n"
        )

        print(f"📤 Đang gửi tin nhắn GTalk cho AM {hub['am']}...")
        if send_gtalk_photo(img_path, caption, channel_id):
            success_count += 1
        time.sleep(1)

    print(f"\n🏁 HOÀN TẤT: Đã gửi {success_count}/{len(hubs)} bưu cục sang GTalk thành công!")

if __name__ == "__main__":
    main()
