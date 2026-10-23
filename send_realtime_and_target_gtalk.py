# -*- coding: utf-8 -*-
"""
Script: send_realtime_and_target_gtalk.py
Mô tả:
  Kết hợp BÁO CÁO NĂNG SUẤT NVPTT REAL-TIME + THEO DÕI TARGET % GTC BƯU CỤC trong 1 ảnh duy nhất:
  1. Đọc dữ liệu Năng suất NVPTT từ sheet: 1-p9VUXndK_7BoiT-a81UfTCbUi953XNmVBoXaTGis_c (tab BaoCao).
  2. Đọc dữ liệu Target GTC, Volume hôm nay, GTC hôm qua từ sheet: 14cN4YAf01NqKtREtmhBngRP58OCdgby-PjVoixVlpZo.
  3. Tự động tính toán 5 mốc Target linh hoạt và số đơn còn thiếu cho từng mốc.
  4. Render 1 ảnh báo cáo tổng hợp tuyệt đẹp (Theme Vàng Kim / Amber sang trọng).
  5. Tạo caption GTalk đầy đủ thông tin: Tiến độ Target bưu cục + Năng suất nhân sự.
  6. Hỗ trợ tùy chọn --no-send để test và xem trước hình ảnh mà KHÔNG gửi thật.
"""

import os
import re
import sys
import json
import time
import argparse
import unicodedata
from datetime import datetime

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import urllib3
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as UserCredentials
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ===== CẤU HÌNH GOOGLE SPREADSHEETS =====
BAOCAO_SHEET_KEY = "1-p9VUXndK_7BoiT-a81UfTCbUi953XNmVBoXaTGis_c"
BAOCAO_TAB_NAME = "BaoCao"

TARGET_SHEET_KEY = "14cN4YAf01NqKtREtmhBngRP58OCdgby-PjVoixVlpZo"

# ===== CẤU HÌNH GTALK =====
GTALK_OA_TOKEN = os.environ.get("NVPTT_GTALK_OA_TOKEN") or "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
DEFAULT_TEST_CHANNEL = "2077277510775197696"

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

def get_channel_for_am(am_name):
    if not am_name:
        return DEFAULT_TEST_CHANNEL
    norm_target = unicodedata.normalize('NFC', am_name.strip()).lower()
    for name, channel_id in AM_CHANNEL_MAP.items():
        if unicodedata.normalize('NFC', name.strip()).lower() == norm_target:
            return channel_id
    for name, channel_id in AM_CHANNEL_MAP.items():
        norm_k = unicodedata.normalize('NFC', name.strip()).lower()
        if norm_k in norm_target or norm_target in norm_k:
            return channel_id
    return DEFAULT_TEST_CHANNEL

def get_http_session():
    session = requests.Session()
    retries = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[500, 502, 503, 504, 429],
        raise_on_status=False
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

def parse_pct(val):
    if val is None:
        return 0.0
    s = str(val).strip().replace("%", "").replace(",", ".")
    if s == "":
        return 0.0
    try:
        f = float(s)
    except ValueError:
        return 0.0
    return f * 100 if 0 <= f <= 1.5 else f

def parse_num(val):
    if val is None or str(val).strip() == "":
        return 0
    s = str(val).strip().replace(".", "").replace(",", "")
    try:
        return int(float(s))
    except ValueError:
        return 0

SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, 'credentials.json'),
    r'C:\Users\lap4all\Documents\Auto report\credentials.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json',
    'credentials.json'
]

def get_gspread_client(spreadsheet_id=None):
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = SACredentials.from_service_account_file(cred_path, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception:
                pass

    auth_user_candidates = [
        os.path.join(BASE_DIR, 'authorized_user.json'),
        r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
        'authorized_user.json'
    ]
    for auth_user_file in auth_user_candidates:
        if os.path.exists(auth_user_file):
            try:
                creds = UserCredentials.from_authorized_user_file(auth_user_file, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception:
                pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")

def get_dynamic_milestones(pct_now):
    """
    Xác định 5 mốc target linh hoạt theo tiến độ hiện tại của bưu cục:
    - < 55%: [40%, 45%, 50%, 55%, 60%]
    - 55% - 65%: [50%, 55%, 60%, 65%, 70%]
    - 65% - 75%: [60%, 65%, 70%, 75%, 80%]
    - 75% - 85%: [70%, 75%, 80%, 85%, 90%]
    - >= 85%: [80%, 85%, 90%, 95%, 100%]
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

# ===== 1. ĐỌC DỮ LIỆU TARGET TỪ SHEET TARGET =====
def load_target_data():
    """Đọc dữ liệu từ sheet 14cN4YAf01NqKtREtmhBngRP58OCdgby-PjVoixVlpZo."""
    print("⏳ Đang tải dữ liệu từ Google Sheet Target...", flush=True)
    gc = get_gspread_client(TARGET_SHEET_KEY)
    sh = gc.open_by_key(TARGET_SHEET_KEY)

    # 1. Cơ cấu bưu cục (id, tên, tỉnh, am)
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

    # 2. % GTC ngày N-1 từ tab Data NTB (Đúng chuẩn Ca 1 + Hàng Tồn)
    ws_data = sh.worksheet('Data NTB')
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    df_data['Volume_num'] = pd.to_numeric(df_data['Volume'].astype(str).str.replace(',', '').str.replace('.', ''), errors='coerce').fillna(0)
    df_data['GTC_num'] = pd.to_numeric(df_data['Sản Lượng Giao Thành Công'].astype(str).str.replace(',', '').str.replace('.', ''), errors='coerce').fillna(0)

    sorted_dates = sorted([d for d in df_data['Time'].unique() if d], reverse=True)
    latest_date_n1 = sorted_dates[0] if sorted_dates else "N-1"
    df_n1 = df_data[df_data['Time'] == latest_date_n1]
    agg_n1 = df_n1.groupby('Chi tiết')[['Volume_num', 'GTC_num']].sum().reset_index()
    agg_n1['pct_gtc_n1'] = (agg_n1['GTC_num'] / agg_n1['Volume_num'] * 100).round(2)

    gtc_n1_map = {}
    for _, row in agg_n1.iterrows():
        gtc_n1_map[unicodedata.normalize('NFC', str(row['Chi tiết'])).strip()] = {
            'vol_n1': int(row['Volume_num']),
            'gtc_n1': int(row['GTC_num']),
            'pct_gtc_n1': float(row['pct_gtc_n1'])
        }

    # 3. Volume hôm nay từ tab hàng về
    ws_hv = sh.worksheet('hàng về')
    hv_rows = ws_hv.get_all_values()
    vol_today_map = {}
    for r in hv_rows[1:]:
        if len(r) >= 2:
            bc_id = r[0].strip()
            vol_val = r[1].strip().replace(',', '').replace('.', '')
            if vol_val.isdigit():
                vol_today_map[bc_id] = int(vol_val)

    # 4. GTC hiện tại từ tab GTC hiện tại
    ws_gtc = sh.worksheet('GTC hiện tại')
    gtc_rows = ws_gtc.get_all_values()
    gtc_now_map = {}
    for r in gtc_rows[1:]:
        if len(r) >= 2:
            bc_name = unicodedata.normalize('NFC', r[0]).strip()
            gtc_val = r[1].strip().replace(',', '').replace('.', '')
            if gtc_val.isdigit():
                gtc_now_map[bc_name] = int(gtc_val)

    # Tạo map tổng hợp
    target_data_map = {}
    for bc_name, info in hub_info.items():
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
        if gtc_now == 0:
            for k_g, v_g in gtc_now_map.items():
                if k_g in bc_name or bc_name in k_g:
                    gtc_now = v_g
                    break

        if vol_today < gtc_now:
            vol_today = gtc_now

        target_data_map[bc_name] = {
            'bc_name': bc_name,
            'bc_id': bc_id,
            'tinh': info['tinh'],
            'am': info['am'],
            'date_n1': latest_date_n1,
            'pct_gtc_n1': pct_gtc_n1,
            'vol_today': vol_today,
            'gtc_now': gtc_now
        }

    print(f"✅ Đã tải thông tin Target của {len(target_data_map)} bưu cục (Ngày N-1: {latest_date_n1}).", flush=True)
    return target_data_map, latest_date_n1

# ===== 2. ĐỌC DỮ LIỆU NĂNG SUẤT NVPTT REAL-TIME =====
def read_baocao_realtime(sheet_key=BAOCAO_SHEET_KEY, tab_name=BAOCAO_TAB_NAME):
    print("⏳ Đang tải dữ liệu Năng suất NVPTT từ Google Sheet...", flush=True)
    gc = get_gspread_client(sheet_key)
    sh = gc.open_by_key(sheet_key)
    ws = sh.worksheet(tab_name)

    all_values = ws.get_all_values()

    def norm(s):
        return unicodedata.normalize('NFC', str(s)).strip().lower()

    header_idx = None
    for idx, row in enumerate(all_values):
        row_join = " ".join(norm(c) for c in row)
        if ("nhan vien" in row_join.replace("â", "a") or "nhân viên" in row_join) and ("bưu" in row_join or "buu" in row_join):
            header_idx = idx
            break

    if header_idx is None:
        raise ValueError("Không tìm thấy dòng header trong tab BaoCao.")

    headers = [h.strip() for h in all_values[header_idx]]
    data_rows = all_values[header_idx + 1:]

    def col_idx(*candidates):
        for cand in candidates:
            for i, h in enumerate(headers):
                if norm(h).replace(" ", "") == norm(cand).replace(" ", ""):
                    return i
        return None

    i_bc = col_idx("Bưu Cục", "BuuCuc")
    i_am = col_idx("AM")
    i_manv = col_idx("Mã NV", "MaNV")
    i_name = col_idx("Nhân Viên", "NhanVien")
    i_gan = col_idx("Gán Giao", "GanGiao", "Tổng đơn gán giao")
    i_tc = col_idx("Giao TC", "GiaoTC", "Số đơn GTC")
    i_pct = col_idx("%GTC")
    i_ltc = col_idx("LTC")
    i_danhgia = col_idx("Đánh Giá", "DanhGia")

    missing = [name for name, i in [("Bưu Cục", i_bc), ("AM", i_am), ("Nhân Viên", i_name), ("%GTC", i_pct)] if i is None]
    if missing:
        raise ValueError(f"Thiếu cột bắt buộc trong tab BaoCao: {missing}.")

    grouped = {}
    for row in data_rows:
        if not any(c.strip() for c in row):
            continue

        bc = row[i_bc].strip() if i_bc < len(row) else ""
        am = row[i_am].strip() if i_am < len(row) else ""
        name = row[i_name].strip() if i_name < len(row) else ""
        if not bc or not am or not name:
            continue

        ma_nv = row[i_manv].strip() if (i_manv is not None and i_manv < len(row)) else ""
        gan = parse_num(row[i_gan]) if (i_gan is not None and i_gan < len(row)) else 0
        tc = parse_num(row[i_tc]) if (i_tc is not None and i_tc < len(row)) else 0
        pct = parse_pct(row[i_pct]) if i_pct < len(row) else 0.0
        ltc = parse_num(row[i_ltc]) if (i_ltc is not None and i_ltc < len(row)) else 0
        danh_gia = row[i_danhgia].strip() if (i_danhgia is not None and i_danhgia < len(row)) else ""

        # Chống trùng lặp nhân viên trong cùng 1 bưu cục: Nếu đã có, giữ lại bản ghi có đơn giao TC cao nhất (mới nhất)
        existing_list = grouped.setdefault(am, {}).setdefault(bc, [])
        dup_idx = None
        for i, item in enumerate(existing_list):
            if (ma_nv and item["ma_nv"] == ma_nv) or (not ma_nv and item["name"] == name):
                dup_idx = i
                break

        staff_data = {
            "ma_nv": ma_nv,
            "name": name,
            "gan": gan,
            "tc": tc,
            "pct": pct,
            "ltc": ltc,
            "danh_gia": danh_gia
        }

        if dup_idx is not None:
            if tc > existing_list[dup_idx]["tc"] or (tc == existing_list[dup_idx]["tc"] and gan >= existing_list[dup_idx]["gan"]):
                existing_list[dup_idx] = staff_data
        else:
            existing_list.append(staff_data)

    result = []
    for am_name, bc_map in grouped.items():
        bcs = [(bc_name, staff_list) for bc_name, staff_list in bc_map.items()]
        result.append({"am": am_name, "bcs": bcs})

    total_staff = sum(len(s) for am in result for _, s in am["bcs"])
    print(f"📊 Đã tải thành công {total_staff} NVPTT tại {sum(len(am['bcs']) for am in result)} bưu cục thuộc {len(result)} AM.", flush=True)
    return result

# ===== 3. RENDER HTML TỔNG HỢP: NĂNG SUẤT NVPTT + TARGET GTC =====
def build_combined_report_html(am_name, bc_name, staff, target_info, report_date_str, update_time_str):
    total_gan = sum(s["gan"] for s in staff)
    total_tc = sum(s["tc"] for s in staff)
    total_ltc = sum(s["ltc"] for s in staff)
    total_rows = len(staff)

    # Lấy thông tin Target GTC
    tinh = target_info.get('tinh', '')
    date_n1 = target_info.get('date_n1', 'N-1')
    pct_gtc_n1 = target_info.get('pct_gtc_n1', 0.0)
    vol_today = target_info.get('vol_today', total_gan)
    if vol_today < total_tc:
        vol_today = total_tc
    gtc_now = target_info.get('gtc_now', total_tc)

    pct_now = round(gtc_now / vol_today * 100, 2) if vol_today > 0 else 0.0
    milestones = get_dynamic_milestones(pct_now)

    # Tính toán các mốc target
    targets = {}
    for t_pct in milestones:
        t_vol = int(round(vol_today * (t_pct / 100.0)))
        gap = t_vol - gtc_now
        targets[t_pct] = {
            'target_vol': t_vol,
            'gap': gap,
            'achieved': gap <= 0
        }

    tot_ok = sum(1 for s in staff if "ok" in str(s["danh_gia"]).lower())
    tot_dat = sum(1 for s in staff if "đạt" in str(s["danh_gia"]).lower() or "dat" in str(s["danh_gia"]).lower())
    tot_thap = sum(1 for s in staff if "thấp" in str(s["danh_gia"]).lower() or "thap" in str(s["danh_gia"]).lower())

    eval_summary = f"{tot_ok} OK"
    if tot_dat > 0:
        eval_summary += f" · {tot_dat} Đạt"
    if tot_thap > 0:
        eval_summary += f" · {tot_thap} Thấp"

    # HTML 5 thẻ Target nằm ngang
    target_cards_html = ""
    for pct in milestones:
        t_data = targets[pct]
        t_vol = t_data['target_vol']
        gap = t_data['gap']
        achieved = t_data['achieved']

        progress_pct = min(100.0, round((gtc_now / t_vol * 100), 1)) if t_vol > 0 else 100.0

        if achieved:
            card_class = "t-card-achieved"
            status_badge = f'<span class="t-badge t-badge-achieved">✓ Đạt (+{abs(gap):,})</span>'
            bar_color = "linear-gradient(90deg, #10b981 0%, #059669 100%)"
        else:
            card_class = "t-card-missing"
            status_badge = f'<span class="t-badge t-badge-missing">✕ Thiếu {gap:,} đơn</span>'
            bar_color = "linear-gradient(90deg, #f59e0b 0%, #ea580c 100%)"

        target_cards_html += f"""
        <div class="t-card {card_class}">
            <div class="t-card-head">
                <span class="t-pct-chip">{pct}%</span>
                {status_badge}
            </div>
            <div class="t-card-body">
                <div class="t-label">MỤC TIÊU CẦN ĐẠT</div>
                <div class="t-vol-val">{t_vol:,} <span class="t-unit">đơn</span></div>
                <div class="t-bar-wrap">
                    <div class="t-bar" style="width: {progress_pct}%; background: {bar_color};"></div>
                </div>
                <div class="t-bar-label">Đã đạt: <b>{progress_pct:.1f}%</b></div>
            </div>
        </div>
        """

    # Ghi chú mốc target nâng
    elevated_badge = ""
    if min(milestones) > 40:
        elevated_badge = f'<div class="elevated-badge">🔥 Nâng target lên mốc {max(milestones)}%</div>'

    # HTML danh sách nhân viên
    rows_html = ""
    for idx, s in enumerate(staff, start=1):
        ma_nv = s["ma_nv"]
        name = s["name"]
        gan = s["gan"]
        tc = s["tc"]
        pct = s["pct"]
        ltc = s["ltc"]
        danh_gia = s["danh_gia"]

        name_html = f'<span class="emp-code">{ma_nv}</span> <span class="emp-name">{name}</span>' if ma_nv else f'<span class="emp-name">{name}</span>'

        if pct >= 85:
            pct_badge = f'<span class="badge badge-success">{pct:.2f}%</span>'
        elif pct >= 80:
            pct_badge = f'<span class="badge badge-warning">{pct:.2f}%</span>'
        else:
            pct_badge = f'<span class="badge badge-danger">{pct:.2f}%</span>'

        dg_lower = str(danh_gia).lower()
        if "ok" in dg_lower:
            dg_badge = '<span class="badge badge-dg-ok">✓ OK</span>'
        elif "đạt" in dg_lower or "dat" in dg_lower:
            dg_badge = '<span class="badge badge-dg-dat">✓ Đạt</span>'
        elif "thấp" in dg_lower or "thap" in dg_lower:
            dg_badge = '<span class="badge badge-dg-thap">✕ Thấp</span>'
        elif danh_gia:
            dg_badge = f'<span class="badge badge-gray">{danh_gia}</span>'
        else:
            dg_badge = '<span style="color:#cbd5e1;">-</span>'

        even_cls = "row-even" if idx % 2 == 0 else "row-odd"
        rows_html += f"""
        <tr class="{even_cls}">
            <td class="col-stt">{idx}</td>
            <td class="col-name">{name_html}</td>
            <td class="col-num col-gan">{gan:,}</td>
            <td class="col-num col-tc">{tc:,}</td>
            <td class="col-pct">{pct_badge}</td>
            <td class="col-num col-ltc">{ltc:,}</td>
            <td class="col-dg">{dg_badge}</td>
        </tr>
        """

    overall_pct_badge_cls = "badge-solid-emerald" if pct_now >= 85 else ("badge-solid-gold" if pct_now >= 80 else "badge-solid-rose")

    html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:ital,wght@0,400;0,500;0,600;0,700;0,800;1,400&family=Plus+Jakarta+Sans:wght@500;600;700;800&display=swap" rel="stylesheet">
    <style>
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, sans-serif;
            -webkit-font-smoothing: antialiased;
        }}

        body {{
            background: #f4efe6;
            padding: 24px;
            display: inline-block;
            min-width: 1240px;
        }}

        .container {{
            background: #ffffff;
            border-radius: 24px;
            box-shadow: 0 20px 45px -12px rgba(120, 53, 15, 0.22), 0 0 0 1px rgba(120, 53, 15, 0.1);
            overflow: hidden;
            width: 1240px;
        }}

        /* HEADER SANG TRỌNG (DARK AMBER / GOLD GRADIENT) */
        .header {{
            background: linear-gradient(135deg, #451a03 0%, #78350f 30%, #b45309 65%, #d97706 100%);
            color: #ffffff;
            padding: 26px 36px 22px 36px;
            position: relative;
            border-bottom: 3px solid #f59e0b;
        }}

        .header-top {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 18px;
        }}

        .header-title-box {{
            flex: 1;
        }}

        .tags-row {{
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 8px;
        }}

        .title-tag {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: rgba(254, 240, 138, 0.22);
            border: 1px solid rgba(254, 240, 138, 0.55);
            color: #fef08a;
            font-size: 11px;
            font-weight: 800;
            letter-spacing: 1px;
            text-transform: uppercase;
            padding: 5px 14px;
            border-radius: 99px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }}

        .live-dot {{
            display: inline-block;
            width: 8px;
            height: 8px;
            background-color: #fef08a;
            border-radius: 50%;
            box-shadow: 0 0 8px #fef08a;
        }}

        .elevated-badge {{
            display: inline-flex;
            align-items: center;
            background: #fef08a;
            color: #854d0e;
            font-size: 11px;
            font-weight: 800;
            padding: 5px 12px;
            border-radius: 99px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.15);
        }}

        .main-title {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 28px;
            font-weight: 800;
            letter-spacing: -0.5px;
            color: #ffffff;
            margin-bottom: 8px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.25);
        }}

        .sub-info {{
            display: flex;
            align-items: center;
            gap: 12px;
            color: #fef3c7;
            font-size: 14px;
            font-weight: 500;
        }}

        .info-pill {{
            background: rgba(255, 255, 255, 0.16);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.28);
            padding: 5px 14px;
            border-radius: 10px;
            color: #ffffff;
            font-size: 13px;
            font-weight: 600;
        }}
        .info-pill strong {{
            color: #fef08a;
            font-weight: 800;
        }}

        .date-badge {{
            background: rgba(255, 255, 255, 0.18);
            border: 1px solid rgba(255, 255, 255, 0.3);
            padding: 8px 18px;
            border-radius: 14px;
            color: #fef3c7;
            font-size: 12px;
            font-weight: 700;
            text-align: right;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }}
        .date-badge strong {{
            color: #ffffff;
            display: block;
            font-size: 16px;
            font-weight: 800;
            margin-top: 2px;
        }}
        .date-badge .realtime-label {{
            color: #fde047;
            font-size: 11px;
            font-weight: 800;
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }}

        /* 5 THẺ KPI TỔNG QUAN */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 12px;
            margin-top: 4px;
        }}

        .kpi-card {{
            background: rgba(255, 255, 255, 0.16);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.28);
            border-radius: 14px;
            padding: 12px 16px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }}

        .kpi-card.highlight {{
            background: rgba(254, 240, 138, 0.25);
            border-color: rgba(254, 240, 138, 0.55);
        }}

        .kpi-label {{
            font-size: 11px;
            color: rgba(255, 255, 255, 0.92);
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 4px;
        }}

        .kpi-val {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 23px;
            font-weight: 800;
            color: #ffffff;
        }}

        .kpi-val.accent {{
            color: #fef08a;
            text-shadow: 0 1px 3px rgba(0,0,0,0.25);
        }}
        
        .kpi-sub {{
            font-size: 11px;
            color: rgba(254, 243, 199, 0.88);
            margin-top: 2px;
            font-weight: 500;
        }}

        /* SECTION 1: TARGET GTC PROGRESS (5 CARDS) */
        .section-target {{
            padding: 20px 28px 16px 28px;
            background: #fffdfa;
            border-bottom: 1px solid #fde68a;
        }}

        .section-title {{
            font-size: 13px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            color: #92400e;
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}

        .target-cards-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 12px;
        }}

        .t-card {{
            border-radius: 14px;
            padding: 12px 14px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            transition: all 0.2s ease;
        }}

        .t-card-achieved {{
            background: #f0fdf4;
            border: 1.5px solid #86efac;
            box-shadow: 0 2px 8px rgba(16, 185, 129, 0.08);
        }}

        .t-card-missing {{
            background: #fff7ed;
            border: 1.5px solid #fdba74;
            box-shadow: 0 2px 8px rgba(234, 88, 12, 0.08);
        }}

        .t-card-head {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        }}

        .t-pct-chip {{
            background: #78350f;
            color: #fef3c7;
            padding: 3px 8px;
            border-radius: 6px;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 12px;
            font-weight: 800;
        }}

        .t-badge {{
            padding: 3px 8px;
            border-radius: 99px;
            font-size: 11px;
            font-weight: 800;
            white-space: nowrap;
        }}

        .t-badge-achieved {{
            background: #dcfce7;
            color: #15803d;
            border: 1px solid #86efac;
        }}

        .t-badge-missing {{
            background: #fee2e2;
            color: #dc2626;
            border: 1px solid #fca5a5;
        }}

        .t-label {{
            font-size: 10px;
            color: #78350f;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.4px;
        }}

        .t-vol-val {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 20px;
            font-weight: 800;
            color: #1e293b;
            margin: 2px 0 6px 0;
        }}
        .t-vol-val .t-unit {{
            font-size: 12px;
            font-weight: 600;
            color: #64748b;
        }}

        .t-bar-wrap {{
            width: 100%;
            height: 6px;
            background: #e2e8f0;
            border-radius: 99px;
            overflow: hidden;
            margin-bottom: 4px;
        }}

        .t-bar {{
            height: 100%;
            border-radius: 99px;
        }}

        .t-bar-label {{
            font-size: 10px;
            color: #64748b;
            text-align: right;
        }}

        /* SECTION 2: BẢNG CHI TIẾT NVPTT */
        .table-container {{
            padding: 20px 28px 24px 28px;
            background: #ffffff;
        }}

        table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid #fde68a;
        }}

        th {{
            background: #fef3c7;
            color: #78350f;
            font-size: 12px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            padding: 12px 14px;
            border-bottom: 2px solid #d97706;
            text-align: left;
        }}

        th.text-right {{ text-align: right; }}
        th.text-center {{ text-align: center; }}

        td {{
            padding: 12px 14px;
            font-size: 13.5px;
            border-bottom: 1px solid #f3ede2;
            color: #0f172a;
            vertical-align: middle;
        }}

        tr.row-even {{ background-color: #ffffff; }}
        tr.row-odd {{ background-color: #fffdf7; }}

        .col-stt {{
            width: 44px;
            text-align: center;
            color: #854d0e;
            font-size: 13px;
            font-weight: 700;
        }}

        .col-name {{
            font-weight: 600;
            color: #0f172a;
        }}

        .emp-code {{
            display: inline-block;
            background: #fef3c7;
            color: #78350f;
            font-family: monospace;
            font-size: 11.5px;
            font-weight: 700;
            padding: 2px 7px;
            border-radius: 6px;
            border: 1px solid #fde047;
            margin-right: 6px;
        }}

        .emp-name {{
            font-weight: 700;
            color: #1e293b;
        }}

        .col-num {{
            text-align: right;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 14.5px;
            font-weight: 700;
        }}

        .col-gan {{ color: #1e293b; }}
        .col-tc {{ color: #b45309; }}
        .col-ltc {{ color: #475569; }}

        .col-pct {{
            text-align: right;
            width: 125px;
        }}

        .col-dg {{
            text-align: center;
            width: 130px;
        }}

        .badge {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 4px 10px;
            border-radius: 99px;
            font-size: 12px;
            font-weight: 700;
            white-space: nowrap;
        }}

        .badge-success {{
            background: #dcfce7;
            color: #15803d;
            border: 1px solid #86efac;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }}

        .badge-warning {{
            background: #fef3c7;
            color: #b45309;
            border: 1px solid #fde047;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }}

        .badge-danger {{
            background: #ffe4e6;
            color: #be123c;
            border: 1px solid #fda4af;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }}

        .badge-dg-ok {{
            background: #dcfce7;
            color: #166534;
            border: 1px solid #86efac;
            font-weight: 800;
        }}

        .badge-dg-dat {{
            background: #fef3c7;
            color: #854d0e;
            border: 1px solid #fde047;
            font-weight: 800;
        }}

        .badge-dg-thap {{
            background: #fee2e2;
            color: #991b1b;
            border: 1px solid #fca5a5;
            font-weight: 800;
        }}

        .badge-gray {{
            background: #f1f5f9;
            color: #475569;
            border: 1px solid #cbd5e1;
        }}

        .badge-solid-emerald {{
            background: #16a34a;
            color: #ffffff;
            box-shadow: 0 2px 4px rgba(22, 163, 74, 0.2);
        }}

        .badge-solid-gold {{
            background: #d97706;
            color: #ffffff;
            box-shadow: 0 2px 4px rgba(217, 119, 6, 0.2);
        }}

        .badge-solid-rose {{
            background: #dc2626;
            color: #ffffff;
            box-shadow: 0 2px 4px rgba(220, 38, 38, 0.2);
        }}

        .summary-row {{
            background: #fef3c7 !important;
            color: #78350f;
        }}

        .summary-row td {{
            color: #78350f;
            font-weight: 800;
            border-top: 2px solid #d97706;
            border-bottom: 2px solid #d97706;
            padding: 13px 14px;
        }}

        .summary-label {{
            font-size: 13.5px;
            font-weight: 800;
            letter-spacing: 0.5px;
            text-transform: uppercase;
            color: #78350f;
        }}

        .summary-val {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 15px;
            font-weight: 800;
            text-align: right;
            color: #78350f;
        }}

        .footer {{
            background: #fffaf0;
            padding: 10px 28px;
            font-size: 11px;
            color: #92400e;
            display: flex;
            justify-content: space-between;
            border-top: 1px solid #fde68a;
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- HEADER VÀNG ĐẬM HOÀNG GIA -->
        <div class="header">
            <div class="header-top">
                <div class="header-title-box">
                    <div class="tags-row">
                        <div class="title-tag">
                            <span class="live-dot"></span>
                            ⚡ NĂNG SUẤT REAL-TIME & TARGET (CA 1 + TỒN)
                        </div>
                        {elevated_badge}
                    </div>
                    <h1 class="main-title">BƯU CỤC: {bc_name}</h1>
                    <div class="sub-info">
                        <div class="info-pill">📍 Tỉnh: <strong>{tinh if tinh else 'NTB'}</strong></div>
                        <div class="info-pill">👤 AM Quản lý: <strong>{am_name}</strong></div>
                        <div class="info-pill">👥 Quy mô: <strong>{total_rows} NVPTT</strong></div>
                    </div>
                </div>
                <div class="date-badge">
                    <div class="realtime-label">CẬP NHẬT REAL-TIME</div>
                    <strong>{update_time_str}</strong>
                    <div style="font-size: 11px; margin-top: 2px; color: #fef3c7;">So sánh N-1: {date_n1} ({pct_gtc_n1}%)</div>
                </div>
            </div>

            <!-- 5 KPI TỔNG QUAN -->
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">% GTC N-1 (CA 1 + TỒN)</div>
                    <div class="kpi-val">{pct_gtc_n1}%</div>
                    <div class="kpi-sub">Ca 1 + Tồn hôm trước</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">HÀNG VỀ (CA 1 + TỒN)</div>
                    <div class="kpi-val">{vol_today:,}</div>
                    <div class="kpi-sub">Chốt đầu ngày 09h00</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">ĐÃ GIAO HIỆN TẠI</div>
                    <div class="kpi-val accent">{gtc_now:,}</div>
                    <div class="kpi-sub">Số đơn đã hoàn tất</div>
                </div>
                <div class="kpi-card highlight">
                    <div class="kpi-label">% GTC HIỆN TẠI (CA 1 + TỒN)</div>
                    <div class="kpi-val accent">{pct_now:.2f}%</div>
                    <div class="kpi-sub">Tính trên hàng đầu ngày</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">ĐÁNH GIÁ NHÂN SỰ</div>
                    <div class="kpi-val" style="font-size: 17px; line-height: 1.3; margin-top: 2px; color: #fef08a;">
                        {eval_summary}
                    </div>
                    <div class="kpi-sub">Tổng LTC: {total_ltc:,} đơn</div>
                </div>
            </div>
        </div>

        <!-- SECTION 1: TARGET GTC THEO CÁC MỐC -->
        <div class="section-target">
            <div class="section-title">
                <span>🎯 TIẾN ĐỘ TARGET GTC (CA 1 + TỒN) THEO CÁC MỐC</span>
                <span style="font-size: 11px; font-weight: 600; text-transform: none; color: #b45309;">(Cơ số tính: Hàng Ca 1 + Hàng Tồn chốt đầu ngày)</span>
            </div>
            <div class="target-cards-grid">
                {target_cards_html}
            </div>
        </div>

        <!-- SECTION 2: BẢNG CHI TIẾT NĂNG SUẤT NVPTT -->
        <div class="table-container">
            <div class="section-title" style="margin-bottom: 10px;">
                <span>👥 BẢNG CHI TIẾT NĂNG SUẤT TỪNG NHÂN VIÊN PTT (REAL-TIME)</span>
                <span style="font-size: 11px; font-weight: 600; text-transform: none; color: #b45309;">Đạt chuẩn: % GTC ≥ 85% (Xanh lá) | 80% - 85% (Vàng) | &lt; 80% (Đỏ)</span>
            </div>
            <table>
                <thead>
                    <tr>
                        <th class="text-center" style="width: 44px;">#</th>
                        <th>Nhân viên PTT</th>
                        <th class="text-right">Gán Giao</th>
                        <th class="text-right">Giao TC</th>
                        <th class="text-right">% GTC</th>
                        <th class="text-right">LTC</th>
                        <th class="text-center">Đánh Giá</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                    <tr class="summary-row">
                        <td class="text-center" style="color: #854d0e;">-</td>
                        <td class="summary-label">TỔNG CỘNG BƯU CỤC</td>
                        <td class="summary-val">{total_gan:,}</td>
                        <td class="summary-val" style="color: #b45309;">{total_tc:,}</td>
                        <td class="summary-val" style="text-align: right;">
                            <span class="badge {overall_pct_badge_cls}" style="font-size: 13.5px; padding: 4px 12px; font-weight: 800;">{pct_now:.2f}%</span>
                        </td>
                        <td class="summary-val">{total_ltc:,}</td>
                        <td class="text-center" style="font-weight: 800; font-size: 12.5px;">{eval_summary}</td>
                    </tr>
                </tbody>
            </table>
        </div>

        <div class="footer">
            <span>Auto Report System · Vận Hành NTB</span>
            <span>Hệ thống báo cáo tự động kết hợp Năng Suất Nhân Sự Real-Time & Theo Dõi Target GTC</span>
        </div>
    </div>
</body>
</html>"""
    return html

def render_combined_image(am_name, bc_name, staff, target_info, report_date_str, update_time_str, out_path):
    """Render ảnh kết hợp bằng Playwright."""
    from playwright.sync_api import sync_playwright
    html_content = build_combined_report_html(am_name, bc_name, staff, target_info, report_date_str, update_time_str)

    temp_html = out_path.replace(".png", ".html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(device_scale_factor=2)
        page.goto(f"file:///{os.path.abspath(temp_html)}")
        page.wait_for_timeout(400)
        container = page.query_selector(".container")
        if container:
            container.screenshot(path=out_path)
        else:
            page.screenshot(path=out_path, full_page=True)
        browser.close()

    try:
        if os.path.exists(temp_html):
            os.remove(temp_html)
    except Exception:
        pass

    return out_path

# ===== 4. TẠO CAPTION TIN NHẮN GTALK =====
def build_combined_caption(am_name, bc_name, staff, target_info, update_time_str):
    tinh = target_info.get('tinh', '')
    date_n1 = target_info.get('date_n1', 'N-1')
    pct_gtc_n1 = target_info.get('pct_gtc_n1', 0.0)
    vol_today = target_info.get('vol_today', sum(s['gan'] for s in staff))
    gtc_now = target_info.get('gtc_now', sum(s['tc'] for s in staff))
    pct_now = round(gtc_now / vol_today * 100, 2) if vol_today > 0 else 0.0

    milestones = get_dynamic_milestones(pct_now)
    milestone_lines = []
    for pct in milestones:
        t_vol = int(round(vol_today * (pct / 100.0)))
        gap = t_vol - gtc_now
        if gap <= 0:
            milestone_lines.append(f"• Mốc <b>{pct}%</b>: ✅ Đã đạt (Vượt {abs(gap):,} đơn)")
        else:
            milestone_lines.append(f"• Mốc <b>{pct}%</b>: 🔴 <b>Thiếu {gap:,} đơn</b>")

    milestone_str = "\n".join(milestone_lines)

    tot_ok = sum(1 for s in staff if "ok" in str(s["danh_gia"]).lower())
    tot_dat = sum(1 for s in staff if "đạt" in str(s["danh_gia"]).lower() or "dat" in str(s["danh_gia"]).lower())
    tot_thap = sum(1 for s in staff if "thấp" in str(s["danh_gia"]).lower() or "thap" in str(s["danh_gia"]).lower())
    total_ltc = sum(s["ltc"] for s in staff)

    eval_txt = f"{tot_ok} OK"
    if tot_dat > 0:
        eval_txt += f" · {tot_dat} Đạt"
    if tot_thap > 0:
        eval_txt += f" · {tot_thap} Thấp"

    tinh_str = f" ({tinh})" if tinh else ""

    caption = (
        f"⚡ <b>BÁO CÁO NĂNG SUẤT & TARGET GTC (CA 1 + TỒN)</b>\n"
        f"🏢 Bưu cục: <b>{bc_name}</b>{tinh_str}\n"
        f"👤 AM: <b>{am_name}</b> | 🕒 Cập nhật: <b>{update_time_str}</b>\n\n"
        f"📊 <b>TIẾN ĐỘ GTC (CA 1 + TỒN):</b>\n"
        f"• % GTC Hôm qua ({date_n1}): <b>{pct_gtc_n1}% (Ca 1 + Tồn)</b>\n"
        f"• Hàng đầu ngày (Ca 1 + Tồn): <b>{vol_today:,} đơn</b> | Đã giao: <b>{gtc_now:,} đơn ({pct_now:.2f}%)</b>\n\n"
        f"🎯 <b>SỐ ĐƠN CÒN THIẾU THEO TARGET (CA 1 + TỒN):</b>\n"
        f"{milestone_str}"
    )
    return caption

# ===== 5. GỬI ẢNH LÊN GTALK (HỖ TRỢ RETRY) =====
def upload_image_to_gtalk(image_path, channel_id, oa_token, session=None):
    if session is None:
        session = get_http_session()

    file_name = os.path.basename(image_path)
    file_size = os.path.getsize(image_path)
    with open(image_path, "rb") as f:
        file_bytes = f.read()

    with Image.open(image_path) as im:
        width, height = im.size

    init_payload = {
        "ChannelId": channel_id,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": width, "height": height}),
        "oaToken": oa_token
    }

    presigned_url = None
    upload_id = None

    for attempt in range(1, 4):
        try:
            resp_init = session.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload, timeout=30)
            if resp_init.status_code == 200:
                init_data = resp_init.json()
                if init_data.get("errorCode") == "success":
                    presigned_url = init_data["data"]["PresignedURL"]
                    upload_id = init_data["data"]["UploadId"]
                    break
            time.sleep(attempt * 2)
        except Exception:
            time.sleep(attempt * 2)
    else:
        return None, None

    for attempt in range(1, 4):
        try:
            resp_put = session.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"}, timeout=60)
            if resp_put.status_code == 200:
                break
            time.sleep(attempt * 2)
        except Exception:
            time.sleep(attempt * 2)
    else:
        return None, None

    for attempt in range(1, 4):
        try:
            resp_comp = session.post(
                "https://mbff.ghn.vn/api/gtalk/complete-upload",
                json={"oaToken": oa_token, "UploadId": upload_id},
                timeout=30
            )
            if resp_comp.status_code == 200:
                comp_data = resp_comp.json()
                if comp_data.get("errorCode") == "success":
                    return comp_data["data"]["Id"], (width, height)
            time.sleep(attempt * 2)
        except Exception:
            time.sleep(attempt * 2)
    return None, None

def send_report_to_gtalk(image_path, caption, channel_id, oa_token=None, session=None):
    if session is None:
        session = get_http_session()

    oa_token = oa_token or GTALK_OA_TOKEN
    file_id, size = upload_image_to_gtalk(image_path, channel_id, oa_token, session=session)
    if not file_id:
        return False

    width, height = size
    send_payload = {
        "channelId": channel_id,
        "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
        "content": {
            "parseMode": "HTML",
            "attachment": {
                "caption": caption,
                "items": [
                    {"image": {"fileId": file_id, "width": width, "height": height}}
                ]
            }
        },
        "oaToken": oa_token
    }

    for attempt in range(1, 4):
        try:
            r_send = session.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload, timeout=30)
            if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
                return True
            time.sleep(attempt * 2)
        except Exception:
            time.sleep(attempt * 2)

    return False

# ===== MAIN CHÍNH =====
def main():
    parser = argparse.ArgumentParser(description="Báo cáo Năng Suất NVPTT & Target GTC kết hợp")
    parser.add_argument("--limit", type=int, default=0, help="Giới hạn số bưu cục chạy test")
    parser.add_argument("--bc", type=str, default="", help="Chỉ định 1 bưu cục cụ thể (ví dụ: --bc 'Bắc Bình')")
    parser.add_argument("--am", type=str, default="", help="Chỉ định 1 AM cụ thể")
    parser.add_argument("--no-send", action="store_true", help="Chỉ tạo ảnh và in caption test, KHÔNG gửi thật lên GTalk")
    parser.add_argument("--out-dir", type=str, default="", help="Thư mục xuất ảnh")
    args = parser.parse_args()

    now = datetime.now()
    report_date_str = now.strftime("%d/%m/%Y")
    update_time_str = now.strftime("%H:%M · %d/%m/%Y")

    print("=" * 70, flush=True)
    print(f"🚀 BẮT ĐẦU CHẠY BÁO CÁO NĂNG SUẤT & TARGET GTC KẾT HỢP ({update_time_str})", flush=True)
    if args.no_send:
        print("🛡️ [CHẾ ĐỘ TEST - KHÔNG GỬI THỰC TẾ LÊN GTALK]", flush=True)
    print("=" * 70, flush=True)

    # 1. Đọc dữ liệu Target
    target_data_map, latest_date_n1 = load_target_data()

    # 2. Đọc dữ liệu Năng Suất Realtime
    rt_data = read_baocao_realtime()

    # 3. Lọc danh sách công việc
    jobs = []
    for am_block in rt_data:
        am_name = am_block["am"]
        if args.am and args.am.lower() not in am_name.lower():
            continue
        for bc_name, staff in am_block["bcs"]:
            if args.bc and args.bc.lower() not in bc_name.lower():
                continue

            # Tìm target tương ứng
            norm_bc = unicodedata.normalize('NFC', bc_name).strip()
            t_info = target_data_map.get(norm_bc)
            if not t_info:
                for k, v in target_data_map.items():
                    if k in norm_bc or norm_bc in k:
                        t_info = v
                        break
            if not t_info:
                t_info = {
                    'bc_name': bc_name,
                    'bc_id': '',
                    'tinh': '',
                    'am': am_name,
                    'date_n1': latest_date_n1,
                    'pct_gtc_n1': 0.0,
                    'vol_today': sum(s['gan'] for s in staff),
                    'gtc_now': sum(s['tc'] for s in staff)
                }

            jobs.append((am_name, bc_name, staff, t_info))

    if args.limit > 0:
        jobs = jobs[:args.limit]

    print(f"\n📦 Tìm thấy {len(jobs)} bưu cục phù hợp điều kiện lọc.", flush=True)

    out_dir = args.out_dir or os.path.join(BASE_DIR, "output_combined_reports")
    os.makedirs(out_dir, exist_ok=True)

    session = get_http_session()
    success_count = 0

    for i, (am_name, bc_name, staff, t_info) in enumerate(jobs, start=1):
        safe_name = "".join(c if c.isalnum() else "_" for c in f"{am_name}_{bc_name}")
        out_path = os.path.join(out_dir, f"combined_{safe_name}.png")

        print(f"\n[{i}/{len(jobs)}] 🎨 Đang tạo ảnh kết hợp: AM '{am_name}' · Bưu cục '{bc_name}' ({len(staff)} NVPTT)...", flush=True)
        render_combined_image(am_name, bc_name, staff, t_info, report_date_str, update_time_str, out_path)
        print(f"   ✅ Đã tạo ảnh thành công: {out_path}", flush=True)

        caption = build_combined_caption(am_name, bc_name, staff, t_info, update_time_str)

        if args.no_send:
            print("\n" + "—" * 50)
            print("📝 [CAPTION MẪU GỬI GTALK]:")
            print(caption)
            print("—" * 50 + "\n")
            success_count += 1
            continue

        channel_id = get_channel_for_am(am_name)
        print(f"📡 Đang gửi ảnh vào GTalk (Channel: {channel_id})...", flush=True)
        ok = send_report_to_gtalk(out_path, caption, channel_id=channel_id, session=session)
        if ok:
            print(f"   🎉 Đã gửi thành công bưu cục '{bc_name}' sang GTalk!", flush=True)
            success_count += 1
        else:
            print(f"   ❌ Gửi bưu cục '{bc_name}' thất bại.", flush=True)

        time.sleep(1.5)

    print("\n" + "=" * 70, flush=True)
    if args.no_send:
        print(f"🏁 HOÀN TẤT THỬ NGHIỆM! Đã tạo xong {success_count}/{len(jobs)} ảnh mẫu tại thư mục:\n👉 {out_dir}", flush=True)
    else:
        print(f"🏁 HOÀN TẤT GỬI BÁO CÁO! Đã gửi {success_count}/{len(jobs)} bưu cục thành công.", flush=True)
    print("=" * 70, flush=True)

if __name__ == "__main__":
    main()
