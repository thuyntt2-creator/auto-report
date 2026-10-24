# -*- coding: utf-8 -*-
"""
TOOL TỰ ĐỘNG CẢNH BÁO NGUY CƠ GỐI ĐẦU COD CHO TỪNG AM QUA GTALK
==============================================================
Quy trình:
1. Đọc file Excel cảnh báo COD mới nhất (hoặc gọi tao_file_canh_bao_cod.py để tự sinh).
2. Tự động đồng bộ & tách tab chi tiết MVĐ theo từng AM lên Google Sheet.
3. Lấy link tab riêng của từng AM.
4. Tự động gom nhóm bưu cục, tổng hợp số tiền COD, đơn hoãn, đơn fake GPS.
5. Gửi tin nhắn cảnh báo trực tiếp vào Channel GTalk của từng AM có tên.

Cách dùng:
  - Xem trước nội dung, không gửi: python gui_canh_bao_goidau_cod.py --dry-run
  - Gửi thật:                       python gui_canh_bao_goidau_cod.py
"""

import os
import sys
import time
import glob
import argparse
import unicodedata
import requests
import urllib3
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

def normalize_text(s):
    """Chuẩn hóa chuỗi tiếng Việt về Unicode chuẩn NFC và xóa khoảng trắng thừa"""
    if s is None:
        return ""
    return unicodedata.normalize('NFC', str(s)).strip()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GOOGLE_SHEET_ID = "1sQXZ9IJTR7v-z7TSYkvEpVOuEY6LxqFWbMHi0IGiD8M"
GTALK_OA_TOKEN = "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
GTALK_API_URL = "https://mbff.ghn.vn/api/gtalk/send-message"

# Danh sách Channel ID GTalk chuẩn mới nhất do User cập nhật
RAW_AM_CHANNELS = {
    "Nguyễn Ngọc Khánh": "2077277510775197696",
    "Nguyễn Duy Long": "2077277718988836864",
    "Lê Thanh Nhựt": "2077277754418147328",
    "Trần Thị Nhung": "2077277827057745920",
    "Huỳnh Thị Kim Chi": "2077277857186131968",
    "Phan Đình Duy": "2077278383281876992",
    "Thái Thị Thanh Thư": "2077277934947827712",
    "Nguyễn Thanh Long": "2077277974325506048",
    "Nguyễn Hoàng Phi": "2077278021170323456",
    "Nguyễn Lê Nguyên Vũ": "2077278095459835904",
    "Lê Văn Trường": "2077278127814696960",
    "Hồng Bích Nga": "2077278157729837056",
    "Lê Hồng Phong": "2111144704592470016",
    "Cao Thị Thanh Thủy": "2083241927281995776",
    "Nguyễn Thị Tuyết Thơ": "2089391817020141568",
    "Trương Quang Linh": "2094079475020627968",
    "Lê Minh Lợi": "2094079507615027200",
    "Nguyễn Đỗ Minh Nghĩa": "2100062122691026944",
    "Phan Nguyễn Yến Nhi": "2105595062412402688"
}

AM_CHANNELS = {normalize_text(k): v for k, v in RAW_AM_CHANNELS.items()}

def get_latest_excel_file():
    """Tự động tìm file cảnh báo gối đầu COD mới nhất trong Downloads hoặc thư mục hiện tại"""
    patterns = [
        os.path.join(os.path.expanduser("~"), "Downloads", "Bao_Cao_Canh_Bao_Goi_Dau_COD_*.xlsx"),
        os.path.join(BASE_DIR, "Bao_Cao_Canh_Bao_Goi_Dau_COD_*.xlsx")
    ]
    candidates = []
    for pat in patterns:
        candidates.extend(glob.glob(pat))
    if candidates:
        candidates.sort(key=os.path.getmtime, reverse=True)
        return candidates[0]
    return None

def get_gspread_client():
    auth_file = os.path.join(BASE_DIR, 'authorized_user.json')
    creds = Credentials.from_authorized_user_file(
        auth_file,
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    return gspread.authorize(creds)

def load_data_from_excel(excel_path):
    print(f"📂 Đang đọc dữ liệu từ file Excel: {excel_path}")
    df_bc = pd.read_excel(excel_path, 'Tong_Hop_Theo_Buu_Cuc')
    df_nv = pd.read_excel(excel_path, 'Top_NVPTTT_Nguy_Co')
    df_detail = pd.read_excel(excel_path, 'Chi_Tiet_Don_Hang_400k')
    
    if 'AM Quản Lý' in df_bc.columns:
        df_bc['AM Quản Lý'] = df_bc['AM Quản Lý'].apply(normalize_text)
    if 'AM Quản Lý' in df_nv.columns:
        df_nv['AM Quản Lý'] = df_nv['AM Quản Lý'].apply(normalize_text)
    if 'AM Quản Lý' in df_detail.columns:
        df_detail['AM Quản Lý'] = df_detail['AM Quản Lý'].apply(normalize_text)
        
    return df_bc, df_nv, df_detail

def load_data_from_sheet():
    print(f"📥 Đang tải dữ liệu trực tiếp từ Google Sheet ({GOOGLE_SHEET_ID})...")
    gc = get_gspread_client()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    
    ws_bc = sh.worksheet('Tong_Hop_Theo_Buu_Cuc')
    vals_bc = ws_bc.get_all_values()
    df_bc = pd.DataFrame(vals_bc[1:], columns=vals_bc[0])
    
    ws_nv = sh.worksheet('Top_NVPTTT_Nguy_Co')
    vals_nv = ws_nv.get_all_values()
    df_nv = pd.DataFrame(vals_nv[1:], columns=vals_nv[0])
    
    ws_detail = sh.worksheet('Chi_Tiet_Don_Hang_400k')
    vals_detail = ws_detail.get_all_values()
    df_detail = pd.DataFrame(vals_detail[1:], columns=vals_detail[0])
    
    for c in ['Tổng Tiền COD Nghi Vấn (VNĐ)', 'Tổng Số Đơn Hoãn Hẹn', 'Số Đơn POD Giả Mạo GPS', 'Số Lượng NVPTTT Rủi Ro']:
        if c in df_bc.columns:
            df_bc[c] = pd.to_numeric(df_bc[c].astype(str).str.replace(',', '').str.replace('.', '').str.replace('đ', '').str.strip(), errors='coerce').fillna(0)
            
    for c in ['Tổng Tiền COD Nghi Vấn (VNĐ)', 'Số Đơn COD >400k Bị Hoãn', 'Số Đơn POD Giả Mạo GPS']:
        if c in df_nv.columns:
            df_nv[c] = pd.to_numeric(df_nv[c].astype(str).str.replace(',', '').str.replace('.', '').str.replace('đ', '').str.strip(), errors='coerce').fillna(0)
            
    if 'AM Quản Lý' in df_bc.columns:
        df_bc['AM Quản Lý'] = df_bc['AM Quản Lý'].apply(normalize_text)
    if 'AM Quản Lý' in df_nv.columns:
        df_nv['AM Quản Lý'] = df_nv['AM Quản Lý'].apply(normalize_text)
    if 'AM Quản Lý' in df_detail.columns:
        df_detail['AM Quản Lý'] = df_detail['AM Quản Lý'].apply(normalize_text)
        
    return df_bc, df_nv, df_detail

def sync_and_split_sheet_by_am(df_detail, force_sync=False):
    print("🔄 Đang lấy link tab chi tiết MVĐ theo từng AM trên Google Sheet...")
    gc = get_gspread_client()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    existing_sheets = {ws.title: ws for ws in sh.worksheets()}
    headers = [list(df_detail.columns)]
    am_links = {}
    
    for am_name, group_df in df_detail.groupby('AM Quản Lý'):
        clean_name = str(am_name).strip()
        if not clean_name or clean_name.lower() == 'nan':
            continue
        tab_title = f"MVĐ - {clean_name}"
        rows_data = headers + group_df.astype(str).values.tolist()
        
        # Nếu tab đã tồn tại và không bắt buộc cập nhật lại thì tái sử dụng tab (tránh vượt quota 60 write/phút của Google Sheet)
        if tab_title in existing_sheets and not force_sync:
            ws = existing_sheets[tab_title]
        else:
            if tab_title in existing_sheets:
                ws = existing_sheets[tab_title]
                ws.clear()
            else:
                ws = sh.add_worksheet(title=tab_title, rows=len(rows_data) + 20, cols=len(headers[0]) + 2)
                existing_sheets[tab_title] = ws
            
            # Thử cập nhật với cơ chế retry nếu gặp lỗi quota 429
            for attempt in range(3):
                try:
                    ws.update(values=rows_data, range_name='A1')
                    time.sleep(1)
                    break
                except Exception as e:
                    if attempt < 2 and '429' in str(e):
                        print(f"  ⚠️ Quota API 429 vượt hạn mức, chờ 30s thử lại...")
                        time.sleep(30)
                    else:
                        raise e
            
        am_links[clean_name] = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/edit#gid={ws.id}"
        
    print(f"✅ Đã có link {len(am_links)} tab AM trên Google Sheet!")
    return am_links

def build_am_message(am_name, df_bc, df_nv, link_tab):
    sub_bc = df_bc[df_bc['AM Quản Lý'].astype(str).str.strip() == am_name]
    sub_nv = df_nv[df_nv['AM Quản Lý'].astype(str).str.strip() == am_name]
    
    if len(sub_bc) == 0 or len(sub_nv) == 0:
        return None
        
    total_cod = sub_bc['Tổng Tiền COD Nghi Vấn (VNĐ)'].sum()
    total_don = sub_bc['Tổng Số Đơn Hoãn Hẹn'].sum()
    total_fake = sub_bc['Số Đơn POD Giả Mạo GPS'].sum()
    
    bcs_str = " & ".join(sub_bc['Tên Bưu Cục'].tolist())
    fake_header = f" | 🚨 <b>{int(total_fake)} đơn POD Fake GPS</b>" if total_fake > 0 else ""
    
    lines = []
    lines.append(f"🚨 <b>[CẢNH BÁO NGUY CƠ GỐI ĐẦU COD] - AM {am_name.upper()}</b>")
    lines.append(f"Khu vực: <b>{bcs_str}</b>")
    lines.append(f"Tổng COD nghi vấn: <b>{total_cod:,.0f} đ</b> | Số đơn: <b>{int(total_don)} đơn</b>{fake_header}\n")
    
    for bc_name, group_bc in sub_nv.groupby('Tên Bưu Cục'):
        bc_info = sub_bc[sub_bc['Tên Bưu Cục'] == bc_name]
        bc_cod = bc_info['Tổng Tiền COD Nghi Vấn (VNĐ)'].values[0] if len(bc_info) > 0 else group_bc['Tổng Tiền COD Nghi Vấn (VNĐ)'].sum()
        bc_don = bc_info['Tổng Số Đơn Hoãn Hẹn'].values[0] if len(bc_info) > 0 else group_bc['Số Đơn COD >400k Bị Hoãn'].sum()
        bc_fake = bc_info['Số Đơn POD Giả Mạo GPS'].values[0] if len(bc_info) > 0 else 0
        fake_tag = f" | 🚨 {int(bc_fake)} Fake GPS" if bc_fake > 0 else ""
        
        lines.append(f"📍 <b>{bc_name}</b> ({bc_cod:,.0f} đ - {int(bc_don)} đơn{fake_tag}):")
        
        sorted_nv = group_bc.sort_values(by='Tổng Tiền COD Nghi Vấn (VNĐ)', ascending=False)
        for _, r in sorted_nv.iterrows():
            f_val = int(r.get('Số Đơn POD Giả Mạo GPS', 0))
            f_str = f" | 🚨 {f_val} Fake GPS" if f_val > 0 else ""
            lines.append(f"• MSNV: {r.get('Mã NVPTTT', '')} - {r['Họ Tên NVPTTT']} | COD: {r['Tổng Tiền COD Nghi Vấn (VNĐ)']:,.0f} đ ({int(r['Số Đơn COD >400k Bị Hoãn'])} đơn){f_str}")
        lines.append("")
        
    lines.append("🔗 <b>Link tra cứu chi tiết danh sách MVĐ:</b>")
    lines.append(f'<a href="{link_tab}"><b>Bấm vào đây để mở Sheet MVĐ chi tiết của AM {am_name}</b></a>\n')
    lines.append("⚠️ <b>Các đơn trên có nguy cơ cao bưu tá đã thu COD nhưng báo lý do ảo để gối đầu tiền: Yêu cầu AM đối soát ngay giữa HÀNG THỰC TẾ TRONG KHO và gọi điện phúc tra trực tiếp khách hàng trước 11h30 trưa nay!</b>")
    
    return "\n".join(lines)

def send_gtalk_message(channel_id, text):
    payload = {
        "channelId": str(channel_id),
        "clientMsgId": str(int(time.time() * 1000)),
        "content": {"parseMode": "HTML", "text": text},
        "oaToken": GTALK_OA_TOKEN
    }
    r = requests.post(GTALK_API_URL, json=payload, timeout=20, verify=False)
    return r.json()

def main():
    parser = argparse.ArgumentParser(description="Tự động gửi cảnh báo gối đầu COD cho AM qua GTalk")
    parser.add_argument("--file", help="Đường dẫn file Excel cảnh báo COD.")
    parser.add_argument("--sheet", action="store_true", help="Lấy dữ liệu trực tiếp từ Google Sheet")
    parser.add_argument("--dry-run", action="store_true", help="Xem trước nội dung tin nhắn cảnh báo, không gửi GTalk thật")
    args = parser.parse_args()
    
    if args.sheet:
        df_bc, df_nv, df_detail = load_data_from_sheet()
    else:
        excel_path = args.file or get_latest_excel_file()
        if excel_path and os.path.exists(excel_path):
            df_bc, df_nv, df_detail = load_data_from_excel(excel_path)
        else:
            print("⚠️ Không tìm thấy file Excel hợp lệ. Tự động chuyển sang đọc từ Google Sheet...")
            df_bc, df_nv, df_detail = load_data_from_sheet()
    
    am_links = sync_and_split_sheet_by_am(df_detail)
    
    print(f"\n🚀 Bắt đầu xử lý tin nhắn cảnh báo cho từng AM...")
    sent_count = 0
    
    for am_name in df_bc['AM Quản Lý'].unique():
        am_name_clean = str(am_name).strip()
        if not am_name_clean or am_name_clean.lower() == 'nan':
            continue
            
        channel_id = AM_CHANNELS.get(am_name_clean)
        if not channel_id:
            print(f"⚠️ Chưa có Channel ID cho AM: {am_name_clean}")
            continue
            
        link_tab = am_links.get(am_name_clean, f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/edit#gid=0")
        msg = build_am_message(am_name_clean, df_bc, df_nv, link_tab)
        
        if not msg:
            continue
            
        if args.dry_run:
            print(f"\n--- [DRY-RUN XEM TRƯỚC] AM {am_name_clean} (Channel: {channel_id}) ---")
            print(msg[:300] + "\n...")
            sent_count += 1
            continue
            
        print(f"📡 Đang gửi cho AM {am_name_clean} (Channel: {channel_id})...", end="")
        try:
            res = send_gtalk_message(channel_id, msg)
            if res.get("errorCode") == "success":
                print(" -> ✅ THÀNH CÔNG!")
                sent_count += 1
            else:
                print(f" -> ❌ THẤT BẠI: {res}")
        except Exception as e:
            print(f" -> ❌ LỖI: {e}")
            
        time.sleep(0.5)
        
    status_str = "xem trước thành công" if args.dry_run else "gửi cảnh báo thành công"
    print(f"\n🎉 HOÀN THÀNH: Đã {status_str} cho {sent_count} AM!")

if __name__ == "__main__":
    main()
