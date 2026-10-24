# -*- coding: utf-8 -*-
"""
TOOL TỰ ĐỘNG LẤY DỮ LIỆU & TẠO FILE EXCEL CẢNH BÁO GỐI ĐẦU COD & FAKE GPS (MỤC #16)
==================================================================================
Chế độ mặc định: TEST AN TOÀN (CHỈ XUẤT FILE EXCEL, TUYỆT ĐỐI KHÔNG GỬI GROUP CHAT)

Quy trình:
1. Kết nối API / Dashboard để lấy dữ liệu bóc tách rủi ro COD mới nhất.
2. Tự động tính toán & phân loại rủi ro theo Shipper -> Bưu cục -> AM.
3. Xuất file Excel chuẩn 4 Sheet:
   - Tong_Hop_Theo_Buu_Cuc
   - Top_NVPTTT_Nguy_Co
   - Don_That_Bai_Tren_1Trieu
   - Chi_Tiet_Don_Hang_400k
4. Lưu file trực tiếp vào thư mục Downloads của người dùng.
"""

import os
import sys
import re
import json
import time
from datetime import datetime
import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
if not os.path.exists(DOWNLOADS_DIR):
    DOWNLOADS_DIR = os.path.join(BASE_DIR, "downloads")
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)
DASHBOARD_API_URL = "https://strongly-restrict-ads-massage.trycloudflare.com/api/preview-report?id=16"


def fetch_report_content():
    """Lấy dữ liệu bóc tách Mục #16 từ API Dashboard"""
    print("📡 Đang kết nối lấy dữ liệu phân tích Lastmile từ Dashboard Mục #16...")
    try:
        r = requests.get(DASHBOARD_API_URL, timeout=15)
        if r.status_code == 200:
            data = r.json()
            content = data.get("content", "")
            if content and "CẢNH BÁO NGUY CƠ GỐI ĐẦU COD" in content:
                print(f"✅ Đã tải thành công dữ liệu ({len(content):,} ký tự).")
                return content
    except Exception as e:
        print(f"⚠️ Lỗi kết nối Dashboard API: {e}")

    # Fallback to local snapshot if network fails
    local_preview = os.path.join(BASE_DIR, "scratch", "preview_16_full.txt")
    if os.path.exists(local_preview):
        print("📂 Đang sử dụng dữ liệu snapshot dự phòng từ máy...")
        with open(local_preview, "r", encoding="utf-8") as f:
            return f.read()

    return ""

def parse_cod_data(text):
    """Bóc tách dữ liệu từ bản tin thành các DataFrame chuẩn"""
    if not text:
        return None, None, None, None, {}

    # Metadata
    meta = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_don_400k": 0,
        "total_cod": 0,
        "total_fake_gps": 0
    }

    m_time = re.search(r'Dữ liệu bóc tách lúc:\s*`([^`]+)`', text)
    if m_time:
        meta["timestamp"] = m_time.group(1).strip()

    m_don = re.search(r'Số đơn COD > 400\.000đ bị hoãn/hẹn:\*\*\s*`([\d,]+)`', text)
    if m_don:
        meta["total_don_400k"] = int(m_don.group(1).replace(',', ''))

    m_cod = re.search(r'Tổng tiền COD nghi vấn giữ lại:\*\*\s*`([\d,]+)`', text)
    if m_cod:
        meta["total_cod"] = int(m_cod.group(1).replace(',', ''))

    m_fake = re.search(r'🚨\s*\*\*(\d+)\s+đơn bị phát hiện giả mạo toạ độ\*\*', text)
    if m_fake:
        meta["total_fake_gps"] = int(m_fake.group(1))

    # 1. Bóc tách Bưu Cục
    # Example: 1. 🏢 **(DNO) Đức Lập** (Mã BC: `2309`) - AM: **Trần Thị Nhung**
    #    • Tổng COD giữ: **115,103,386 VNĐ** (131 đơn - 12 shipper)
    bc_blocks = re.findall(
        r'(\d+)\.\s+🏢\s+\*\*(.*?)\*\*\s+\(Mã BC:\s+`(\d+)`\)\s+-\s+AM:\s+\*\*(.*?)\*\*\s*\n\s+•\s+Tổng COD giữ:\s+\*\*([\d,]+)\s+VNĐ\*\*\s+\((\d+)\s+đơn\s+-\s+(\d+)\s+shipper\)',
        text, re.UNICODE
    )

    bc_rows = []
    for rank, bc_name, bc_id, am, cod_str, don_str, nv_str in bc_blocks:
        cod_val = int(cod_str.replace(',', ''))
        don_val = int(don_str.replace(',', ''))
        nv_val = int(nv_str.replace(',', ''))
        
        # Đánh giá nguy cơ
        risk = "🔴 RẤT CAO" if cod_val >= 10000000 or don_val >= 20 else "🟡 CẢNH BÁO"

        bc_rows.append({
            "Mã Bưu Cục": int(bc_id),
            "Tên Bưu Cục": bc_name.strip(),
            "AM Quản Lý": am.strip(),
            "Số Lượng NVPTTT Rủi Ro": nv_val,
            "Tổng Số Đơn Hoãn Hẹn": don_val,
            "Tổng Tiền COD Nghi Vấn (VNĐ)": cod_val,
            "Số Đơn POD Giả Mạo GPS": 0,  # Sẽ cộng dồn từ shipper
            "Số Đơn Đã Có POD": 0,
            "Đánh Giá Nguy Cơ": risk
        })

    # 2. Bóc tách từng Shipper theo từng Bưu Cục
    # Tách đoạn theo bưu cục để gán đúng Bưu cục & AM cho shipper
    sections = re.split(r'\n(?=\d+\.\s+🏢)', text)
    nv_rows = []
    bc_fake_map = {}

    for sec in sections[1:]:
        header_m = re.search(r'\d+\.\s+🏢\s+\*\*(.*?)\*\*\s+\(Mã BC:\s+`(\d+)`\)\s+-\s+AM:\s+\*\*(.*?)\*\*', sec)
        if not header_m:
            continue
        cur_bc_name = header_m.group(1).strip()
        cur_bc_id = int(header_m.group(2))
        cur_am = header_m.group(3).strip()

        # Parse shipper lines in this section
        # Example: ↳ [🔴 RẤT CAO] **Lê Hữu Hiệp** (0832387287): **16,246,681 VNĐ** (14 đơn: 10 hẹn, 1 không nghe | 🚨 **2 đơn FAKE GPS**)
        shippers = re.findall(
            r'↳\s+\[(.*?)\]\s+\*\*(.*?)\*\*\s+\((\d+)\):\s+\*\*([\d,]+)\s+VNĐ\*\*\s+\((\d+)\s+đơn:\s+(\d+)\s+hẹn,\s+(\d+)\s+không nghe(?:\s+\|\s+🚨\s+\*\*(\d+)\s+đơn\s+FAKE\s+GPS\*\*)?\)',
            sec, re.UNICODE
        )

        for dg, name, phone, cod_s, don_s, hen_s, knghe_s, fake_s in shippers:
            s_cod = int(cod_s.replace(',', ''))
            s_don = int(don_s.replace(',', ''))
            s_hen = int(hen_s.replace(',', ''))
            s_knghe = int(knghe_s.replace(',', ''))
            s_fake = int(fake_s) if fake_s else 0
            
            clean_phone = phone.strip()
            if not clean_phone.startswith('0') and len(clean_phone) == 9:
                clean_phone = '0' + clean_phone

            bc_fake_map[cur_bc_id] = bc_fake_map.get(cur_bc_id, 0) + s_fake

            nv_rows.append({
                "Mã Bưu Cục": cur_bc_id,
                "Tên Bưu Cục": cur_bc_name,
                "AM Quản Lý": cur_am,
                "Mã NVPTTT": "",  # Tự sinh nếu không có MSNV trực tiếp
                "Họ Tên NVPTTT": name.strip(),
                "Số Điện Thoại": clean_phone,
                "Số Đơn COD >400k Bị Hoãn": s_don,
                "Tổng Tiền COD Nghi Vấn (VNĐ)": s_cod,
                "Số Đơn Báo Khách Hẹn": s_hen,
                "Số Đơn Báo Không Nghe Máy": s_knghe,
                "Số Đơn POD Giả Mạo GPS": s_fake,
                "Số Đơn Đã Có POD": s_fake,
                "Đánh Giá Nguy Cơ": f"[{dg.strip()}]" if not dg.strip().startswith('[') else dg.strip()
            })

    # Cập nhật fake GPS cho Bưu cục
    for bc in bc_rows:
        bc["Số Đơn POD Giả Mạo GPS"] = bc_fake_map.get(bc["Mã Bưu Cục"], 0)
        bc["Số Đơn Đã Có POD"] = bc["Số Đơn POD Giả Mạo GPS"]

    # 3. Bóc tách Đơn thất bại trên 1 triệu
    orders_1m = re.findall(
        r'•\s+📦\s+`([A-Z0-9]+)`:\s+\*\*([\d,]+)\s+VNĐ\*\*\s+-\s+(.*?)\s+\((.*?)\)\s+-\s+(.*?)(?=\n|$)',
        text, re.UNICODE
    )

    don_1m_rows = []
    for code, cod_s, shipper_name, bc_name, reason in orders_1m:
        val = int(cod_s.replace(',', ''))
        don_1m_rows.append({
            "Mã Bưu Cục": "",
            "Tên Bưu Cục": bc_name.strip(),
            "AM Quản Lý": "",
            "Họ Tên NVPTTT": shipper_name.strip(),
            "SĐT NVPTTT": "",
            "Mã Đơn Hàng": code.strip(),
            "Tiền COD (VNĐ)": val,
            "Lý Do Cập Nhật": reason.strip(),
            "Phân Loại Rủi Ro": "Khách hẹn giao lại" if "hẹn" in reason.lower() else "Từ chối / Hoãn giao",
            "Có POD": "Có",
            "Giả Mạo GPS": "🚨 CÓ (Giả mạo)" if "fake" in text.lower() else "Không",
            "Thời Gian Upload POD": meta["timestamp"],
            "Mã Chuyến Đi": "",
            "Trạng Thái Chuyến": "FINISHED"
        })

    # Khớp lại AM và SĐT cho danh sách đơn > 1 triệu từ danh sách Shipper
    shipper_info_map = {r["Họ Tên NVPTTT"]: r for r in nv_rows}
    for d in don_1m_rows:
        sh = shipper_info_map.get(d["Họ Tên NVPTTT"])
        if sh:
            d["Mã Bưu Cục"] = sh["Mã Bưu Cục"]
            d["AM Quản Lý"] = sh["AM Quản Lý"]
            d["SĐT NVPTTT"] = sh["Số Điện Thoại"]

    # 4. Quét danh sách Mã Vận Đơn THỰC TẾ từ GHN Lastmile API cho các bưu cục rủi ro
    print("📡 Đang quét chi tiết Mã Vận Đơn (MVĐ) thực tế từ GHN Lastmile API...")
    don_400k_rows = []
    try:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        try:
            from fetch_lastmile_productivity import get_ghn_token
            token_ghn = get_ghn_token()
        except Exception:
            token_ghn = 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJvcmdDb2RlIjoiZ2huZXhwcmVzcyIsInBhcnRuZXJDb2RlIjoiIiwic2VlZCI6NzkzNTk5NTY0OTU2NDAxMjIyOCwic3NvSWQiOiIzMDY2MDIxIiwidXNlcklkIjoiNjRlMmUyYzI2MmNhZDU1YzZiODRlZTBkIn0.CZd7x6iRVCKRap5NkO_wRs3vusfFTVUBMvgfgM8HbSw'
        headers_lm = {'Authorization': token_ghn, 'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'}

        def scan_hub_orders(bc):
            h_id = str(bc["Mã Bưu Cục"])
            h_name = bc["Tên Bưu Cục"]
            h_am = bc["AM Quản Lý"]
            h_orders = []
            s = requests.Session()
            trips = []
            for st in ["FINISHED", "ON_TRIP"]:
                for attempt in range(2):
                    try:
                        r = s.post('https://nhanh-api.ghn.vn/api/lastmile/trip/get-trip-list-by-hub',
                                   headers=headers_lm, json={'hub_id': h_id, 'status': st, 'limit': 35, 'reverse': 1}, timeout=15)
                        if r.status_code == 200:
                            trips.extend(r.json().get('data') or [])
                            break
                    except Exception:
                        time.sleep(0.5)
            
            seen_t = set()
            for t in trips:
                tc = t.get('tripCode')
                if not tc or tc in seen_t:
                    continue
                seen_t.add(tc)
                dname = str(t.get('driverName') or '').strip()
                dphone = str(t.get('driverPhone') or '').strip()
                if not dphone.startswith('0') and len(dphone) == 9:
                    dphone = '0' + dphone
                tstat = str(t.get('status') or 'FINISHED').strip()

                try:
                    r_it = s.post('https://nhanh-api.ghn.vn/api/lastmile/trip/get-trip-items',
                                  headers=headers_lm, json={'tripCode': tc, 'limit': 5000}, timeout=20)
                    if r_it.status_code == 200:
                        for it in (r_it.json().get('data') or []):
                            if it.get('type') == 'DELIVER':
                                is_succ = bool(it.get('isSucceeded'))
                                is_up = bool(it.get('isUpdated'))
                                cod = int(it.get('collectAmount') or 0)
                                if not is_succ and (is_up or it.get('failCode') or it.get('failNote')):
                                    if cod >= 400000:
                                        oc = str(it.get('orderCode') or '').strip()
                                        fn = str(it.get('failNote') or it.get('failCode') or 'Chưa hoàn tất giao hàng').strip()
                                        fn_low = fn.lower()
                                        if 'hẹn' in fn_low:
                                            rc = 'Khách hẹn giao lại'
                                        elif 'không nghe' in fn_low or 'chặn' in fn_low or 'không liên lạc' in fn_low:
                                            rc = 'Không nghe máy'
                                        elif 'đổi ý' in fn_low or 'từ chối' in fn_low or 'không mua' in fn_low:
                                            rc = 'Từ chối nhận hàng'
                                        else:
                                            rc = 'Chưa giao / Hoãn giao'
                                        h_orders.append({
                                            "Mã Đơn Hàng": oc,
                                            "Mã Chuyến Đi": tc,
                                            "Trạng Thái Chuyến": tstat,
                                            "Mã Bưu Cục": h_id,
                                            "Tên Bưu Cục": h_name,
                                            "AM Quản Lý": h_am,
                                            "NVPTTT": dname if dname else (it.get('updateUserName') or 'Shipper tuyến'),
                                            "SĐT NVPTTT": dphone,
                                            "Tiền COD (VNĐ)": cod,
                                            "Lý Do Cập Nhật": fn,
                                            "Phân Loại Rủi Ro": rc,
                                            "Có POD": "Có" if it.get('updatedGeometry') or is_up else "Chưa có",
                                            "Giả Mạo GPS": "Không",
                                            "Thời Gian Upload POD": str(it.get('lastUpdatedTime') or it.get('createdTime') or '---')[:19]
                                        })
                except Exception:
                    pass
            return h_orders

        # Danh sách bưu cục quét thực tế: Giữ nguyên toàn bộ bưu cục rủi ro từ Dashboard + Bổ sung bưu cục Đơn Dương theo yêu cầu rà soát
        EXTRA_MONITORED_HUBS = [
            {
                "Mã Bưu Cục": 21377000,
                "Tên Bưu Cục": "(LDO) Đơn Dương",
                "AM Quản Lý": "Phan Nguyễn Yến Nhi"
            }
        ]

        hubs_to_scan = list(bc_rows)
        existing_bc_ids = {str(b["Mã Bưu Cục"]) for b in hubs_to_scan}
        extra_hubs_to_add = []
        for eh in EXTRA_MONITORED_HUBS:
            if str(eh["Mã Bưu Cục"]) not in existing_bc_ids:
                hubs_to_scan.append(eh)
                extra_hubs_to_add.append(eh)

        with ThreadPoolExecutor(max_workers=6) as executor:
            futs = [executor.submit(scan_hub_orders, b) for b in hubs_to_scan]
            for f in as_completed(futs):
                don_400k_rows.extend(f.result())

        # Loại bỏ trùng mã đơn hàng nếu có
        seen_orders = set()
        unique_400k = []
        for d in don_400k_rows:
            oc = d.get("Mã Đơn Hàng")
            if oc and oc not in seen_orders:
                seen_orders.add(oc)
                unique_400k.append(d)
        don_400k_rows = unique_400k
        print(f"✅ Đã quét thành công {len(don_400k_rows)} đơn hàng thực tế từ Lastmile GHN (bao gồm bưu cục bổ sung)!")

        # Bổ sung thông tin tổng hợp cho các bưu cục rà soát thêm nếu có đơn phát sinh
        for eh in extra_hubs_to_add:
            eh_id_str = str(eh["Mã Bưu Cục"])
            eh_orders = [o for o in don_400k_rows if str(o["Mã Bưu Cục"]) == eh_id_str]
            if eh_orders:
                eh_total_cod = sum(o["Tiền COD (VNĐ)"] for o in eh_orders)
                eh_total_don = len(eh_orders)

                # Gom theo shipper
                sh_groups = {}
                for o in eh_orders:
                    sh_key = (o["NVPTTT"], o["SĐT NVPTTT"])
                    if sh_key not in sh_groups:
                        sh_groups[sh_key] = {
                            "Mã Bưu Cục": eh["Mã Bưu Cục"],
                            "Tên Bưu Cục": eh["Tên Bưu Cục"],
                            "AM Quản Lý": eh["AM Quản Lý"],
                            "Mã NVPTTT": "",
                            "Họ Tên NVPTTT": o["NVPTTT"],
                            "Số Điện Thoại": o["SĐT NVPTTT"],
                            "Số Đơn COD >400k Bị Hoãn": 0,
                            "Tổng Tiền COD Nghi Vấn (VNĐ)": 0,
                            "Số Đơn Báo Khách Hẹn": 0,
                            "Số Đơn Báo Không Nghe Máy": 0,
                            "Số Đơn POD Giả Mạo GPS": 0,
                            "Số Đơn Đã Có POD": 0,
                            "Đánh Giá Nguy Cơ": "🔴 RẤT CAO"
                        }
                    g = sh_groups[sh_key]
                    g["Số Đơn COD >400k Bị Hoãn"] += 1
                    g["Tổng Tiền COD Nghi Vấn (VNĐ)"] += o["Tiền COD (VNĐ)"]
                    if "hẹn" in str(o["Phân Loại Rủi Ro"]).lower():
                        g["Số Đơn Báo Khách Hẹn"] += 1
                    if "không nghe" in str(o["Phân Loại Rủi Ro"]).lower():
                        g["Số Đơn Báo Không Nghe Máy"] += 1

                for sh_info in sh_groups.values():
                    sh_cod = sh_info["Tổng Tiền COD Nghi Vấn (VNĐ)"]
                    sh_don = sh_info["Số Đơn COD >400k Bị Hoãn"]
                    sh_risk = "[🔴 RẤT CAO]" if sh_cod >= 10000000 or sh_don >= 10 else "[🟡 CẢNH BÁO]"
                    sh_info["Đánh Giá Nguy Cơ"] = sh_risk
                    nv_rows.append(sh_info)

                # Thêm vào bc_rows
                bc_risk = "🔴 RẤT CAO" if eh_total_cod >= 10000000 or eh_total_don >= 20 else "🟡 CẢNH BÁO"
                bc_rows.append({
                    "Mã Bưu Cục": eh["Mã Bưu Cục"],
                    "Tên Bưu Cục": eh["Tên Bưu Cục"],
                    "AM Quản Lý": eh["AM Quản Lý"],
                    "Số Lượng NVPTTT Rủi Ro": len(sh_groups),
                    "Tổng Số Đơn Hoãn Hẹn": eh_total_don,
                    "Tổng Tiền COD Nghi Vấn (VNĐ)": eh_total_cod,
                    "Số Đơn POD Giả Mạo GPS": 0,
                    "Số Đơn Đã Có POD": 0,
                    "Đánh Giá Nguy Cơ": bc_risk
                })
    except Exception as e:
        print(f"⚠️ Không thể quét Lastmile API ({e}).")

    df_bc = pd.DataFrame(bc_rows)
    df_nv = pd.DataFrame(nv_rows)
    df_1m = pd.DataFrame(don_1m_rows)
    df_400k = pd.DataFrame(don_400k_rows)
    if len(df_400k) > 0:
        df_400k.drop_duplicates(subset=['Mã Đơn Hàng'], inplace=True)
        df_400k.sort_values(by=['Tiền COD (VNĐ)'], ascending=False, inplace=True)

    return df_bc, df_nv, df_1m, df_400k, meta

def export_styled_excel(df_bc, df_nv, df_1m, df_400k, output_path):
    """Xuất file Excel với định dạng chuyên nghiệp chuẩn GHN"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_bc.to_excel(writer, sheet_name='Tong_Hop_Theo_Buu_Cuc', index=False)
        df_nv.to_excel(writer, sheet_name='Top_NVPTTT_Nguy_Co', index=False)
        df_1m.to_excel(writer, sheet_name='Don_That_Bai_Tren_1Trieu', index=False)
        df_400k.to_excel(writer, sheet_name='Chi_Tiet_Don_Hang_400k', index=False)

    # Style workbook with openpyxl
    wb = openpyxl.load_workbook(output_path)
    
    header_fill = PatternFill(start_color="F26522", end_color="F26522", fill_type="solid") # GHN Orange
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        
        # Style Header
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            ws.row_dimensions[1].height = 28

        # Style Rows
        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 20
            for col_idx in range(1, ws.max_column + 1):
                c = ws.cell(row=row_idx, column=col_idx)
                c.font = data_font
                c.border = thin_border
                
                header_name = str(ws.cell(row=1, column=col_idx).value or '')
                val = c.value

                # Currency formatting
                if any(k in header_name for k in ['Tiền COD', 'VNĐ']):
                    if isinstance(val, (int, float)):
                        c.number_format = '#,##0 "đ"'
                        c.alignment = right_align
                    else:
                        c.alignment = right_align
                # Integer counts
                elif any(k in header_name for k in ['Số Lượng', 'Số Đơn', 'Mã Bưu Cục', 'Mã NVPTTT']):
                    c.alignment = center_align
                    if isinstance(val, (int, float)):
                        c.number_format = '#,##0'
                # Risk Badges
                elif 'Đánh Giá' in header_name or 'Giả Mạo' in header_name:
                    c.alignment = center_align
                    if 'RẤT CAO' in str(val) or 'Giả mạo' in str(val):
                        c.font = Font(name="Calibri", size=10, bold=True, color="9C0006")
                        c.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    elif 'CẢNH BÁO' in str(val):
                        c.font = Font(name="Calibri", size=10, bold=True, color="9C6500")
                        c.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                else:
                    c.alignment = left_align

        # Auto fit column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)

    wb.save(output_path)
    print(f"📁 Đã tạo và định dạng xong file Excel: {output_path}")

GOOGLE_SHEET_ID = "1sQXZ9IJTR7v-z7TSYkvEpVOuEY6LxqFWbMHi0IGiD8M"

def sync_to_google_sheet(df_bc, df_nv, df_1m, df_400k):
    """Đồng bộ 4 sheet chính và tách tab MVĐ theo từng AM lên Google Sheet (KHÔNG GỬI CHAT)"""
    print(f"\n🔄 Đang kết nối và đẩy dữ liệu lên Google Sheet ({GOOGLE_SHEET_ID})...")
    import gspread
    from google.oauth2.credentials import Credentials
    
    auth_file = os.path.join(BASE_DIR, 'authorized_user.json')
    creds = Credentials.from_authorized_user_file(
        auth_file,
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)

    existing_sheets = {ws.title: ws for ws in sh.worksheets()}

    import time
    def safe_update_worksheet(ws_target, data_rows):
        for attempt in range(3):
            try:
                ws_target.update(values=data_rows, range_name='A1')
                time.sleep(0.5)
                break
            except Exception as e:
                if attempt < 2 and '429' in str(e):
                    print("    ⚠️ Quota 429 hit. Waiting 25s before retry...")
                    time.sleep(25)
                else:
                    raise e

    sheets_to_update = [
        ('Tong_Hop_Theo_Buu_Cuc', df_bc),
        ('Top_NVPTTT_Nguy_Co', df_nv),
        ('Don_That_Bai_Tren_1Trieu', df_1m),
        ('Chi_Tiet_Don_Hang_400k', df_400k)
    ]

    for title, df in sheets_to_update:
        print(f"  ⬆️ Đang cập nhật tab '{title}' ({len(df)} dòng)...")
        rows = [list(df.columns)] + df.astype(str).values.tolist()
        if title in existing_sheets:
            ws = existing_sheets[title]
            ws.clear()
        else:
            ws = sh.add_worksheet(title=title, rows=len(rows) + 20, cols=len(df.columns) + 2)
            existing_sheets[title] = ws
        safe_update_worksheet(ws, rows)

    # Tách tab MVĐ theo từng AM
    print("  ✂️ Đang tách tab chi tiết MVĐ cho từng AM...")
    headers_400k = [list(df_400k.columns)]
    for am_name, group_df in df_400k.groupby('AM Quản Lý'):
        clean_name = str(am_name).strip()
        if not clean_name or clean_name.lower() == 'nan':
            continue
        tab_title = f"MVĐ - {clean_name}"
        rows_data = headers_400k + group_df.astype(str).values.tolist()
        if tab_title in existing_sheets:
            ws = existing_sheets[tab_title]
            ws.clear()
        else:
            ws = sh.add_worksheet(title=tab_title, rows=len(rows_data) + 20, cols=len(headers_400k[0]) + 2)
            existing_sheets[tab_title] = ws
        safe_update_worksheet(ws, rows_data)

    print(f"✅ ĐÃ ĐỒNG BỘ THÀNH CÔNG LÊN GOOGLE SHEET: https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/")

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Tạo file Excel và đẩy Google Sheet cảnh báo gối đầu COD")
    parser.add_argument("--upload-sheet", action="store_true", help="Tự động đồng bộ lên Google Sheet")
    args = parser.parse_args()

    print("=" * 70)
    print("🚀 CHƯƠNG TRÌNH ĐỐI SOÁT & TẠO FILE EXCEL CẢNH BÁO GỐI ĐẦU COD (MỤC #16)")
    print("🛡️  CHẾ ĐỘ AN TOÀN: TUYỆT ĐỐI KHÔNG GỬI GROUP CHAT / GTALK")
    print("=" * 70)

    # 1. Lấy dữ liệu
    raw_text = fetch_report_content()
    if not raw_text:
        print("❌ Không lấy được dữ liệu bóc tách.")
        sys.exit(1)

    # 2. Xử lý & Bóc tách
    df_bc, df_nv, df_1m, df_400k, meta = parse_cod_data(raw_text)

    print("\n📊 TỔNG QUAN KẾT QUẢ ĐỐI SOÁT:")
    print(f"• Thời điểm bóc tách: {meta['timestamp']}")
    print(f"• Tổng số bưu cục có rủi ro: {len(df_bc)} bưu cục")
    print(f"• Tổng số NVPTTT có rủi ro: {len(df_nv)} nhân sự")
    print(f"• Tổng số đơn COD > 400k bị hoãn/hẹn: {meta['total_don_400k']} đơn")
    print(f"• Tổng tiền COD nghi vấn giữ lại: {meta['total_cod']:,.0f} VNĐ")
    print(f"• Đơn phát hiện Fake GPS: {meta['total_fake_gps']} đơn")
    print(f"• Đơn thất bại trên 1 triệu: {len(df_1m)} đơn")

    # 3. Tạo file Excel chuẩn
    time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"Bao_Cao_Canh_Bao_Goi_Dau_COD_{time_str}.xlsx"
    
    downloads_path = os.path.join(DOWNLOADS_DIR, file_name)
    local_path = os.path.join(BASE_DIR, file_name)

    export_styled_excel(df_bc, df_nv, df_1m, df_400k, downloads_path)
    export_styled_excel(df_bc, df_nv, df_1m, df_400k, local_path)

    # 4. Đẩy lên Google Sheet nếu có yêu cầu
    if args.upload_sheet:
        sync_to_google_sheet(df_bc, df_nv, df_1m, df_400k)

    print("\n" + "=" * 70)
    print(f"🎉 THÀNH CÔNG! File Excel đối soát đã được lưu tại:")
    print(f"👉 1. {downloads_path}")
    print(f"👉 2. {local_path}")
    print("=" * 70)

if __name__ == "__main__":
    main()
