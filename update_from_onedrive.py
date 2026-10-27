# -*- coding: utf-8 -*-
"""
Script cập nhật dữ liệu dây điện cho ứng dụng 80SERV trực tiếp từ link chia sẻ Microsoft OneDrive (.xlsx).
Cách dùng:
    python update_from_onedrive.py "https://1drv.ms/x/c/..."
hoặc chỉ cần chạy:
    python update_from_onedrive.py
rồi dán link khi được hỏi.
"""

import sys
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET
import urllib.request
import urllib.parse
import http.cookiejar

def download_onedrive_file(share_url, output_path="checklist_finall.xlsx"):
    share_url = share_url.strip()
    if not share_url:
        raise ValueError("Đường link OneDrive không được để trống.")
    
    # Chuẩn hoá URL tải về: đảm bảo có tham số download=1
    if "download=1" not in share_url:
        delim = "&" if "?" in share_url else "?"
        download_url = share_url + delim + "download=1"
    else:
        download_url = share_url

    print(f"[*] Đang kết nối tới OneDrive: {share_url[:70]}...")

    # Sử dụng CookieJar để xử lý redirect xác thực của OneDrive
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    opener.addheaders = [
        ('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    ]

    try:
        resp = opener.open(download_url, timeout=30)
        data = resp.read()
    except Exception as e:
        raise RuntimeError(f"Lỗi khi tải file từ OneDrive: {e}")

    # Kiểm tra xem dữ liệu tải về có phải là file ZIP/XLSX hợp lệ không (magic bytes PK\x03\x04)
    if not data.startswith(b'PK\x03\x04'):
        raise RuntimeError("Dữ liệu tải về không phải file Excel (.xlsx) hợp lệ. Vui lòng kiểm tra quyền chia sẻ của link OneDrive (chọn 'Bất kỳ ai có liên kết đều có thể xem').")

    with open(output_path, "wb") as f:
        f.write(data)

    print(f"[+] Tải file thành công: {output_path} ({len(data):,} bytes)")
    return output_path

def parse_xlsx_cables(xlsx_path):
    print(f"[*] Đang phân tích file Excel: {xlsx_path}...")
    with zipfile.ZipFile(xlsx_path) as z:
        wb_xml = z.read('xl/workbook.xml')
        wb_root = ET.fromstring(wb_xml)
        
        sheets = [s.attrib['name'] for s in wb_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')]
        
        sheets_list = []
        counts = {}
        total_obj = {'name': 'TỔNG', 'gid': '0'}
        all_cables = {}
        
        for i, name in enumerate(sheets):
            fname = f'xl/worksheets/sheet{i+1}.xml'
            try:
                s_xml = z.read(fname)
            except KeyError:
                continue
            s_root = ET.fromstring(s_xml)
            rows = s_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
            cables = []
            for r in rows:
                cells = []
                for c in r.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    t = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                    v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    val = t.text if t is not None else (v.text if v is not None else '')
                    cells.append(val)
                # Bỏ hàng tiêu đề hoặc hàng rỗng
                if cells and cells[0].strip():
                    first_cell = cells[0].strip().lower()
                    if not first_cell.startswith('cable') and not first_cell.startswith('tổng') and not first_cell.startswith('số dây'):
                        c6 = (cells + ['']*6)[:6]
                        cables.append(c6)
                    
            all_cables[name] = cables
            if name.lower() in ('total', 'tổng'):
                total_obj = {'name': name, 'gid': '0'}
            else:
                sheets_list.append({'name': name, 'gid': str(i)})
                counts[name] = len(cables)

    print(f"[+] Đã đọc: {len(sheets_list)} sheet tủ điện, tổng số dây: {len(all_cables.get('TỔNG', [])):,}")
    return total_obj, sheets_list, counts, all_cables

def update_index_html(total_obj, sheets_list, counts, all_cables, html_path="index.html"):
    print(f"[*] Đang cập nhật vào {html_path}...")
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    new_seed = {
        'total': total_obj,
        'sheets': sheets_list,
        'counts': counts
    }
    seed_json = json.dumps(new_seed, ensure_ascii=False, separators=(',', ':'))
    cables_json = json.dumps(all_cables, ensure_ascii=False, separators=(',', ':'))

    # Thay thế WIRE_SEED và WIRE_SEED_CABLES trong index.html
    seed_pattern = r'const WIRE_SEED=\{.*?\};'
    if not re.search(seed_pattern, content):
        raise RuntimeError("Không tìm thấy const WIRE_SEED trong index.html")
    content = re.sub(seed_pattern, f'const WIRE_SEED={seed_json};', content, count=1)

    cables_pattern = r'const WIRE_SEED_CABLES=\{.*?\};'
    if re.search(cables_pattern, content):
        content = re.sub(cables_pattern, f'const WIRE_SEED_CABLES={cables_json};', content, count=1)
    else:
        # Nếu chưa có WIRE_SEED_CABLES, chèn ngay sau WIRE_SEED
        content = content.replace(f'const WIRE_SEED={seed_json};', f'const WIRE_SEED={seed_json};\nconst WIRE_SEED_CABLES={cables_json};')

    with open(html_path, "w", encoding="utf-8", newline="") as f:
        f.write(content)

    print(f"[✔] HOÀN TẤT! Đã cập nhật thành công toàn bộ dữ liệu vào {html_path}!")

DEFAULT_ONEDRIVE_URL = "https://1drv.ms/x/c/e898f0bace036960/IQD2rR5GS2FjTYh2peM3EgROAY2D9geLBA0U7pVrTl7Nqmg?e=NuDBrO"

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].strip():
        onedrive_url = sys.argv[1].strip()
    else:
        onedrive_url = DEFAULT_ONEDRIVE_URL
        print(f"[*] Dùng link OneDrive mặc định: {DEFAULT_ONEDRIVE_URL}")

    if not onedrive_url:
        print("Lỗi: Chưa nhập link.")
        sys.exit(1)

    try:
        xlsx_file = download_onedrive_file(onedrive_url)
        total_obj, sheets_list, counts, all_cables = parse_xlsx_cables(xlsx_file)
        update_index_html(total_obj, sheets_list, counts, all_cables)
    except Exception as err:
        print(f"\n[X] LỖI: {err}")
        sys.exit(1)
