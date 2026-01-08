from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
import pandas as pd
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import io
import re

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def extract_all_awb_candidates(text):
    """Ekstrak semua kandidat AWB dari teks halaman PDF."""
    candidates = []
    patterns = [
        r'\bSPXID\d{12,15}\b',          # Shopee Express AWB
        r'\bSPX[A-Z]{2}\d{12,}\b',      # SPX variant
        r'\b\d{6}[A-Z0-9]{6,10}\b',     # Shopee Order ID: 260108ENX0N8NS
        r'\bJT\d{12,15}\b',             # J&T Express
        r'\bJX\d{12,15}\b',             # J&T Cargo
        r'\bTKP\d{10,15}\b',            # AnterAja/Tokopedia
        r'\bID\d{12,18}\b',             # Lazada LEX
        r'\bLEX\d{10,15}\b',            # LEX
        r'\b[A-Z]{2,4}\d{12,18}\b',     # General format
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        candidates.extend([m.upper() for m in matches])
    
    seen = set()
    unique = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            unique.append(c)
    return unique


def find_matching_awb(candidates, known_awbs):
    """Cari AWB yang match dengan Excel."""
    for candidate in candidates:
        if candidate in known_awbs:
            return candidate
        for known in known_awbs:
            if candidate in known or known in candidate:
                return known
    return None


def find_clear_start_position(page):
    """
    Cari posisi Y untuk memulai area putih (clear area).
    Logika:
    1. Cari teks "TANPA VIDEO UNBOXING, KOMPLIEN TIDAK DITERIMA"
    2. Jika ada "Catatan Pembeli:" di bawahnya, start dari bawah Catatan Pembeli
    3. Jika tidak ada Catatan Pembeli, start dari bawah teks TANPA VIDEO...
    4. Fallback ke y=300 jika tidak ditemukan
    """
    text_instances = page.get_text("dict")
    blocks = text_instances.get("blocks", [])
    
    unboxing_y = None
    catatan_y = None
    
    # Cari posisi teks
    for block in blocks:
        if block.get("type") == 0:  # Text block
            for line in block.get("lines", []):
                line_text = ""
                line_bottom = 0
                for span in line.get("spans", []):
                    line_text += span.get("text", "")
                    bbox = span.get("bbox", [0, 0, 0, 0])
                    line_bottom = max(line_bottom, bbox[3])  # y1 (bottom)
                
                # Cek teks TANPA VIDEO UNBOXING
                if "TANPA VIDEO UNBOXING" in line_text.upper():
                    unboxing_y = line_bottom
                    print(f"[DEBUG] Found 'TANPA VIDEO UNBOXING' at y={unboxing_y}")
                
                # Cek teks Catatan Pembeli (harus setelah TANPA VIDEO)
                if "CATATAN PEMBELI" in line_text.upper() and unboxing_y is not None:
                    # Cari sampai akhir catatan pembeli (bisa multi-line)
                    catatan_y = line_bottom
                    print(f"[DEBUG] Found 'Catatan Pembeli' at y={catatan_y}")
    
    # Tentukan posisi start clear area
    if catatan_y is not None:
        # Ada Catatan Pembeli, mulai dari bawahnya + margin
        clear_start = catatan_y + 5
        print(f"[DEBUG] Clear start after Catatan Pembeli: y={clear_start}")
    elif unboxing_y is not None:
        # Tidak ada Catatan Pembeli, mulai dari bawah TANPA VIDEO + margin
        clear_start = unboxing_y + 5
        print(f"[DEBUG] Clear start after TANPA VIDEO: y={clear_start}")
    else:
        # Fallback
        clear_start = 300
        print(f"[DEBUG] Fallback clear start: y={clear_start}")
    
    return clear_start


def calculate_available_rows(page_height, clear_start_y, row_height=18, margin_bottom=30):
    """
    Hitung berapa baris yang muat di area yang tersedia.
    """
    available_height = page_height - clear_start_y - margin_bottom
    # Kurangi 1 row untuk header
    max_rows = int(available_height / row_height) - 1
    return max(1, max_rows)  # Minimal 1 row


def create_table(data):
    """Buat tabel ReportLab untuk MSKU dan Qty"""
    t = Table(data, colWidths=[220, 50], rowHeights=18)
    
    style = [
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
    ]
    
    t.setStyle(TableStyle(style))
    return t


@app.post("/process-labels")
async def process_labels(
    excel_file: UploadFile = File(...),
    pdf_files: list[UploadFile] = File(...)
):
    try:
        # 1. Baca File Excel Ginee
        excel_content = await excel_file.read()
        df = pd.read_excel(io.BytesIO(excel_content))
        
        print(f"[DEBUG] Excel columns: {list(df.columns)}")
        
        # Normalize kolom
        col_mapping = {
            'AWB/No. Tracking': ['AWB/No. Tracking', 'AWB', 'No. Tracking', 'Tracking Number', 'Resi', 'No Resi'],
            'MSKU': ['MSKU', 'SKU', 'Nama SKU', 'Product SKU', 'Master SKU'],
            'Jumlah': ['Jumlah', 'Qty', 'Quantity', 'QTY']
        }
        
        for target_col, alternatives in col_mapping.items():
            if target_col not in df.columns:
                for alt in alternatives:
                    if alt in df.columns:
                        df = df.rename(columns={alt: target_col})
                        break
        
        required_cols = ['AWB/No. Tracking', 'MSKU', 'Jumlah']
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise HTTPException(status_code=400, detail=f"Kolom tidak ditemukan: {missing}")

        # 2. Buat mapping AWB ke items
        awb_to_items = {}
        all_awbs = set()
        
        for _, row in df.iterrows():
            awb = str(row['AWB/No. Tracking']).strip().upper()
            if awb and awb not in ['NAN', 'NONE', 'NULL', '']:
                all_awbs.add(awb)
                if awb not in awb_to_items:
                    awb_to_items[awb] = []
                awb_to_items[awb].append({
                    'msku': str(row['MSKU']),
                    'jumlah': int(row['Jumlah']) if pd.notna(row['Jumlah']) else 1
                })
        
        for awb in awb_to_items:
            awb_to_items[awb].sort(key=lambda x: x['msku'])
        
        print(f"[DEBUG] Unique AWBs: {len(all_awbs)}")

        # 3. Baca PDF
        all_pages = []
        for pdf in pdf_files:
            pdf_content = await pdf.read()
            doc = fitz.open(stream=pdf_content, filetype="pdf")
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                text = page.get_text()
                candidates = extract_all_awb_candidates(text)
                matched_awb = find_matching_awb(candidates, all_awbs)
                
                # Cari posisi clear area untuk halaman ini
                clear_start_y = find_clear_start_position(page)
                
                all_pages.append({
                    'page_num': page_num,
                    'page': page,
                    'doc': doc,
                    'width': page.rect.width,
                    'height': page.rect.height,
                    'awb': matched_awb,
                    'clear_start_y': clear_start_y
                })
                
                print(f"[DEBUG] Page {page_num + 1}: AWB={matched_awb}, clear_y={clear_start_y}")

        master_doc = fitz.open()
        matched_count = 0

        # 4. Proses setiap halaman
        for page_info in all_pages:
            page_awb = page_info['awb']
            page_num = page_info['page_num']
            W_pts = page_info['width']
            H_pts = page_info['height']
            clear_start_y = page_info['clear_start_y']
            
            if not page_awb or page_awb not in awb_to_items:
                # Copy tanpa modifikasi
                temp_doc = fitz.open()
                temp_doc.insert_pdf(page_info['doc'], from_page=page_num, to_page=page_num)
                master_doc.insert_pdf(temp_doc)
                temp_doc.close()
                continue
            
            matched_count += 1
            items = awb_to_items[page_awb]
            
            # Hitung berapa row yang muat di halaman pertama
            available_rows = calculate_available_rows(H_pts, clear_start_y)
            print(f"[DEBUG] Page {page_num + 1}: {len(items)} items, {available_rows} rows available")
            
            # Pagination dinamis
            limit_extra = 25
            chunks = []
            if len(items) <= available_rows:
                chunks.append(items)
            else:
                chunks.append(items[:available_rows])
                remaining = items[available_rows:]
                for i in range(0, len(remaining), limit_extra):
                    chunks.append(remaining[i:i + limit_extra])

            # Generate halaman
            for i, chunk in enumerate(chunks):
                if i == 0:
                    # Halaman 1: Overlay
                    temp_doc = fitz.open()
                    temp_doc.insert_pdf(page_info['doc'], from_page=page_num, to_page=page_num)
                    page_copy = temp_doc[0]
                    
                    # Clear area dimulai dari posisi dinamis
                    rect_clear = fitz.Rect(0, clear_start_y, W_pts, H_pts)
                    page_copy.draw_rect(rect_clear, color=(1, 1, 1), fill=(1, 1, 1))
                    
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=(W_pts, H_pts))
                    
                    table_data = [['MSKU', 'Qty']]
                    for item in chunk:
                        table_data.append([item['msku'], str(item['jumlah'])])

                    t = create_table(table_data)
                    t.wrapOn(can, W_pts, H_pts)
                    # Posisi tabel: di bawah clear area
                    table_y = H_pts - clear_start_y - (len(chunk) + 1) * 18 - 10
                    t.drawOn(can, 7, max(10, table_y))
                    can.save()
                    
                    packet.seek(0)
                    overlay_pdf = fitz.open("pdf", packet.read())
                    page_copy.show_pdf_page(page_copy.rect, overlay_pdf, 0)
                    master_doc.insert_pdf(temp_doc, from_page=0, to_page=0)
                    overlay_pdf.close()
                    temp_doc.close()
                else:
                    # Halaman lanjutan
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=(W_pts, H_pts))
                    
                    can.setFont("Helvetica-Bold", 10)
                    can.drawString(10, H_pts - 25, f"Lanjutan AWB: {page_awb}")
                    
                    table_data = [['MSKU', 'Qty']]
                    for item in chunk:
                        table_data.append([item['msku'], str(item['jumlah'])])

                    t = create_table(table_data)
                    t.wrapOn(can, W_pts, H_pts)
                    t_height = (len(chunk) + 1) * 18
                    t.drawOn(can, 7, H_pts - t_height - 45)
                    can.save()
                    
                    packet.seek(0)
                    new_page_doc = fitz.open("pdf", packet.read())
                    master_doc.insert_pdf(new_page_doc)
                    new_page_doc.close()

        print(f"\n[RESULT] Matched: {matched_count}")

        if len(master_doc) == 0:
            raise HTTPException(status_code=404, detail="Tidak ada halaman yang diproses.")

        output_stream = io.BytesIO()
        master_doc.save(output_stream)
        master_doc.close()
        output_stream.seek(0)

        return StreamingResponse(
            output_stream,
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=SIAP_PRINT_ALL.pdf"}
        )

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/health")
async def health_check():
    return {"status": "ok", "message": "Backend is running"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
