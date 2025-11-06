import os
import pytesseract
from PIL import Image, ImageTk
import customtkinter as ctk
from tkinter import filedialog
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import cv2
import numpy as np

# -------------------- KONFIGURASI TESSERACT --------------------
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# -------------------- INISIALISASI GUI --------------------
app = ctk.CTk()
app.title("📄 DocuVision AI OCR - Enhanced Layout")
app.geometry("780x640")
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

selected_image = None

# -------------------- FUNGSI ENHANCED --------------------
def preprocess_image(image_path):
    """Preprocess image untuk meningkatkan kualitas OCR"""
    img = cv2.imread(image_path)
    
    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Noise reduction
    denoised = cv2.medianBlur(gray, 3)
    
    # Thresholding untuk meningkatkan kontras
    _, thresh = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    return thresh

def process_ocr_enhanced():
    global selected_image
    if not selected_image:
        status_label.configure(text="⚠️ Harap pilih gambar terlebih dahulu!", text_color="red")
        return

    try:
        status_label.configure(text="🤖 Memproses dengan AI OCR (Enhanced Layout)...", text_color="blue")
        app.update()

        # Preprocess image
        processed_image = preprocess_image(selected_image)
        pil_image = Image.fromarray(processed_image)

        # OCR dengan config yang lebih baik untuk preservasi layout
        custom_config = r'--psm 6 -c preserve_interword_spaces=1'
        
        # Gunakan image_to_data untuk mendapatkan posisi teks
        data = pytesseract.image_to_data(pil_image, output_type=pytesseract.Output.DICT, config=custom_config)
        
        # Juga gunakan image_to_string dengan hOCR untuk layout yang lebih baik
        hocr_data = pytesseract.image_to_pdf_or_hocr(pil_image, extension='hocr', config=custom_config)

        # ---------------- DOCX DENGAN LAYOUT LEBIH BAIK ----------------
        doc = Document()
        
        # Group text by lines untuk preservasi layout horizontal
        lines = {}
        for i in range(len(data['text'])):
            if int(data['conf'][i]) > 30:  # Hanya teks dengan confidence > 30%
                text = data['text'][i].strip()
                if text:
                    top = data['top'][i]
                    left = data['left'][i]
                    
                    # Group by line (menggunakan top coordinate)
                    line_key = top // 10  # Group pixels dalam range 10px
                    if line_key not in lines:
                        lines[line_key] = []
                    
                    lines[line_key].append((left, text))
        
        # Sort lines by vertical position dan text oleh horizontal position
        for line_key in sorted(lines.keys()):
            line_texts = sorted(lines[line_key], key=lambda x: x[0])
            line_text = ' '.join([text for _, text in line_texts])
            if line_text.strip():
                doc.add_paragraph(line_text.strip())

        docx_path = "hasil_ai_enhanced.docx"
        doc.save(docx_path)

        # ---------------- PDF DENGAN LAYOUT LEBIH AKURAT ----------------
        pdf_path = "hasil_ai_enhanced.pdf"
        c = canvas.Canvas(pdf_path, pagesize=A4)
        width, height = A4
        
        # Scale factor yang lebih baik
        img_pil = Image.open(selected_image)
        img_width, img_height = img_pil.size
        
        scale_x = width / img_width * 0.9
        scale_y = height / img_height * 0.9
        scale = min(scale_x, scale_y)
        
        # Tambahkan gambar asli sebagai background (transparan)
        c.drawImage(selected_image, 50, height - (img_height * scale) - 50, 
                   width=img_width * scale, height=img_height * scale)
        
        # Tambahkan teks di posisi yang sesuai
        for i in range(len(data['text'])):
            if int(data['conf'][i]) > 30:
                text = data['text'][i].strip()
                if text:
                    x = 50 + (data['left'][i] * scale)
                    y = height - 50 - (data['top'][i] * scale)
                    
                    # Set font size yang sesuai
                    font_size = max(8, data['height'][i] * scale * 0.7)
                    c.setFont("Helvetica", font_size)
                    c.setFillColorRGB(0, 0, 0)  # Warna hitam
                    
                    c.drawString(x, y, text)
        
        c.save()

        status_label.configure(
            text=f"✅ OCR selesai dengan preservasi layout!\nFile tersimpan:\n📄 {docx_path}\n📘 {pdf_path}",
            text_color="green"
        )

    except Exception as e:
        status_label.configure(text=f"❌ Terjadi kesalahan: {e}", text_color="red")

def upload_image():
    global selected_image
    file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.jpg *.jpeg *.png *.bmp *.tiff")])
    if file_path:
        selected_image = file_path
        img = Image.open(file_path)
        img.thumbnail((450, 450))
        tk_img = ImageTk.PhotoImage(img)
        preview_label.configure(image=tk_img, text="")
        preview_label.image = tk_img
        status_label.configure(text="📷 Gambar dimuat, siap diproses.", text_color="green")

# -------------------- LAYOUT GUI DENGAN TOMBOL BARU --------------------
sidebar = ctk.CTkFrame(app, width=200, corner_radius=0)
sidebar.pack(side="left", fill="y")

title_label = ctk.CTkLabel(sidebar, text="📄 DocuVision AI Pro", font=("Segoe UI", 20, "bold"))
title_label.pack(pady=20)

upload_btn = ctk.CTkButton(sidebar, text="📁 Pilih Gambar", command=upload_image)
upload_btn.pack(pady=10)

process_btn = ctk.CTkButton(sidebar, text="🚀 OCR Enhanced", fg_color="#0078D7",
                            hover_color="#005EA6", command=process_ocr_enhanced)
process_btn.pack(pady=10)

footer = ctk.CTkLabel(sidebar, text="© 2025 Kusuma-pixel", text_color="gray", font=("Arial", 10))
footer.pack(side="bottom", pady=10)

content_frame = ctk.CTkFrame(app)
content_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

preview_label = ctk.CTkLabel(content_frame, text="(Pratinjau Gambar di sini)", text_color="gray")
preview_label.pack(pady=20)

status_label = ctk.CTkLabel(content_frame, text="", font=("Arial", 12))
status_label.pack(pady=10)

app.mainloop()