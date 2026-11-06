import os
import shutil
from pathlib import Path
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pypdf import PdfReader
from docx import Document
import openpyxl
from PIL import Image as PILImage
import pytesseract

# دالة استخراج النص من مختلف أنواع الملفات
def extract_text(file_path):
    ext = file_path.suffix.lower()
    text = ""
    try:
        if ext == '.pdf':
            reader = PdfReader(file_path)
            for page in reader.pages:
                text += page.extract_text() or ""
        elif ext in ['.docx', '.doc']:
            doc = Document(file_path)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif ext in ['.xlsx', '.xls']:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    row_str = " ".join([str(cell) for cell in row if cell is not None])
                    text += row_str + "\n"
        elif ext in ['.png', '.jpg', '.jpeg']:
            image = PILImage.open(file_path)
            text = pytesseract.image_to_string(image, lang='ara+eng')
    except Exception as e:
        pass
    
    return (file_path.name + " " + text).lower()

# دالة التصنيف بالقواعد البرمجية (Rule-Based Classification)
def classify_by_rules(content_text):
    if any(word in content_text for word in ["فاتورة", "invoice", "receipt", "ايصال", "حساب"]):
        return "المالية/الفواتير"
    elif any(word in content_text for word in ["راتب", "kashf", "salary", "مفردات مرتب", "كشف حساب"]):
        return "المالية/الرواتب"
    elif any(word in content_text for word in ["عقد", "agreement", "contract", "اتفاقية"]):
        return "العمل/العقود"
    elif any(word in content_text for word in ["شهادة", "certificate", "تخرج", "دبلوم"]):
        return "الشخصية/الشهادات"
    elif any(word in content_text for word in ["تقرير", "report", "دراسة", "بحث"]):
        return "العمل/التقارير"
    elif any(word in content_text for word in ["هوية", "passport", "جواز", "بطاقة", "إقامة"]):
        return "الشخصية/الهويات_والوثائق"
    else:
        return "أخرى/ملفات_غير_مفروزة"

class DocumentOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("منظم المستندات الذكي (مع ميزة التشغيل التجريبي)")
        self.root.geometry("650x600")
        self.root.config(bg="#f0f0f0")
        
        self.source_dir = tk.StringVar()
        self.target_dir = tk.StringVar()
        self.dry_run_var = tk.BooleanVar(value=True) # مفعلة افتراضياً للأمان

        self.create_widgets()

    def create_widgets(self):
        # عنوان رئيسي
        title_label = tk.Label(self.root, text="منظم الوثائق والملفات (تشغيل تجريبي / حقيقي)", font=("Arial", 14, "bold"), bg="#f0f0f0", fg="#333")
        title_label.pack(pady=15)

        # مجلد المصدر
        src_frame = tk.Frame(self.root, bg="#f0f0f0")
        src_frame.pack(fill="x", padx=20, pady=5)
        
        tk.Label(src_frame, text="مجلد الملفات:", width=18, anchor="e", bg="#f0f0f0").pack(side="right", padx=5)
        tk.Entry(src_frame, textvariable=self.source_dir, width=45).pack(side="right", padx=5)
        tk.Button(src_frame, text="استعراض...", command=self.browse_source, bg="#ddd").pack(side="right")

        # مجلد الوجهة
        tgt_frame = tk.Frame(self.root, bg="#f0f0f0")
        tgt_frame.pack(fill="x", padx=20, pady=5)
        
        tk.Label(tgt_frame, text="مجلد الحفظ:", width=18, anchor="e", bg="#f0f0f0").pack(side="right", padx=5)
        tk.Entry(tgt_frame, textvariable=self.target_dir, width=45).pack(side="right", padx=5)
        tk.Button(tgt_frame, text="استعراض...", command=self.browse_target, bg="#ddd").pack(side="right")

        # إطار خيار Dry Run
        opt_frame = tk.Frame(self.root, bg="#f0f0f0")
        opt_frame.pack(fill="x", padx=20, pady=8)
        
        chk_dry = tk.Checkbutton(opt_frame, text="تشغيل تجريبي (Dry Run) - محاكاة بدون نقل حقيقي للملفات", variable=self.dry_run_var, bg="#f0f0f0", font=("Arial", 10, "bold"), fg="#d32f2f")
        chk_dry.pack(side="right", padx=5)

        # زر بدء التنظيم
        self.start_btn = tk.Button(self.root, text="ابدأ العملية", command=self.start_thread, bg="#2196F3", fg="white", font=("Arial", 11, "bold"), width=25, height=2)
        self.start_btn.pack(pady=15)

        # صندوق السجل (Log)
        self.log_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, width=75, height=13, font=("Courier", 9))
        self.log_area.pack(padx=20, pady=5)
        self.log_area.insert(tk.END, "جاهز للعمل... (الوضع التجريبي مفعل تلقائياً للأمان)\n")

    def browse_source(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.source_dir.set(dir_path)

    def browse_target(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.target_dir.set(dir_path)

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def start_thread(self):
        src = self.source_dir.get()
        tgt = self.target_dir.get()
        is_dry_run = self.dry_run_var.get()
        
        if not src or not tgt:
            messagebox.showerror("خطأ", "الرجاء تحديد مجلد المصدر ومجلد الوجهة أولاً!")
            return
            
        self.start_btn.config(state=tk.DISABLED, bg="#9e9e9e")
        threading.Thread(target=self.run_organization, args=(src, tgt, is_dry_run), daemon=True).start()

    def run_organization(self, source_dir, target_base_dir, is_dry_run):
        mode_text = "[وضع تجريبي - Dry Run]" if is_dry_run else "[وضع حقيقي - نقل فعلي]"
        self.log(f"--- بدأ العملية {mode_text} ---")
        source_path = Path(source_dir)
        base_target = Path(target_base_dir)
        
        supported_extensions = ['.pdf', '.docx', '.doc', '.xlsx', '.xls', '.png', '.jpg', '.jpeg']
        
        count = 0
        for file_path in source_path.glob('**/*.*'):
            if file_path.suffix.lower() in supported_extensions:
                self.log(f"فحص: {file_path.name}")
                
                content_text = extract_text(file_path)
                category = classify_by_rules(content_text)
                
                destination_folder = base_target / category
                target_file_path = destination_folder / file_path.name
                
                if is_dry_run:
                    # محاكاة فقط دون نقل أو إنشاء مجلدات حقيقية
                    self.log(f"   (محاكاة) سيتم نقله إلى $\leftarrow$ {category}")
                    count += 1
                else:
                    # تنفيذ حقيقي
                    destination_folder.mkdir(parents=True, exist_ok=True)
                    try:
                        shutil.move(str(file_path), str(target_file_path))
                        self.log(f"   تم النقل بنجاح إلى $\leftarrow$ {category}")
                        count += 1
                    except Exception as e:
                        self.log(f"   خطأ أثناء النقل: {e}")
                    
        self.log(f"--- انتهت العملية! تم معالجة {count} ملف بنجاح. ---")
        self.start_btn.config(state=tk.NORMAL, bg="#2196F3")
        
        msg = f"تم الانتهاء من محاكاة تنظيم {count} ملف بنجاح!" if is_dry_run else f"تم الانتهاء من تنظيم ونقل {count} ملف بنجاح!"
        messagebox.showinfo("تم", msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentOrganizerApp(root)
    root.mainloop()
