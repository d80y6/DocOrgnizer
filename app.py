import os
import shutil
import json
from pathlib import Path
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pypdf import PdfReader
from docx import Document
import openpyxl
from PIL import Image as PILImage
import pytesseract
from google import genai

CONFIG_FILE = "settings.json"

def load_settings():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {"api_key": "", "source_dir": "", "target_dir": ""}

def save_settings(settings_data):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings_data, f, ensure_ascii=False, indent+4 if 'indent+4' in globals() else 4)
    except:
        pass

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
    
    return (file_path.name + "\n" + text)[:3000]

# التصنيف الذكي الحر (Gemini يقترح التصنيف من تلقاء نفسه بالكامل)
def classify_with_gemini_fully_dynamic(content_text, api_key):
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
        أنت مساعد ذكي لتنظيم الملفات والوثائق. مهمتك قراءة النص أو اسم الملف أدناه، واستنتاج واقتراح مسار مجلد رئيسي ومجلد فرعي مناسب جداً من تفكيرك الخاص (بالصيغة التالية: المجلد_الرئيسي/المجلد_الفرعي).
        
        شروط هامة:
        1. لا تتقيد بأي قوائم قديمة، ابتكر تصنيفات دقيقة واحترافية باللغة العربية (أو الإنجليزية إذا كان الملف تقنياً بحتاً وتفضّل ذلك)، مثل: (الأمن_السيبراني/التحقيق_الجنائى)، (الشبكات/تقارير_DNS)، (الكتب_والأبحاث/تقنية)، (الصور/لقطات_شاشة)، (المالية/الفواتير)، وغيرها بحسب طبيعة المحتوى.
        2. استخدم الشرطة السفلية (_) بدلاً من المسافات في أسماء المجلدات لكي تكون متوافقة مع أنظمة الملفات.
        3. أجب بصيغة المسار فقط (المجلد_الرئيسي/المجلد_الفرعي) بدون أي مقدمات، رموز تزيينية، أو شرح إضافي نهائياً.

        محتوى المستند أو اسم الملف:
        {content_text}
        """
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        # تنظيف النتيجة لضمان الحصول على مسار نظيف فقط
        result = response.text.strip().replace("'", "").replace('"', "").replace("`", "")
        # التأكد من أنه مسار صالح
        if "/" not in result and "\\" not in result:
            return "مواضيع_عامة/" + result
        return result
    except Exception as e:
        # ⚠️ هنا التعديل: سنقوم بإرجاع وصف الخطأ أو طباعته لنعرف السبب الحقيقي
        print(f"Gemini API Error: {e}")
        return "أخرى/ملفات_غير_مفروزة"

class DocumentOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("منظم المستندات الذكي (استنتاج حر للتصنيفات عبر Gemini)")
        self.root.geometry("700x650")
        self.root.config(bg="#f0f0f0")
        
        self.settings = load_settings()

        self.source_dir = tk.StringVar(value=self.settings.get("source_dir", ""))
        self.target_dir = tk.StringVar(value=self.settings.get("target_dir", ""))
        self.api_key_var = tk.StringVar(value=self.settings.get("api_key", ""))
        self.dry_run_var = tk.BooleanVar(value=True)
        
        self.is_running = False

        self.create_widgets()

    def create_widgets(self):
        title_label = tk.Label(self.root, text="منظم الوثائق (استنتاج حر للتصنيفات بالذكاء الاصطناعي)", font=("Arial", 13, "bold"), bg="#f0f0f0", fg="#1a73e8")
        title_label.pack(pady=12)

        key_frame = tk.Frame(self.root, bg="#f0f0f0")
        key_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(key_frame, text="Gemini API Key:", width=18, anchor="e", bg="#f0f0f0", font=("Arial", 9, "bold")).pack(side="right", padx=5)
        tk.Entry(key_frame, textvariable=self.api_key_var, width=48, show="*").pack(side="right", padx=5)

        src_frame = tk.Frame(self.root, bg="#f0f0f0")
        src_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(src_frame, text="مجلد الملفات:", width=18, anchor="e", bg="#f0f0f0").pack(side="right", padx=5)
        tk.Entry(src_frame, textvariable=self.source_dir, width=48).pack(side="right", padx=5)
        tk.Button(src_frame, text="استعراض...", command=self.browse_source, bg="#ddd").pack(side="right")

        tgt_frame = tk.Frame(self.root, bg="#f0f0f0")
        tgt_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(tgt_frame, text="مجلد الحفظ:", width=18, anchor="e", bg="#f0f0f0").pack(side="right", padx=5)
        tk.Entry(tgt_frame, textvariable=self.target_dir, width=48).pack(side="right", padx=5)
        tk.Button(tgt_frame, text="استعراض...", command=self.browse_target, bg="#ddd").pack(side="right")

        opt_frame = tk.Frame(self.root, bg="#f0f0f0")
        opt_frame.pack(fill="x", padx=20, pady=8)
        chk_dry = tk.Checkbutton(opt_frame, text="تشغيل تجريبي (Dry Run) - معاينة الهيكل المستنتج", variable=self.dry_run_var, bg="#f0f0f0", font=("Arial", 10, "bold"), fg="#d32f2f")
        chk_dry.pack(side="right", padx=5)

        btn_frame = tk.Frame(self.root, bg="#f0f0f0")
        btn_frame.pack(pady=10)

        self.start_btn = tk.Button(btn_frame, text="ابدأ الاستنتاج والتنظيم الذكي", command=self.start_thread, bg="#1a73e8", fg="white", font=("Arial", 11, "bold"), width=22, height=2)
        self.start_btn.pack(side="right", padx=5)

        self.stop_btn = tk.Button(btn_frame, text="إيقاف العملية", command=self.stop_process, bg="#d32f2f", fg="white", font=("Arial", 11, "bold"), width=15, height=2, state=tk.DISABLED)
        self.stop_btn.pack(side="right", padx=5)

        self.log_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, width=82, height=12, font=("Courier", 9))
        self.log_area.pack(padx=20, pady=5)
        self.log_area.insert(tk.END, "جاهز للعمل... سيقوم Gemini باستنتاج التصنيفات من تلقاء نفسه لكل ملف.\n")

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
        api_key = self.api_key_var.get().strip()
        is_dry_run = self.dry_run_var.get()
        
        if not src or not tgt:
            messagebox.showerror("خطأ", "الرجاء تحديد مجلد المصدر ومجلد الوجهة!")
            return
        if not api_key:
            messagebox.showerror("خطأ", "الرجاء إدخال مفتاح Gemini API Key!")
            return
            
        save_settings({"api_key": api_key, "source_dir": src, "target_dir": tgt})

        self.is_running = True
        self.start_btn.config(state=tk.DISABLED, bg="#9e9e9e")
        self.stop_btn.config(state=tk.NORMAL)
        
        threading.Thread(target=self.run_organization, args=(src, tgt, api_key, is_dry_run), daemon=Thread if 'Thread' in globals() else daemon=True).start() if 'daemon' in locals() else threading.Thread(target=self.run_organization, args=(src, tgt, api_key, is_dry_run), daemon=True).start()

    def stop_process(self):
        self.is_running = False
        self.log("--- تم طلب إيقاف العملية من قبل المستخدم ---")

    def run_organization(self, source_dir, target_base_dir, api_key, is_dry_run):
        mode_text = "[وضع تجريبي - Dry Run]" if is_dry_run else "[وضع حقيقي]"
        self.log(f"--- بدأ الاستنتاج الذكي الحر عبر Gemini {mode_text} ---")
        
        source_path = Path(source_dir)
        base_target = Path(target_base_dir)
        supported_extensions = ['.pdf', '.docx', '.doc', '.xlsx', '.xls', '.png', '.jpg', '.jpeg']
        
        success_count = 0
        fail_count = 0
        categories_stats = {}

        for file_path in source_path.glob('**/*.*'):
            if not self.is_running:
                break

            if file_path.suffix.lower() in supported_extensions:
                self.log(f"قراءة وتحليل: {file_path.name}")
                
                content_text = extract_text(file_path)
                # استنتاج التصنيف بشكل حر كلياً بواسطة الذكاء الاصطناعي
                category = classify_with_gemini_fully_dynamic(content_text, api_key)
                
                self.log(f"   🧠 التصنيف المستنتج: {category}")
                
                destination_folder = base_target / category
                target_file_path = destination_folder / file_path.name
                
                if is_dry_run:
                    self.log(f"   (محاكاة) سيتم إنشاؤه ونقله إلى ← {category}")
                    success_count += 1
                    categories_stats[category] = categories_stats.get(category, 0) + 1
                else:
                    destination_folder.mkdir(parents=True, exist_ok=True)
                    try:
                        shutil.move(str(file_path), str(target_file_path))
                        self.log(f"   ✨ تم النقل بنجاح إلى ← {category}")
                        success_count += 1
                        categories_stats[category] = categories_stats.get(category, 0) + 1
                    except Exception as e:
                        fail_count += 1
                        self.log(f"   ❌ خطأ أثناء النقل: {e}")
                    
        self.log("\n--- 📊 تقرير الهيكل التنظيمي المستنتج ---")
        self.log(f"• إجمالي الملفات المعالجة: {success_count}")
        self.log(f"• الملفات التي فشلت: {fail_count}")
        self.log("• هيكل المجلدات التي تم استنتاجها:")
        for cat, count in categories_stats.items():
            self.log(f"   📁 {cat} ({count} ملفات)")
            
        self.start_btn.config(state=tk.NORMAL, bg="#1a73e8")
        self.stop_btn.config(state=tk.DISABLED)
        self.is_running = False
        
        messagebox.showinfo("تم", f"تم استنتأج الهيكل التنظيمي ومعالجة {success_count} ملف بنجاح!")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentOrganizerApp(root)
    root.mainloop()
