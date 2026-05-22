import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
from google import genai
import arabic_reshaper
from bidi.algorithm import get_display
import re, os, sys

# 🔴 مفتاح الـ API مدمج
GEMINI_API_KEY = "AIzaSyC0Aom0fJY41PFFVB0h_foP0NMG7HiLbWg"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class UltimateOCRApp:
    def __init__(self, root, on_confirm_cb=None):
        self.root = root
        self.on_confirm_cb = on_confirm_cb
        self.root.title("نظام قراءة وثائق رخص السياقة الجزائرية")
        self.root.geometry("860x720")
        self.root.configure(bg="#f5f6fa")
        self.root.minsize(700, 600)
        self._data = None

        try:
            self.client = genai.Client(api_key=GEMINI_API_KEY)
        except Exception:
            self.client = None

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root,
                 text="القارئ الآلي الذكي لرخص السياقة الجزائرية (عربي / Français)",
                 font=("Arial", 14, "bold"), bg="#f5f6fa", fg="#2f3640").pack(pady=15)

        btn_frame = tk.Frame(self.root, bg="#f5f6fa")
        btn_frame.pack(pady=10)

        self.btn_open = tk.Button(btn_frame,
                                   text="فتح صورة رخصة السياقة فوراً",
                                   font=("Arial", 11, "bold"), bg="#2ecc71", fg="white",
                                   bd=0, padx=20, pady=8, command=self.process_image)
        self.btn_open.pack(side=tk.RIGHT, padx=10)

        tk.Button(btn_frame, text="📂 PDF متعدد الرخص",
                  font=("Arial", 11, "bold"), bg="#2980b9", fg="white",
                  bd=0, padx=14, pady=8,
                  command=self.process_batch_pdf).pack(side=tk.RIGHT, padx=6)

        tk.Button(btn_frame, text="إغلاق البرنامج",
                  font=("Arial", 11, "bold"), bg="#e74c3c", fg="white",
                  bd=0, padx=20, pady=8, command=self.root.destroy).pack(side=tk.LEFT, padx=10)

        # ── الجزء السفلي الثابت ──────────────────────────────
        bottom_frame = tk.Frame(self.root, bg="#f5f6fa")
        bottom_frame.pack(side=tk.BOTTOM, fill="x", padx=20, pady=8)

        self.lbl_status = tk.Label(bottom_frame, text="",
                                    font=("Arial", 9), bg="#f5f6fa", fg="#e67e22")
        self.lbl_status.pack(pady=(0,4))

        self.btn_confirm = tk.Button(bottom_frame,
                                      text="✅ تأكيد وحفظ في قاعدة البيانات",
                                      font=("Arial", 12, "bold"),
                                      bg="#1a6b3c", fg="white",
                                      bd=0, padx=30, pady=10,
                                      state="disabled",
                                      command=self._confirm)
        self.btn_confirm.pack(fill="x")

        # ── منطقة النص ───────────────────────────────────────
        self.text_area = tk.Text(self.root, font=("Arial", 13, "bold"),
                                  wrap=tk.WORD, bd=1, relief=tk.SOLID,
                                  fg="#2f3640", bg="white", padx=20, pady=20)
        self.text_area.pack(padx=20, pady=(5,0), fill=tk.BOTH, expand=True)

    def fix_bilingual_text(self, text):
        """نفس دالة الكود الأصلي بالضبط"""
        lines = text.split('\n')
        fixed_lines = []
        for line in lines:
            if not line.strip():
                fixed_lines.append("")
                continue
            reshaped_line = arabic_reshaper.reshape(line)
            bidi_line = get_display(reshaped_line)
            fixed_lines.append(bidi_line)
        return '\n'.join(fixed_lines)

    def process_image(self):
        file_path = filedialog.askopenfilename(
            title="اختر صورة أو PDF رخصة السياقة",
            filetypes=[
                ("الصور و PDF", "*.jpg *.jpeg *.png *.bmp *.pdf"),
                ("صور", "*.jpg *.jpeg *.png *.bmp"),
                ("PDF", "*.pdf"),
            ]
        )
        if not file_path:
            return

        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, "جاري تحليل وقراءة بيانات الرخصة...")
        self.btn_confirm.config(state="disabled")
        self.lbl_status.config(text="")
        self.root.update()

        try:
            # ── تحويل PDF لصورة إذا لزم ──────────────────────
            if file_path.lower().endswith(".pdf"):
                try:
                    import fitz
                except ImportError:
                    import subprocess, sys
                    self.text_area.delete("1.0", tk.END)
                    self.text_area.insert(tk.END, "⏳ جاري تثبيت PyMuPDF...")
                    self.root.update()
                    installed = False
                    for args in [
                        [sys.executable, "-m", "pip", "install", "pymupdf", "-q"],
                        [sys.executable, "-m", "pip", "install", "pymupdf", "-q", "--user"],
                        [sys.executable, "-m", "pip", "install", "PyMuPDF", "-q"],
                    ]:
                        try:
                            subprocess.check_call(args, stdout=subprocess.DEVNULL,
                                                  stderr=subprocess.DEVNULL)
                            installed = True
                            break
                        except Exception:
                            continue
                    if not installed:
                        messagebox.showerror("خطأ",
                            "فشل تثبيت PyMuPDF\n"
                            "شغّل يدوياً في cmd:\n"
                            "pip install pymupdf")
                        return
                    try:
                        import fitz
                    except ImportError:
                        messagebox.showerror("خطأ", "لم يتم تثبيت PyMuPDF — أعد تشغيل البرنامج")
                        return
                try:
                    pdf_doc = fitz.open(file_path)
                    page    = pdf_doc[0]
                    mat     = fitz.Matrix(2.0, 2.0)   # دقة ×2
                    pix     = page.get_pixmap(matrix=mat)
                    from PIL import Image as _PIL_Image
                    import io as _io
                    img = _PIL_Image.open(_io.BytesIO(pix.tobytes("png")))
                    pdf_doc.close()
                except Exception as pdf_err:
                    messagebox.showerror("خطأ", f"خطأ في تحويل PDF:\n{pdf_err}")
                    return
            else:
                img = Image.open(file_path)

            # إرسال الصورة لـ Gemini مع retry تلقائي
            gemini_prompt = (
                "أنت خبير في قراءة رخص السياقة البيومترية الجزائرية. "
                "استخرج البيانات بالتسميات الرسمية التالية بالضبط:\n"
                "1. اللقب: الكلمة العربية - الكلمة الفرنسية\n"
                "2. الاسم: الكلمة العربية - الكلمة الفرنسية\n"
                "3. تاريخ ومكان الازدياد: DD/MM/YYYY - المدينة بالعربية\n"
                "4. رقم التعريف الوطني: الرقم الطويل 15-18 رقم\n"
                "5. رقم رخصة السياقة: الرقم الذي يبدأ بحرف L\n"
                "6. فئات الرخصة: الحروف اللاتينية فقط مثل B أو B,C — لا تضع الجنس هنا\n"
                "7. الجنس: M أو F فقط\n"
                "8. تاريخ صدور رخصة السياقة: DD/MM/YYYY\n"
                "9. تاريخ انتهاء رخصة السياقة: DD/MM/YYYY\n"
                "10. مكان إصدار الرخصة: اسم المدينة\n"
                "11. زمرة الدم: مثل A+ أو B-\n"
                "12. العنوان: عنوان السكن إن وجد\n"
                "تعليمات مهمة:\n"
                "- اللقب والاسم بالعربية: انسخهما حرفياً كما هما مكتوبان على الرخصة بالعربية بدون أي تعديل\n"
                "- إذا كانت الكتابة العربية غير واضحة، حوّل الاسم الفرنسي للعربية: BEL=بل، BEN=بن، BOU=بو، MOUN=مون، OUN=ون، MA=ما، MOU=مو\n"
                "- فئات الرخصة: لا تضع حرف الجنس (M أو F) ضمن الفئات\n"
                "أجب بأسطر مرقمة فقط بدون مقدمة أو خاتمة."
            )

            import time as _time
            response = None
            last_error = None
            for attempt in range(4):
                try:
                    response = self.client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[img, gemini_prompt]
                    )
                    break
                except Exception as e:
                    last_error = e
                    err_str = str(e)
                    if '503' in err_str or 'UNAVAILABLE' in err_str:
                        wait = 5 * (attempt + 1)
                        self.text_area.delete("1.0", tk.END)
                        self.text_area.insert(tk.END,
                            f"⏳ الخادم مشغول — انتظار {wait} ثانية...\n"
                            f"المحاولة {attempt+1}/4")
                        self.root.update()
                        _time.sleep(wait)
                    elif '429' in err_str or 'QUOTA' in err_str or 'EXHAUSTED' in err_str:
                        wait = 20
                        self.text_area.delete("1.0", tk.END)
                        self.text_area.insert(tk.END,
                            f"⏳ تجاوز الحد — انتظار {wait} ثانية...\n"
                            f"المحاولة {attempt+1}/4")
                        self.root.update()
                        _time.sleep(wait)
                    else:
                        raise
            else:
                raise last_error

            self.text_area.delete("1.0", tk.END)
            if response.text:
                raw = response.text.strip()
                final_text = self.fix_bilingual_text(raw)
                self.text_area.insert(tk.END, final_text)
                # تحليل البيانات
                self._data = parse_text_to_dict(raw)
                self.btn_confirm.config(state="normal")
                # فحص قاعدة البيانات
                nin = self._data.get("رقم التعريف الوطني", "")
                if nin:
                    action, info = check_in_database(nin)
                    if action == "found":
                        self.lbl_status.config(
                            text=f"✅ مترشح موجود: {info} — سيتم التحديث",
                            fg="#27ae60")
                    else:
                        self.lbl_status.config(
                            text="➕ مترشح جديد — سيتم طلب الإذن",
                            fg="#e67e22")
            else:
                self.text_area.insert(tk.END, "لم يتم العثور على بيانات.")

        except Exception as e:
            self.text_area.delete("1.0", tk.END)
            messagebox.showerror("خطأ", f"حدثت مشكلة: {str(e)}")

    def process_batch_pdf(self):
        """استيراد PDF يحتوي رخص متعددة — كل صفحة رخصة"""
        file_path = filedialog.askopenfilename(
            title="اختر PDF يحتوي رخص متعددة",
            filetypes=[("PDF", "*.pdf")]
        )
        if not file_path: return

        # تحقق من PyMuPDF
        try:
            import fitz
        except ImportError:
            import subprocess, sys
            self.text_area.delete("1.0", tk.END)
            self.text_area.insert(tk.END, "⏳ جاري تثبيت PyMuPDF...")
            self.root.update()
            installed = False
            for args in [
                [sys.executable,"-m","pip","install","pymupdf","-q"],
                [sys.executable,"-m","pip","install","pymupdf","-q","--user"],
            ]:
                try:
                    subprocess.check_call(args, stdout=subprocess.DEVNULL,
                                          stderr=subprocess.DEVNULL)
                    installed = True; break
                except Exception: continue
            if not installed:
                messagebox.showerror("خطأ","شغّل: pip install pymupdf"); return
            import fitz

        from PIL import Image as _PIL_Image
        import io as _io

        pdf_doc = fitz.open(file_path)
        total   = len(pdf_doc)

        if total == 0:
            messagebox.showwarning("⚠️", "الملف فارغ"); return

        if not messagebox.askyesno("تأكيد",
            f"الملف يحتوي {total} صفحة.\n"
            f"سيتم معالجة كل صفحة كرخصة منفصلة.\n\n"
            f"⚠️ الحد المجاني لـ Gemini: 20 طلب/يوم\n"
            f"هل تريد المتابعة؟"):
            return

        # اختيار نطاق الصفحات
        range_win = tk.Toplevel(self.root)
        range_win.title("اختر نطاق الصفحات")
        range_win.grab_set(); range_win.resizable(False, False)
        range_win.configure(bg="#1a1a2e")

        tk.Label(range_win, text=f"الملف يحتوي {total} صفحة",
                 bg="#1a1a2e", fg="white", font=("Arial",11,"bold")).pack(pady=(16,4))
        tk.Label(range_win, text="⚠️ الحد المجاني: 20 طلب/يوم",
                 bg="#1a1a2e", fg="#f39c12", font=("Arial",9)).pack(pady=(0,10))

        fr = tk.Frame(range_win, bg="#1a1a2e"); fr.pack(padx=20, pady=6)
        tk.Label(fr, text="من صفحة:", bg="#1a1a2e", fg="white",
                 font=("Arial",10)).grid(row=0, column=0, padx=6, pady=4, sticky="e")
        var_from = tk.IntVar(value=1)
        tk.Spinbox(fr, from_=1, to=total, textvariable=var_from,
                   width=6, font=("Arial",10)).grid(row=0, column=1, padx=6)

        tk.Label(fr, text="إلى صفحة:", bg="#1a1a2e", fg="white",
                 font=("Arial",10)).grid(row=1, column=0, padx=6, pady=4, sticky="e")
        var_to = tk.IntVar(value=min(total, 20))
        tk.Spinbox(fr, from_=1, to=total, textvariable=var_to,
                   width=6, font=("Arial",10)).grid(row=1, column=1, padx=6)

        confirmed = tk.BooleanVar(value=False)
        def _ok():
            confirmed.set(True); range_win.destroy()
        def _cancel():
            range_win.destroy()

        bf = tk.Frame(range_win, bg="#1a1a2e"); bf.pack(pady=12)
        tk.Button(bf, text="✅ ابدأ", bg="#2ecc71", fg="white",
                  font=("Arial",10,"bold"), padx=16, pady=6,
                  command=_ok).pack(side="right", padx=6)
        tk.Button(bf, text="إلغاء", bg="#e74c3c", fg="white",
                  font=("Arial",10,"bold"), padx=16, pady=6,
                  command=_cancel).pack(side="left", padx=6)

        range_win.wait_window()
        if not confirmed.get(): return

        page_from = max(1, var_from.get()) - 1        # 0-indexed
        page_to   = min(total, var_to.get())          # exclusive end
        pages_to_process = range(page_from, page_to)

        # نافذة التقدم
        prog_win = tk.Toplevel(self.root)
        prog_win.title("جاري الاستيراد..."); prog_win.grab_set()
        prog_win.geometry("400x200"); prog_win.configure(bg="#1a1a2e")
        prog_win.resizable(False, False)

        lbl_prog = tk.Label(prog_win, text="", bg="#1a1a2e", fg="white",
                             font=("Arial",11,"bold")); lbl_prog.pack(pady=20)

        import tkinter.ttk as ttk
        bar = ttk.Progressbar(prog_win, length=350, maximum=total)
        bar.pack(pady=8)

        lbl_res = tk.Label(prog_win, text="", bg="#1a1a2e", fg="#2ecc71",
                            font=("Arial",9)); lbl_res.pack(pady=4)

        results = {"ok": 0, "skip": 0, "err": 0}
        log = []
        mat = fitz.Matrix(2.0, 2.0)
        total_to_process = len(list(pages_to_process))
        pages_to_process = range(page_from, page_to)  # reset iterator

        import time as _time
        for idx, page_num in enumerate(pages_to_process):
            lbl_prog.config(text=f"معالجة الصفحة {page_num+1} / (دفعة {idx+1}/{total_to_process})")
            bar["maximum"] = total_to_process
            bar["value"] = idx + 1
            prog_win.update()

            try:
                page = pdf_doc[page_num]
                pix  = page.get_pixmap(matrix=mat)
                img  = _PIL_Image.open(_io.BytesIO(pix.tobytes("png")))

                # إرسال لـ Gemini مع retry
                response = None
                last_err = None
                for attempt in range(3):
                    try:
                        response = self.client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[img, (
                                "أنت خبير في قراءة رخص السياقة البيومترية الجزائرية. "
                                "استخرج البيانات بالتسميات الرسمية التالية بالضبط:\n"
                                "1. اللقب: الكلمة العربية - الكلمة الفرنسية\n"
                                "2. الاسم: الكلمة العربية - الكلمة الفرنسية\n"
                                "3. تاريخ ومكان الازدياد: DD/MM/YYYY - المدينة بالعربية\n"
                                "4. رقم التعريف الوطني: الرقم الطويل 15-18 رقم\n"
                                "5. رقم رخصة السياقة: الرقم الذي يبدأ بحرف L\n"
                                "6. فئات الرخصة: الحروف اللاتينية فقط مثل B أو B,C — لا تضع الجنس هنا\n"
                                "7. الجنس: M أو F فقط\n"
                                "8. تاريخ صدور رخصة السياقة: DD/MM/YYYY\n"
                                "9. تاريخ انتهاء رخصة السياقة: DD/MM/YYYY\n"
                                "10. مكان إصدار الرخصة: اسم المدينة\n"
                                "11. زمرة الدم: مثل A+ أو B-\n"
                                "12. العنوان: عنوان السكن إن وجد\n"
                                "تعليمات مهمة:\n"
                                "- اللقب والاسم بالعربية: انسخهما حرفياً كما هما مكتوبان على الرخصة بالعربية بدون أي تعديل\n"
                                "- إذا كانت الكتابة العربية غير واضحة، حوّل الاسم الفرنسي للعربية: BEL=بل، BEN=بن، BOU=بو، MOUN=مون، OUN=ون، MA=ما\n"
                                "- فئات الرخصة: لا تضع حرف الجنس (M أو F) ضمن الفئات\n"
                                "أجب بأسطر مرقمة فقط بدون مقدمة أو خاتمة."
                            )]
                        )
                        break
                    except Exception as e:
                        last_err = e
                        err_str = str(e)
                        if '503' in err_str or '429' in err_str or 'QUOTA' in err_str:
                            _time.sleep(10 * (attempt + 1))
                        else:
                            raise

                if not response or not response.text:
                    results["err"] += 1
                    log.append(f"❌ صفحة {page_num+1}: لا استجابة")
                    continue

                data = parse_text_to_dict(response.text.strip())
                nom  = data.get("الاسم واللقب", "")
                nin  = data.get("رقم التعريف الوطني", "")

                if not nin:
                    results["skip"] += 1
                    log.append(f"⚠️  صفحة {page_num+1}: لم يُقرأ رقم التعريف")
                    continue

                # حفظ تلقائي
                action, info = check_in_database(nin)
                if action == "found":
                    update_in_database(data)
                    results["ok"] += 1
                    log.append(f"🔄 صفحة {page_num+1}: تحديث — {nom}")
                else:
                    if self.on_confirm_cb:
                        self.on_confirm_cb(data, "new")
                    results["ok"] += 1
                    log.append(f"✅ صفحة {page_num+1}: إضافة — {nom}")

                lbl_res.config(text=f"✅ {results['ok']} | ⚠️ {results['skip']} | ❌ {results['err']}")
                prog_win.update()
                _time.sleep(1)  # تجنب تجاوز حد Gemini

            except Exception as e:
                results["err"] += 1
                log.append(f"❌ صفحة {page_num+1}: {str(e)[:60]}")

        pdf_doc.close()
        prog_win.destroy()

        # عرض ملخص النتائج
        summary = (
            f"✅ تم حفظ: {results['ok']} رخصة\n"
            f"⚠️  تجاهل: {results['skip']} صفحة (بيانات ناقصة)\n"
            f"❌ أخطاء:  {results['err']} صفحة\n\n"
            + "\n".join(log[-20:])
        )
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, summary)
        messagebox.showinfo("✅ اكتمل الاستيراد",
                            f"تمت معالجة {total_to_process} رخصة "
                            f"(صفحة {page_from+1}→{page_to})\n"
                            f"محفوظ: {results['ok']} | تجاهل: {results['skip']} | خطأ: {results['err']}\n\n"
                            f"للصفحات المتبقية استخدم نفس الملف غداً.")

    def _confirm(self):
        if not self._data: return
        action, info = check_in_database(
            self._data.get("رقم التعريف الوطني", ""))

        if action == "found":
            # تحديث تلقائي
            update_in_database(self._data)
            messagebox.showinfo("✅ تم التحديث",
                                f"تم تحديث بيانات: {info}")
            if self.on_confirm_cb:
                self.on_confirm_cb(self._data, "update")
            self.root.destroy()
        else:
            nom = self._data.get("الاسم واللقب", "")
            nin = self._data.get("رقم التعريف الوطني", "")
            if messagebox.askyesno("➕ مترشح جديد",
                                   f"لم يُعثر على:\n{nom}\nرقم: {nin}\n\n"
                                   f"هل تريد تسجيله؟"):
                if self.on_confirm_cb:
                    self.on_confirm_cb(self._data, "new")
                self.root.destroy()


def parse_text_to_dict(raw_text):
    """تحويل النص المستخرج لقاموس البيانات"""
    data = {
        "الاسم واللقب": "", "اللقب بالفرنسية": "", "الاسم بالفرنسية": "",
        "تاريخ ومكان الازدياد": "", "رقم التعريف الوطني": "",
        "رقم رخصة السياقة": "", "تاريخ صدور رخصة السياقة": "",
        "تاريخ انتهاء رخصة السياقة": "", "مكان إصدار الرخصة": "",
        "فئات الرخصة": "", "الجنس": "", "زمرة الدم": "",
        "العنوان": "", "_raw": raw_text
    }
    laqab_ar = ""; ism_ar = ""; laqab_fr = ""; ism_fr = ""
    dob = ""; mob = ""

    for line in raw_text.splitlines():
        line = line.strip()
        if not line: continue

        # 1. اللقب
        if re.match(r'^1[.)]\s*اللقب', line):
            parts = re.split(r'\s*[-–]\s*', line, maxsplit=1)
            laqab_ar = re.sub(r'^1[.)]\s*اللقب[:\s]*', '', parts[0]).strip()
            if len(parts) > 1: laqab_fr = parts[1].strip()

        # 2. الاسم
        elif re.match(r'^2[.)]\s*الاسم', line):
            parts = re.split(r'\s*[-–]\s*', line, maxsplit=1)
            ism_ar = re.sub(r'^2[.)]\s*الاسم[:\s]*', '', parts[0]).strip()
            if len(parts) > 1: ism_fr = parts[1].strip()

        # 3. تاريخ ومكان الازدياد
        elif re.match(r'^3[.)]\s*تاريخ ومكان الازدياد', line):
            m = re.search(r'\d{2}[./]\d{2}[./]\d{4}', line)
            if m: dob = m.group().replace('.', '/')
            parts = re.split(r'\s*[-–]\s*', line, maxsplit=1)
            if len(parts) > 1:
                mob = re.sub(r'.*\d{4}[./]\s*', '', parts[-1]).strip()
                mob = re.sub(r'\d{2}[./]\d{2}[./]\d{4}', '', mob).strip(' -–')

        # 4. رقم التعريف الوطني
        elif re.match(r'^4[.)]\s*رقم التعريف الوطني', line):
            clean = re.sub(r'[\s.]', '', line)
            m = re.search(r'\d{15,18}', clean)
            if m: data["رقم التعريف الوطني"] = m.group()
            else:
                digits = re.sub(r'\D', '', line)
                if len(digits) >= 15: data["رقم التعريف الوطني"] = digits[:18]

        # 5. رقم رخصة السياقة
        elif re.match(r'^5[.)]\s*رقم رخصة السياقة', line):
            m = re.search(r'[A-Za-z]\d{5,}|\d{6,}', line)
            if m: data["رقم رخصة السياقة"] = m.group()

        # 6. فئات الرخصة
        elif re.match(r'^6[.)]\s*فئات الرخصة', line):
            after = line.split(':', 1)[-1].strip() if ':' in line else line
            cats = re.findall(r'\b[ABCDE]\d*\b', after.upper())
            if cats: data["فئات الرخصة"] = ", ".join(cats)

        # 7. الجنس
        elif re.match(r'^7[.)]\s*الجنس', line):
            if 'ذكر' in line or re.search(r'[-:]\s*M\b', line):
                data["الجنس"] = "M"
            elif 'أنثى' in line or re.search(r'[-:]\s*F\b', line):
                data["الجنس"] = "F"

        # 8. تاريخ صدور رخصة السياقة
        elif re.match(r'^8[.)]\s*تاريخ صدور رخصة السياقة', line):
            m = re.search(r'\d{2}[./]\d{2}[./]\d{4}', line)
            if m: data["تاريخ صدور رخصة السياقة"] = m.group().replace('.', '/')

        # 9. تاريخ انتهاء رخصة السياقة
        elif re.match(r'^9[.)]\s*تاريخ انتهاء رخصة السياقة', line):
            m = re.search(r'\d{2}[./]\d{2}[./]\d{4}', line)
            if m: data["تاريخ انتهاء رخصة السياقة"] = m.group().replace('.', '/')

        # 10. مكان إصدار الرخصة
        elif re.match(r'^10[.)]\s*مكان إصدار الرخصة', line):
            clean = re.sub(r'^10[.)]\s*مكان[^:]*[:\s]*', '', line).strip()
            if clean: data["مكان إصدار الرخصة"] = clean

        # 11. زمرة الدم
        elif re.match(r'^11[.)]\s*زمرة الدم', line):
            m = re.search(r'[ABO]{1,2}[+-]', line)
            if m: data["زمرة الدم"] = m.group()

        # 12. العنوان
        elif re.match(r'^12[.)]\s*العنوان', line):
            clean = re.sub(r'^12[.)]\s*العنوان[:\s]*', '', line).strip()
            if clean and 'غير' not in clean: data["العنوان"] = clean

        # رقم التعريف الوطني كـ fallback (أي سطر يحتوي 15-18 رقم)
        elif not data["رقم التعريف الوطني"]:
            if 'رقم التعريف' in line or 'الرقم الوطني' in line:
                clean = re.sub(r'[\s.]', '', line)
                m = re.search(r'\d{15,18}', clean)
                if m: data["رقم التعريف الوطني"] = m.group()

    data["الاسم واللقب"]          = f"{laqab_ar} {ism_ar}".strip()
    data["اللقب بالفرنسية"]       = laqab_fr
    data["الاسم بالفرنسية"]       = ism_fr
    data["تاريخ ومكان الازدياد"]  = f"{dob} {mob}".strip() if mob else dob
    return data


def check_in_database(nin):
    """البحث عن المترشح في Excel"""
    try:
        import openpyxl
        excel = os.path.join(BASE_DIR, "بيانات_مركز_التكوين.xlsx")
        if not os.path.exists(excel): return "new", {}
        wb = openpyxl.load_workbook(excel)
        ws = wb["المترشحون"]
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            r_nin = str(row[6] or "").strip() if len(row) > 6 else ""
            if r_nin == nin.strip():
                return "found", str(row[1] or "")
    except: pass
    return "new", {}


def update_in_database(data):
    """تحديث بيانات مترشح موجود"""
    try:
        import openpyxl
        excel = os.path.join(BASE_DIR, "بيانات_مركز_التكوين.xlsx")
        wb = openpyxl.load_workbook(excel)
        ws = wb["المترشحون"]
        nin = data.get("رقم التعريف الوطني", "")
        for i, row in enumerate(ws.iter_rows(min_row=2), 2):
            if str(row[6].value or "").strip() == nin.strip():
                mp = {2:"الاسم واللقب", 3:"اللقب بالفرنسية", 4:"الاسم بالفرنسية",
                      5:"تاريخ ومكان الازدياد", 8:"رقم رخصة السياقة",
                      9:"تاريخ صدور رخصة السياقة", 12:"فئات الرخصة", 14:"الجنس"}
                for col, key in mp.items():
                    v = data.get(key, "")
                    if v: ws.cell(row=i, column=col, value=v)
                wb.save(excel)
                return True
    except: pass
    return False


def open_reader(parent_app=None):
    """فتح نافذة قراءة الرخصة"""
    if parent_app:
        try:
            win = tk.Toplevel(parent_app.winfo_toplevel())
            win.transient(parent_app.winfo_toplevel())
        except:
            win = tk.Toplevel()
    else:
        win = tk.Tk()

    def on_confirm(data, action):
        if not parent_app:
            import tkinter.messagebox as _mb
            _mb.showwarning("debug", "parent_app = None!")
            return
        if action == "new":
            try:
                mv = parent_app.mv
                if not mv:
                    import tkinter.messagebox as _mb
                    _mb.showwarning("debug", "mv فارغ!")
                    return
                mapping = {
                    "الاسم واللقب":              "الاسم واللقب",
                    "اللقب بالفرنسية":           "اللقب بالفرنسية",
                    "الاسم بالفرنسية":           "الاسم بالفرنسية",
                    "تاريخ ومكان الازدياد":      "تاريخ ومكان الازدياد",
                    "رقم التعريف الوطني":        "رقم التعريف الوطني",
                    "رقم رخصة السياقة":          "رقم رخصة السياقة",
                    "تاريخ صدور رخصة السياقة":   "تاريخ صدور رخصة السياقة",
                    "تاريخ انتهاء رخصة السياقة": "تاريخ انتهاء رخصة السياقة",
                    "مكان إصدار الرخصة":         "مكان إصدار الرخصة",
                    "فئات الرخصة":               "فئات الرخصة",
                    "الجنس":                     "الجنس",
                    "زمرة الدم":                 "زمرة الدم",
                    "العنوان":                   "العنوان",
                }
                count = 0
                for src_k, dst_k in mapping.items():
                    val = data.get(src_k, "")
                    if val and dst_k in mv:
                        mv[dst_k].set(val)
                        count += 1
                if count > 0:
                    messagebox.showinfo("✅",
                        f"تم ملء {count} خانة تلقائياً\n"
                        "أضف التخصص ورقم الدورة ثم اضغط ➕ إضافة")
                else:
                    messagebox.showwarning("⚠️", "لم يتم ملء أي خانة")
            except Exception as e:
                messagebox.showerror("❌", f"خطأ في ملء الفورم:\n{e}")

        elif action == "update":
            try:
                if hasattr(parent_app, 'refresh'):
                    parent_app.refresh()
                messagebox.showinfo("✅", "تم تحديث البيانات")
            except: pass

    UltimateOCRApp(win, on_confirm_cb=on_confirm)
    if not parent_app:
        win.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateOCRApp(root)
    root.mainloop()
