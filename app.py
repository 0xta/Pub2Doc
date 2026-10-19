"""
Pub2Doc Studio - Desktop Application
Tool for converting Microsoft Publisher (.pub) files to Microsoft Word (.docx)
Preserves Thai typography, layout tables, data tables, and high-res images.
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Theme
try:
    import sv_ttk
    HAS_SV_TTK = True
except ImportError:
    HAS_SV_TTK = False

from pub2doc_engine import PubConverter, check_publisher_available

class Pub2DocApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pub2Doc Studio - ตัวแปลงไฟล์ Publisher เป็น Word DOCX")
        self.root.geometry("820x680")
        self.root.minsize(720, 560)

        # Apply modern Windows 11 theme if available
        if HAS_SV_TTK:
            sv_ttk.set_theme("light")

        self.files_to_convert = []
        self.output_dir = os.path.abspath(os.getcwd())
        self.is_converting = False

        self._build_ui()
        self._check_environment()

    def _build_ui(self):
        # Top Container
        main_frame = ttk.Frame(self.root, padding=16)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Header & Engine Status
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 12))

        title_lbl = ttk.Label(header_frame, text="Pub2Doc Studio", font=("Segoe UI", 16, "bold"))
        title_lbl.pack(anchor=tk.W)
        subtitle_lbl = ttk.Label(
            header_frame, 
            text="แปลงไฟล์ Microsoft Publisher (.pub) เป็น Word (.docx) โดยคงภาษาไทย ตาราง และรูปภาพสมบูรณ์แบบ",
            font=("Segoe UI", 9)
        )
        subtitle_lbl.pack(anchor=tk.W, pady=(2, 6))

        # Status Badge
        self.status_badge = ttk.Label(
            header_frame, 
            text=" กำลังตรวจสอบระบบ... ", 
            font=("Segoe UI", 9, "bold"),
            padding=(8, 4),
            relief="groove"
        )
        self.status_badge.pack(anchor=tk.W)

        # 2. File Selection Area
        file_box = ttk.LabelFrame(main_frame, text=" รายการไฟล์ที่ต้องการแปลง ", padding=10)
        file_box.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Button Row
        btn_row = ttk.Frame(file_box)
        btn_row.pack(fill=tk.X, pady=(0, 8))

        btn_add_file = ttk.Button(btn_row, text="➕ เพิ่มไฟล์ .PUB", command=self._add_files)
        btn_add_file.pack(side=tk.LEFT, padx=(0, 6))

        btn_add_folder = ttk.Button(btn_row, text="📁 เพิ่มทั้งโฟลเดอร์", command=self._add_folder)
        btn_add_folder.pack(side=tk.LEFT, padx=(0, 6))

        btn_clear = ttk.Button(btn_row, text="🗑️ ล้างรายการ", command=self._clear_files)
        btn_clear.pack(side=tk.RIGHT)

        # Treeview / List
        tree_frame = ttk.Frame(file_box)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("name", "path", "status")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=6)
        self.tree.heading("name", text="ชื่อไฟล์")
        self.tree.heading("path", text="ที่อยู่ไฟล์")
        self.tree.heading("status", text="สถานะ")
        self.tree.column("name", width=260)
        self.tree.column("path", width=340)
        self.tree.column("status", width=120, anchor=tk.CENTER)

        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 3. Output Options
        opts_frame = ttk.LabelFrame(main_frame, text=" ตั้งค่าการบันทึกผลลัพธ์ ", padding=10)
        opts_frame.pack(fill=tk.X, pady=(0, 10))

        out_row = ttk.Frame(opts_frame)
        out_row.pack(fill=tk.X)

        ttk.Label(out_row, text="โฟลเดอร์ปลายทาง:").pack(side=tk.LEFT, padx=(0, 6))
        self.out_dir_var = tk.StringVar(value=self.output_dir)
        out_entry = ttk.Entry(out_row, textvariable=self.out_dir_var)
        out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        btn_browse_out = ttk.Button(out_row, text="เลือก...", width=8, command=self._browse_output_dir)
        btn_browse_out.pack(side=tk.LEFT, padx=(0, 6))

        btn_open_out = ttk.Button(out_row, text="📂 เปิดโฟลเดอร์", command=self._open_output_dir)
        btn_open_out.pack(side=tk.LEFT)

        # 4. Progress & Action
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.progress_bar = ttk.Progressbar(action_frame, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        self.btn_convert = ttk.Button(
            action_frame, 
            text="🚀 เริ่มแปลงไฟล์เป็น Word (.docx)", 
            style="Accent.TButton" if HAS_SV_TTK else "TButton",
            command=self._start_conversion
        )
        self.btn_convert.pack(fill=tk.X, ipady=6)

        # 5. Log Console
        log_box = ttk.LabelFrame(main_frame, text=" บันทึกการทำงาน (Log) ", padding=8)
        log_box.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_box, height=7, font=("Consolas", 9), wrap=tk.WORD)
        log_scroll = ttk.Scrollbar(log_box, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _check_environment(self):
        has_pub, ver = check_publisher_available()
        if has_pub:
            self.status_badge.config(
                text=f"🟢 ตรวจพบ MS Publisher (v{ver}) - ใช้งานโหมดความแม่นยำสูง (Native COM Mode)",
                foreground="#0d6832"
            )
            self._log(f"ตรวจสอบระบบ: พบ Microsoft Publisher เวอร์ชัน {ver} (พร้อมแปลงข้อมูลระดับ Native)")
        else:
            self.status_badge.config(
                text="🟡 ไม่พบ MS Publisher ในเครื่อง - ใช้งานโหมด Standalone Fallback",
                foreground="#946200"
            )
            self._log("ตรวจสอบระบบ: ไม่พบ Microsoft Publisher จะใช้โหมด Standalone สำรอง")

    def _log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="เลือกไฟล์ Publisher",
            filetypes=[("Publisher Documents", "*.pub"), ("All Files", "*.*")]
        )
        if files:
            for f in files:
                f_abs = os.path.abspath(f)
                if f_abs not in self.files_to_convert:
                    self.files_to_convert.append(f_abs)
                    fname = os.path.basename(f_abs)
                    self.tree.insert("", tk.END, values=(fname, f_abs, "รอการแปลง"))
            self._log(f"เพิ่มไฟล์ {len(files)} ไฟล์เรียบร้อย")

    def _add_folder(self):
        folder = filedialog.askdirectory(title="เลือกโฟลเดอร์ที่มีไฟล์ Publisher")
        if folder:
            count = 0
            for root_dir, _, filenames in os.walk(folder):
                for fname in filenames:
                    if fname.lower().endswith(".pub"):
                        f_abs = os.path.abspath(os.path.join(root_dir, fname))
                        if f_abs not in self.files_to_convert:
                            self.files_to_convert.append(f_abs)
                            self.tree.insert("", tk.END, values=(fname, f_abs, "รอการแปลง"))
                            count += 1
            self._log(f"พบไฟล์ .PUB ในโฟลเดอร์ {count} ไฟล์")

    def _clear_files(self):
        self.files_to_convert.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._log("ล้างรายการไฟล์ทั้งหมดแล้ว")

    def _browse_output_dir(self):
        dir_selected = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึกไฟล์ Word")
        if dir_selected:
            self.output_dir = os.path.abspath(dir_selected)
            self.out_dir_var.set(self.output_dir)

    def _open_output_dir(self):
        target = self.out_dir_var.get()
        if os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showwarning("แจ้งเตือน", "ไม่พบโฟลเดอร์ปลายทางที่ระบุ")

    def _start_conversion(self):
        if not self.files_to_convert:
            messagebox.showinfo("แจ้งเตือน", "กรุณาเพิ่มไฟล์ .PUB อย่างน้อย 1 ไฟล์ก่อนเริ่มแปลง")
            return

        if self.is_converting:
            return

        self.is_converting = True
        self.btn_convert.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0

        # Run conversion in background thread so GUI stays responsive
        t = threading.Thread(target=self._run_batch_conversion, daemon=True)
        t.start()

    def _run_batch_conversion(self):
        total = len(self.files_to_convert)
        out_dir = self.out_dir_var.get()
        os.makedirs(out_dir, exist_ok=True)

        items = self.tree.get_children()
        successful = 0

        for idx, item in enumerate(items):
            f_path = self.files_to_convert[idx]
            fname = os.path.basename(f_path)
            base_name, _ = os.path.splitext(fname)
            out_docx = os.path.join(out_dir, f"{base_name}_converted.docx")

            self.tree.item(item, values=(fname, f_path, "⏳ กำลังแปลง..."))
            self._log(f"\n--- เริ่มแปลง [{idx+1}/{total}]: {fname} ---")

            try:
                converter = PubConverter(f_path, out_docx, log_fn=self._log)
                converter.convert()
                self.tree.item(item, values=(fname, f_path, "✅ สำเร็จ"))
                successful += 1
            except Exception as e:
                self.tree.item(item, values=(fname, f_path, "❌ ผิดพลาด"))
                self._log(f"ข้อผิดพลาด: {e}")

            progress_val = int(((idx + 1) / total) * 100)
            self.root.after(0, lambda v=progress_val: self.progress_bar.configure(value=v))

        self.root.after(0, lambda: self._on_conversion_finished(successful, total))

    def _on_conversion_finished(self, successful, total):
        self.is_converting = False
        self.btn_convert.config(state=tk.NORMAL)
        self._log(f"\n🎉 การแปลงเสร็จสิ้นทั้งหมด: สำเร็จ {successful} จาก {total} ไฟล์")
        
        res = messagebox.askyesno(
            "การแปลงสำเร็จ", 
            f"แปลงเสร็จสิ้น {successful}/{total} ไฟล์เรียบร้อยแล้ว!\nต้องการเปิดโฟลเดอร์ผลลัพธ์เลยหรือไม่?"
        )
        if res:
            self._open_output_dir()

def main():
    root = tk.Tk()
    app = Pub2DocApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
