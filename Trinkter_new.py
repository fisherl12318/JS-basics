import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import re
from tkinter import filedialog

from Pdf_filling import fill_pdf_fields
from Excel import find_items


class PDFFillerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF自动填写")
        self.root.geometry("900x700")

        # ================= 设备 → 探头映射 =================
        self.device_map = {
            "Acclarix AX3 Series": {
                "name": "AX3",
                "probes": ["C5-2Q", "P5-1Q", "E8-4Q", "L17-7HQ", "L17-7SQ", "MC8-4Q","P7-3Q", "MC9-3TQ",
                           "C6-2MQ", "C5-1Q", "E10-3HQ", "E10-3BQ", "L12-5HQ", "L12-5Q", "L12-5WQ",
                           "ECL12-3Q", "E10-3Q", "C5-2XQ"
                           ]
            },
            "Acclarix LX3 Series": {
                "name": "LX3",
                "probes": ["L17-7HQ", "E8-4Q", "C5-2Q", "L12-5Q", "MC8-4Q", "L17-7SQ",
                           "P5-1Q", "P7-3Q", "MC9-3TQ", "E10-3HQ", "C6-2MQ", "C5-1Q", "L12-5HQ",
                           "E10-3BQ", "L12-5WQ", "C5-2XQ", "E10-3Q"
                           ]
            },
            "Acclarix LX25": {
                "name": "LX25",
                "probes": [
                    "C5-2Q", "P5-1Q", "E8-4Q", "L17-7HQ", "L17-7SQ",
                    "MC8-4Q", "P7-3Q", "MC9-3TQ", "C6-2MQ",
                    "C5-1Q", "E10-3HQ", "E10-3BQ", "L12-5HQ", "P7-3Q",
                    "L12-5Q", "L12-5WQ", "ECL12-3Q", "E10-3Q", "C5-2XQ"
                ]
            },
            "DUS60": {
                "name": "DUS 60",
                "probes": ["C361-2", "L761-2", "L743-2", "C611-2", "E611-2", "L741-2"]
            },
            "U50": {
                "name": "U50 Diagnostic Ultrasound System",
                "probes": ["L15-7b", "P5-1b", "C352UB", "L742UB"]
            },
            "U60": {
                "name": "U60 Diagnostic Ultrasound System",
                "probes": ["L15-7b", "P5-1b", "C352UB"]
            }
        }

        self.probe_list = []  # 初始为空（等用户选设备）

        self.setup_ui()

    def setup_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # ================= 文件 =================
        ttk.Label(main, text="PDF文件").grid(row=0, column=0)

        self.pdf_path_var = tk.StringVar()

        ttk.Entry(main, textvariable=self.pdf_path_var, width=50).grid(row=0, column=1)

        ttk.Button(main, text="浏览", command=self.browse_pdf).grid(row=0, column=2)

        # ================= 设备信息 =================
        ttk.Label(main, text="设备型号").grid(row=2, column=0)

        self.device_var = tk.StringVar()
        device_combo = ttk.Combobox(
            main,
            textvariable=self.device_var,
            values=list(self.device_map.keys()),
            state="readonly",
            width=40
        )
        device_combo.grid(row=2, column=1)
        device_combo.bind("<<ComboboxSelected>>", self.on_device_selected)

        ttk.Label(main, text="公司名称").grid(row=1, column=0)
        self.firma_name_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.firma_name_var, width=50).grid(row=1, column=1)

        ttk.Label(main, text="设备名称").grid(row=3, column=0)
        self.name_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.name_var, width=50).grid(row=3, column=1)

        ttk.Label(main, text="设备号(SN)").grid(row=4, column=0)
        self.geraet_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.geraet_var).grid(row=4, column=1)
        self.geraet_var.trace_add("write", self._on_sn_changed)

        self.baujahr_var = tk.StringVar()

        ttk.Label(main, text="交货日期").grid(row=5, column=0)
        self.datum_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.datum_var).grid(row=5, column=1)

        # ================= 探头 =================
        self.left_listbox = tk.Listbox(main, selectmode=tk.MULTIPLE, height=10)
        self.left_listbox.grid(row=7, column=0)

        self.right_listbox = tk.Listbox(main, selectmode=tk.MULTIPLE, height=10)
        self.right_listbox.grid(row=7, column=2)

        # ttk.Button(main, text="→", command=self.add_selected).grid(row=7, column=1)
        # ttk.Button(main, text="←", command=self.remove_selected).grid(row=8, column=1)

        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=7, column=1)

        ttk.Button(btn_frame, text="→", command=self.add_selected).pack(pady=5)
        ttk.Button(btn_frame, text="←", command=self.remove_selected).pack(pady=5)


        # ================= 输出 =================
        ttk.Label(main, text="输出文件").grid(row=9, column=0)
        self.output_var = tk.StringVar(value="output.pdf")
        ttk.Entry(main, textvariable=self.output_var).grid(row=9, column=1)

        # ================= 日志 =================
        self.log_text = scrolledtext.ScrolledText(main, height=9)
        self.log_text.grid(row=10, column=0, columnspan=3)

        ttk.Button(main, text="开始", command=self.start_filling).grid(row=11, column=1)


        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.grid(row=12, column=0, columnspan=3, sticky="ew")

    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            self.pdf_path_var.set(file_path)
            self.log(f"已选择PDF: {file_path}")

    # ================= SN → Baujahr =================
    def _on_sn_changed(self, *_):
        sn = self.geraet_var.get()
        m = re.search(r'M(\d{2})', sn)
        if m:
            self.baujahr_var.set("20" + m.group(1))
        else:
            self.baujahr_var.set("")

    # ================= 联动核心 =================
    def on_device_selected(self, event):
        key = self.device_var.get()
        data = self.device_map.get(key)

        if not data:
            return

        # 自动填名称
        self.name_var.set(data["name"])

        # 更新探头
        self.probe_list = data["probes"]

        self.left_listbox.delete(0, tk.END)
        for p in self.probe_list:
            self.left_listbox.insert(tk.END, p)

        # 清空已选
        self.right_listbox.delete(0, tk.END)
        self.log(f"已选择设备: {key}")

    def add_selected(self):
        for i in self.left_listbox.curselection():
            self.right_listbox.insert(tk.END, self.left_listbox.get(i))

    def remove_selected(self):
        selected = self.right_listbox.curselection()
        for index in reversed(selected):
            self.right_listbox.delete(index)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)

    def start_filling(self):
        probes = list(self.right_listbox.get(0, tk.END))
        pdf = self.pdf_path_var.get()
        output = self.output_var.get()

        items = find_items(" ".join(probes))

        device_name = self.name_var.get()

        if device_name == "DUS 60":
            items.append("AK 20.12")
        else:
            items.append("AK 20.11")

        self.log(f"AK结果: {items}")

        self.progress.start()

        def run():
            try:
                fill_pdf_fields(
                    pdf_path=pdf,
                    probes=probes,
                    items_to_check=items,  # 这里直接传 AK 列表
                    output_path=output,
                    adresse=self.firma_name_var.get(),
                    name=self.name_var.get(),
                    geraet=self.geraet_var.get(),
                    baujahr=self.baujahr_var.get(),
                    datum=self.datum_var.get()
                )
                self.log("✅ 完成")
            except Exception as e:
                self.log(f"❌ {e}")
            finally:
                self.progress.stop()

        threading.Thread(target=run).start()


def main():
    root = tk.Tk()
    PDFFillerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()