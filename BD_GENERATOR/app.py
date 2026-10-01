"""Local Tk review UI and headless CLI for the two-stage workflow."""
from __future__ import annotations
import argparse
import copy
import logging
import os
import queue
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from generator import VERSION, Cancelled, generate, initial_rows, load_package, load_review, merge_rows, save_review


class App:
    def __init__(self, root):
        self.root = root
        root.title(VERSION); root.geometry("1200x780")
        self.analysis = None; self.rows = []; self.busy = False
        self.cancel = threading.Event(); self.events = queue.Queue()
        log_dir = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "변대생기" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = log_dir / ("alpha09_" + time.strftime("%Y%m%d_%H%M%S") + ".log")
        logging.basicConfig(filename=self.log_path, encoding="utf-8", level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
        self.name = tk.StringVar(); self.mode = tk.StringVar(value="editable")
        self.status = tk.StringVar(value="변분기 Alpha 0.4의 .bdsg 결과물을 선택하십시오.")
        self.actions = []
        bar = ttk.Frame(root, padding=8); bar.pack(fill="x")
        self.button(bar, ".bdsg 선택", self.open_package)
        ttk.Label(bar, text="문서명").pack(side="left", padx=8)
        self.name_entry = ttk.Entry(bar, textvariable=self.name, width=65); self.name_entry.pack(side="left", fill="x", expand=True)
        self.button(bar, "로그 열기", lambda: os.startfile(self.log_path), managed=False)
        toolbar = ttk.Frame(root, padding=8); toolbar.pack(fill="x")
        for label, fn in (("Section·page 수정", self.edit), ("선택 병합", self.merge), ("포함/제외", self.toggle), ("검토 초기화", self.reset), ("검토 저장", self.save), ("검토 불러오기", self.load)):
            self.button(toolbar, label, fn)
        pane = ttk.Panedwindow(root, orient="vertical"); pane.pack(fill="both", expand=True, padx=8)
        frame = ttk.Frame(pane); pane.add(frame, weight=2)
        self.tree = ttk.Treeview(frame, columns=("include", "id", "section", "page", "type", "warning"), show="headings", selectmode="extended")
        for key, label, width in (("include", "포함", 50), ("id", "변경 ID", 170), ("section", "Section", 290), ("page", "page", 150), ("type", "유형", 90), ("warning", "확인사항", 230)):
            self.tree.heading(key, text=label); self.tree.column(key, width=width)
        self.tree.tag_configure("excluded", foreground="#999999")
        self.tree.tag_configure("review", foreground="#9B4F00")
        self.tree.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview); scroll.pack(side="right", fill="y"); self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.preview()); self.tree.bind("<Double-1>", lambda e: self.edit())
        bottom = ttk.Frame(pane); pane.add(bottom, weight=1)
        self.detail = tk.Text(bottom, wrap="word", height=12, state="disabled"); self.detail.pack(fill="both", expand=True)
        footer = ttk.Frame(root, padding=8); footer.pack(fill="x")
        self.radio_edit = ttk.Radiobutton(footer, text="편집 가능한 텍스트·표", variable=self.mode, value="editable"); self.radio_edit.pack(side="left")
        self.radio_images = ttk.Radiobutton(footer, text="패키지 이미지 (실패 시 텍스트)", variable=self.mode, value="images"); self.radio_images.pack(side="left", padx=12)
        self.button(footer, "변경대비표 생성", self.create)
        self.cancel_button = ttk.Button(footer, text="작업 취소", command=self.cancel.set, state="disabled"); self.cancel_button.pack(side="left", padx=8)
        self.progress = ttk.Progressbar(root, mode="determinate"); self.progress.pack(fill="x", padx=8)
        ttk.Label(root, textvariable=self.status, padding=8).pack(fill="x")
        root.protocol("WM_DELETE_WINDOW", self.close); root.after(100, self.poll)

    def button(self, frame, text, command, managed=True):
        b = ttk.Button(frame, text=text, command=command); b.pack(side="left", padx=3)
        if managed: self.actions.append(b)

    def notify(self, title, text):
        self.root.lift(); self.root.bell(); messagebox.showinfo(title, text, parent=self.root)

    def task(self, fn, kind):
        self.busy = True; self.cancel.clear(); self.progress["value"] = 0
        self.status.set("패키지 읽는 중…" if kind == "loaded" else "변경대비표 생성 중…")
        for b in self.actions + [self.name_entry, self.radio_edit, self.radio_images]: b.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        def work():
            try: self.events.put((kind, fn()))
            except Cancelled: self.events.put(("cancelled", None))
            except Exception as exc:
                logging.exception("작업 실패")
                self.events.put(("error", f"{type(exc).__name__}: {str(exc)[:400]}\n로그: {self.log_path}"))
        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "progress":
                    n, total = value; self.progress["value"] = n / total * 95; self.status.set(f"변경대비표 작성 {n}/{total}"); continue
                self.busy = False
                for b in self.actions + [self.name_entry, self.radio_edit, self.radio_images]: b.configure(state="normal")
                self.cancel_button.configure(state="disabled")
                if kind == "loaded":
                    self.analysis = value; self.rows = initial_rows(value); self.name.set(value.manifest["document_name"]); self.refresh()
                    self.progress["value"] = 100; self.status.set(f"변경 항목 {len(self.rows)}개 · {value.path}")
                    self.notify("패키지 읽기 완료", f"변경 항목 {len(self.rows)}개를 읽었습니다.\nSection·page와 확인사항을 검토한 뒤 생성하십시오.")
                elif kind == "generated":
                    path, warnings = value; self.progress["value"] = 100; self.status.set(f"완료: {path}")
                    self.notify("변경대비표 생성 완료", f"{path}\n확인사항 {len(warnings)}개는 문서 끝에 기록했습니다.")
                elif kind == "error": self.status.set("작업 실패"); messagebox.showerror("오류", value, parent=self.root)
                else: self.status.set("작업을 취소했습니다. 이전 결과 파일은 보존됩니다.")
        except queue.Empty: pass
        self.root.after(100, self.poll)

    def open_package(self):
        path = filedialog.askopenfilename(filetypes=[("변분기 분석 패키지", "*.bdsg")])
        if path: self.task(lambda: load_package(path, self.cancel), "loaded")

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        if not self.analysis: return
        items = {i["change_id"]: i for i in self.analysis.manifest["items"]}
        for idx, row in enumerate(self.rows):
            check = any(items[cid].get("needs_review") or items[cid].get("image_error") for cid in row["ids"])
            tags = ("excluded",) if not row["included"] else ("review",) if check else ()
            self.tree.insert("", "end", iid=str(idx), values=("포함" if row["included"] else "제외", ", ".join(row["ids"]), row["section"], row["page"], "/".join(dict.fromkeys(items[cid]["context_type"] for cid in row["ids"])), "원문·대응 확인 필요" if check else ""), tags=tags)

    def indices(self): return sorted(int(s) for s in self.tree.selection())

    def preview(self):
        if not self.analysis: return
        items = {i["change_id"]: i for i in self.analysis.manifest["items"]}
        text = []
        for idx in self.indices():
            for cid in self.rows[idx]["ids"]:
                i = items[cid]; text.append(f"{cid}\n[변경 전]\n{i['before']}\n[변경 후]\n{i['after']}\n[확인사항]\n{i.get('image_error', '')}")
        if not text: text = self.analysis.manifest.get("warnings", [])
        self.detail.configure(state="normal"); self.detail.delete("1.0", "end"); self.detail.insert("end", "\n\n".join(text)); self.detail.configure(state="disabled")

    def edit(self):
        if self.busy: return
        selected = self.indices()
        if len(selected) != 1: return
        row = self.rows[selected[0]]
        section = simpledialog.askstring("Section 수정", "Section", initialvalue=row["section"], parent=self.root)
        if section is None: return
        page = simpledialog.askstring("page 수정", "전후 페이지", initialvalue=row["page"], parent=self.root)
        if page is None: return
        row.update(section=section, page=page); self.refresh()

    def merge(self):
        try: merge_rows(self.rows, self.indices()); self.refresh()
        except ValueError as exc: messagebox.showwarning("병합", str(exc))

    def toggle(self):
        for idx in self.indices(): self.rows[idx]["included"] = not self.rows[idx]["included"]
        self.refresh()

    def reset(self):
        if self.analysis and messagebox.askyesno("검토 초기화", "병합·제외·수정 내용을 초기화합니까?"):
            self.rows = initial_rows(self.analysis); self.refresh()

    def save(self):
        if not self.analysis: return
        path = filedialog.asksaveasfilename(defaultextension=".json", initialfile=self.analysis.path.stem + "_검토.json", filetypes=[("검토 파일", "*.json")])
        if path:
            try: save_review(self.analysis, self.rows, path); self.status.set(f"검토 저장: {path}")
            except Exception as exc: messagebox.showerror("검토 저장 오류", str(exc))

    def load(self):
        if not self.analysis: return
        path = filedialog.askopenfilename(filetypes=[("검토 파일", "*.json")])
        if path:
            try: self.rows = load_review(self.analysis, path); self.refresh(); self.status.set("검토 내용을 복원했습니다.")
            except Exception as exc: messagebox.showerror("검토 파일 오류", str(exc))

    def create(self):
        if not self.analysis: return
        path = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=self.analysis.path.stem + "_변경대비표.docx", filetypes=[("Word 문서", "*.docx")])
        if not path: return
        rows = copy.deepcopy(self.rows); mode = self.mode.get(); name = self.name.get()
        def work():
            warnings = generate(self.analysis, rows, path, mode=mode, document_name=name, cancel=self.cancel,
                                progress=lambda n, total: self.events.put(("progress", (n, total))))
            logging.info("생성 완료: %s / 확인사항 %s개", path, len(warnings)); return path, warnings
        self.task(work, "generated")

    def close(self):
        if self.busy:
            self.cancel.set(); self.status.set("취소 중입니다. 작업이 끝나면 창을 닫으십시오."); return
        self.root.destroy()


def main():
    parser = argparse.ArgumentParser(description=VERSION)
    parser.add_argument("--input", type=Path); parser.add_argument("--output", type=Path)
    parser.add_argument("--review", type=Path); parser.add_argument("--mode", choices=["editable", "images"], default="editable")
    args = parser.parse_args()
    if args.input:
        if not args.output: parser.error("--input 사용 시 --output이 필요합니다")
        analysis = load_package(args.input)
        rows = load_review(analysis, args.review) if args.review else initial_rows(analysis)
        warnings = generate(analysis, rows, args.output, mode=args.mode)
        print(f"완료: {args.output} / 확인사항 {len(warnings)}개")
    else:
        if args.output or args.review: parser.error("--input이 필요합니다")
        root = tk.Tk(); App(root); root.mainloop()


if __name__ == "__main__": main()
