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

from generator import VERSION, Cancelled, generate, initial_rows, load_package, load_review, merge_rows, reanalysis_warnings, save_review, set_matching


class App:
    def __init__(self, root):
        self.root = root
        root.title(VERSION); root.geometry("1200x780")
        self.analysis = None; self.rows = []; self.busy = False; self.review_dirty = False
        self.cancel = threading.Event(); self.events = queue.Queue()
        log_dir = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "변대생기" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = log_dir / ("alpha11_" + time.strftime("%Y%m%d_%H%M%S") + ".log")
        logging.basicConfig(filename=self.log_path, encoding="utf-8", level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
        self.name = tk.StringVar(); self.mode = tk.StringVar(value="editable")
        self.status = tk.StringVar(value="변분기 Alpha 1.1의 .bdsg를 선택하십시오. 기존 Alpha 1.0·0.5·0.4도 지원합니다.")
        self.actions = []
        bar = ttk.Frame(root, padding=8); bar.pack(fill="x")
        self.button(bar, ".bdsg 선택", self.open_package)
        ttk.Label(bar, text="문서명").pack(side="left", padx=8)
        self.name_entry = ttk.Entry(bar, textvariable=self.name, width=65); self.name_entry.pack(side="left", fill="x", expand=True)
        self.button(bar, "로그 열기", lambda: os.startfile(self.log_path), managed=False)
        toolbar = ttk.Frame(root, padding=8); toolbar.pack(fill="x")
        for label, fn in (("전후 매칭 수정", self.match), ("Section·page 수정", self.edit), ("선택 병합", self.merge), ("포함/제외", self.toggle), ("검토 초기화", self.reset), ("검토 저장", self.save), ("검토 불러오기", self.load)):
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
                    self.analysis = value; self.rows = initial_rows(value); self.review_dirty = False; self.name.set(value.manifest["document_name"]); self.refresh()
                    self.progress["value"] = 100; self.status.set(f"변경 항목 {len(self.rows)}개 · {value.path}")
                    scope_note = '\n'.join(reanalysis_warnings(value.manifest)) or "문서 앞부분·본문 포함, 목차·별첨·부록 제외."
                    self.notify("패키지 읽기 완료", f"변경 항목 {len(self.rows)}개를 읽었습니다.\n{scope_note}\n전후 내용·대응 관계와 Section·page를 검토한 뒤 생성하십시오.")
                elif kind == "generated":
                    path, warnings = value; self.progress["value"] = 100; self.status.set(f"완료: {path}")
                    note = "\n수정한 검토 내용을 다시 사용하려면 ‘검토 저장’을 누르십시오." if self.review_dirty else ""
                    self.notify("변경대비표 생성 완료", f"{path}\n확인사항 {len(warnings)}개는 문서 끝에 기록했습니다.{note}")
                elif kind == "error": self.status.set("작업 실패"); messagebox.showerror("오류", value, parent=self.root)
                else: self.status.set("작업을 취소했습니다. 이전 결과 파일은 보존됩니다.")
        except queue.Empty: pass
        self.root.after(100, self.poll)

    def open_package(self):
        if self.review_dirty and not messagebox.askyesno("저장하지 않은 검토", "현재 검토 내용을 저장하지 않고 다른 패키지를 열겠습니까?", parent=self.root): return
        path = filedialog.askopenfilename(filetypes=[("변분기 분석 패키지", "*.bdsg")])
        if path: self.task(lambda: load_package(path, self.cancel), "loaded")

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        if not self.analysis: return
        items = {i["change_id"]: i for i in self.analysis.manifest["items"]}
        for idx, row in enumerate(self.rows):
            check = any(items[cid].get("needs_review") or items[cid].get("image_error") for cid in row["ids"])
            changed = any((items[cid].get(side + "_block") is not None) != (cid in row[side + "_ids"]) for cid in row["ids"] for side in ("before", "after"))
            check = check or changed
            tags = ("excluded",) if not row["included"] else ("review",) if check else ()
            note = "수동 매칭 수정" if changed else "원문·대응 확인 필요" if check else ""
            self.tree.insert("", "end", iid=str(idx), values=("포함" if row["included"] else "제외", ", ".join(row["ids"]), row["section"], row["page"], "/".join(dict.fromkeys(items[cid]["context_type"] for cid in row["ids"])), note), tags=tags)
        self.preview()

    def indices(self): return sorted(int(s) for s in self.tree.selection())

    def preview(self):
        if not self.analysis: return
        items = {i["change_id"]: i for i in self.analysis.manifest["items"]}
        text = []
        for idx in self.indices():
            row = self.rows[idx]
            text.append(f"{row['section']} · {row['page']}")
            for side, label in (("before", "변경 전"), ("after", "변경 후")):
                text.append(f"[{label}]")
                for cid in row[side + "_ids"]:
                    item = items[cid]
                    text.append(f"{cid}\n{item[side]}\n{item.get('image_error', '')}")
                if not row[side + "_ids"]: text.append("(해당 없음)")
        if not text: text = self.analysis.manifest.get("warnings", []) + reanalysis_warnings(self.analysis.manifest)
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
        row.update(section=section, page=page); self.review_dirty = True; self.refresh()

    def match(self):
        if self.busy or not self.analysis: return
        selected = self.indices()
        if not selected or selected != list(range(selected[0], selected[-1] + 1)):
            messagebox.showinfo("전후 매칭 수정", "수정할 항목을 한 개 이상 연속으로 선택하십시오. 여러 항목은 Shift 또는 Ctrl 키로 선택할 수 있습니다.", parent=self.root); return
        if any(not self.rows[index]["included"] for index in selected):
            messagebox.showinfo("전후 매칭 수정", "제외된 항목은 포함으로 복원한 뒤 수정하십시오.", parent=self.root); return
        dialog = MatchingDialog(self.root, self.analysis, [self.rows[index] for index in selected])
        self.root.wait_window(dialog.window)
        if dialog.result is None: return
        try:
            set_matching(self.analysis, self.rows, selected, dialog.result)
        except ValueError as exc:
            messagebox.showerror("매칭 수정 오류", str(exc), parent=self.root); return
        self.review_dirty = True; self.refresh()
        self.status.set("전후 매칭을 수정했습니다. Section·page를 확인하고 ‘검토 저장’으로 별도 저장하십시오.")

    def merge(self):
        try: merge_rows(self.rows, self.indices()); self.review_dirty = True; self.refresh()
        except ValueError as exc: messagebox.showwarning("병합", str(exc))

    def toggle(self):
        for idx in self.indices(): self.rows[idx]["included"] = not self.rows[idx]["included"]
        if self.indices(): self.review_dirty = True
        self.refresh()

    def reset(self):
        if self.analysis and messagebox.askyesno("검토 초기화", "병합·제외·수정 내용을 초기화합니까?"):
            self.rows = initial_rows(self.analysis); self.review_dirty = True; self.refresh()

    def save(self):
        if not self.analysis: return
        path = filedialog.asksaveasfilename(defaultextension=".json", initialfile=self.analysis.path.stem + "_검토.json", filetypes=[("검토 파일", "*.json")])
        if path:
            try: save_review(self.analysis, self.rows, path); self.review_dirty = False; self.status.set(f"검토 저장: {path}")
            except Exception as exc: messagebox.showerror("검토 저장 오류", str(exc))

    def load(self):
        if not self.analysis: return
        if self.review_dirty and not messagebox.askyesno("저장하지 않은 검토", "현재 검토 내용을 저장하지 않고 다른 검토 파일을 불러오겠습니까?", parent=self.root): return
        path = filedialog.askopenfilename(filetypes=[("검토 파일", "*.json")])
        if path:
            try: self.rows = load_review(self.analysis, path); self.review_dirty = False; self.refresh(); self.status.set("검토 내용을 복원했습니다.")
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
        if self.review_dirty and not messagebox.askyesno("저장하지 않은 검토", "검토 내용을 저장하지 않고 종료하겠습니까?", parent=self.root): return
        self.root.destroy()


class MatchingDialog:
    """Reassign existing source blocks; never edit source text or the package."""
    def __init__(self, parent, analysis, rows):
        self.result = None
        # The caller supplies validated canonical rows from the current review.
        self.items = {item["change_id"]: item for item in analysis.manifest["items"]}
        self.pairs = [(list(row["before_ids"]), list(row["after_ids"])) for row in rows]
        self.pool = {side: [cid for row in rows for cid in row[side + "_ids"]] for side in ("before", "after")}
        self.available = {"before": [], "after": []}
        self.window = tk.Toplevel(parent); self.window.title("전후 매칭 확인·수정"); self.window.geometry("1120x820")
        self.window.transient(parent); self.window.grab_set()
        ttk.Label(self.window, text="1. 잘못 연결된 행을 선택해 ‘연결 해제’ → 2. 양쪽 원문을 선택해 ‘선택 내용 연결’ → 3. ‘검토에 적용’", padding=8).pack(fill="x")
        ttk.Label(self.window, text="한쪽만 선택하면 추가/삭제 항목이 됩니다. 여러 내용을 한 행에 연결하려면 Ctrl 또는 Shift로 선택하십시오.", padding=(8, 0, 8, 8)).pack(fill="x")
        current = ttk.Frame(self.window, padding=8); current.pack(fill="both", expand=True)
        self.current = ttk.Treeview(current, columns=("before", "after"), show="headings", selectmode="extended", height=6)
        for side, label in (("before", "현재 변경 전"), ("after", "현재 변경 후")):
            self.current.heading(side, text=label); self.current.column(side, width=490)
        self.current.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(current, orient="vertical", command=self.current.yview); scroll.pack(side="right", fill="y"); self.current.configure(yscrollcommand=scroll.set)
        self.current.bind("<<TreeviewSelect>>", lambda event: self.preview_current())
        actions = ttk.Frame(self.window, padding=(8, 0)); actions.pack(fill="x")
        ttk.Button(actions, text="선택 행 연결 해제", command=self.release).pack(side="left")
        ttk.Button(actions, text="행 위로", command=lambda: self.move(-1)).pack(side="left", padx=6)
        ttk.Button(actions, text="행 아래로", command=lambda: self.move(1)).pack(side="left")
        lists = ttk.Frame(self.window, padding=8); lists.pack(fill="both", expand=True)
        self.lists = {}
        for column, (side, label) in enumerate((("before", "연결할 변경 전 원문"), ("after", "연결할 변경 후 원문"))):
            frame = ttk.Frame(lists); frame.grid(row=0, column=column, sticky="nsew", padx=(0, 8) if column == 0 else (8, 0))
            lists.columnconfigure(column, weight=1); lists.rowconfigure(0, weight=1)
            ttk.Label(frame, text=label).pack(anchor="w")
            box = tk.Listbox(frame, selectmode="extended", exportselection=False, height=7)
            box.pack(side="left", fill="both", expand=True)
            scroll = ttk.Scrollbar(frame, orient="vertical", command=box.yview); scroll.pack(side="right", fill="y"); box.configure(yscrollcommand=scroll.set)
            box.bind("<<ListboxSelect>>", lambda event: self.preview_available())
            self.lists[side] = box
        controls = ttk.Frame(self.window, padding=8); controls.pack(fill="x")
        ttk.Button(controls, text="선택 내용 연결", command=self.connect).pack(side="left")
        self.remaining = tk.StringVar(); ttk.Label(controls, textvariable=self.remaining).pack(side="left", padx=12)
        detail = ttk.Frame(self.window, padding=8); detail.pack(fill="both", expand=True)
        self.detail = tk.Text(detail, wrap="word", state="disabled", height=11); self.detail.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(detail, orient="vertical", command=self.detail.yview); scroll.pack(side="right", fill="y"); self.detail.configure(yscrollcommand=scroll.set)
        footer = ttk.Frame(self.window, padding=8); footer.pack(fill="x")
        ttk.Button(footer, text="검토에 적용", command=self.apply).pack(side="right")
        ttk.Button(footer, text="취소", command=self.window.destroy).pack(side="right", padx=8)
        ttk.Label(footer, text="원본 .bdsg는 보존됩니다. 적용 후 메인 화면에서 ‘검토 저장’을 누르십시오.").pack(side="left")
        self.refresh()

    def summary(self, ids, side):
        if not ids: return "(해당 없음)"
        return " | ".join((self.items[cid][side].replace("\n", " ") or "[표·그림·개체]")[:95] for cid in ids)

    def refresh(self):
        self.current.delete(*self.current.get_children())
        for index, (before, after) in enumerate(self.pairs):
            self.current.insert("", "end", iid=str(index), values=(self.summary(before, "before"), self.summary(after, "after")))
        used = {side: {cid for pair in self.pairs for cid in pair[index]} for index, side in enumerate(("before", "after"))}
        for side in ("before", "after"):
            self.available[side] = [cid for cid in self.pool[side] if cid not in used[side]]
            self.lists[side].delete(0, "end")
            for cid in self.available[side]: self.lists[side].insert("end", cid + " · " + self.summary([cid], side))
        self.remaining.set(f"미연결: 변경 전 {len(self.available['before'])}개 / 변경 후 {len(self.available['after'])}개")

    def show_detail(self, before, after):
        lines = []
        for side, ids, label in (("before", before, "변경 전"), ("after", after, "변경 후")):
            lines.append(f"[{label}]")
            for cid in ids: lines.append(cid + "\n" + (self.items[cid][side] or "[텍스트가 없는 표·그림·개체: 원본에서 확인하십시오.]"))
            if not ids: lines.append("(해당 없음)")
        self.detail.configure(state="normal"); self.detail.delete("1.0", "end"); self.detail.insert("end", "\n\n".join(lines)); self.detail.configure(state="disabled")

    def preview_current(self):
        before, after = [], []
        for index in sorted(int(value) for value in self.current.selection()):
            before.extend(self.pairs[index][0]); after.extend(self.pairs[index][1])
        self.show_detail(before, after)

    def chosen(self, side):
        return [self.available[side][index] for index in self.lists[side].curselection()]

    def preview_available(self): self.show_detail(self.chosen("before"), self.chosen("after"))

    def release(self):
        for index in sorted((int(value) for value in self.current.selection()), reverse=True): self.pairs.pop(index)
        self.refresh(); self.show_detail([], [])

    def connect(self):
        before, after = self.chosen("before"), self.chosen("after")
        if not before and not after:
            messagebox.showinfo("선택 내용 연결", "연결할 원문을 한쪽 이상 선택하십시오.", parent=self.window); return
        self.pairs.append((before, after)); self.refresh(); self.show_detail(before, after)

    def move(self, offset):
        selected = self.current.selection()
        if len(selected) != 1: return
        index = int(selected[0]); destination = index + offset
        if destination < 0 or destination >= len(self.pairs): return
        self.pairs[index], self.pairs[destination] = self.pairs[destination], self.pairs[index]
        self.refresh(); self.current.selection_set(str(destination)); self.current.see(str(destination)); self.preview_current()

    def apply(self):
        if self.available["before"] or self.available["after"]:
            messagebox.showinfo("아직 연결되지 않은 원문", "모든 원문을 연결하십시오. 추가/삭제는 한쪽만 선택해 연결하면 됩니다.", parent=self.window); return
        self.result = copy.deepcopy(self.pairs); self.window.destroy()


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
