# -*- coding: utf-8 -*-
"""
엑셀 취합 자동화
설치:  pip install customtkinter openpyxl pywin32 tkinterdnd2
실행:  python excel_consolidator.py
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import customtkinter as ctk
import openpyxl
import os
import stat
from copy import copy
import win32com.client as win32
import time

# ───────── PDF 도우미와 같은 색상·글꼴 ─────────
BG = "#EEF1F6"          # 바탕색
NAVY = "#1F2A44"        # 왼쪽 메뉴 색
NAVY_L = "#2B3A5C"      # 남색 위의 목록 상자 색
BLUE = "#3B8ED0"        # 포인트(파랑)
GREEN = "#2F9E7A"       # 실행 버튼(초록)
GREEN_H = "#25805F"
GRAY = "#8A94A6"        # 보조 버튼(회색)
GRAY_H = "#6F7A8E"
RED = "#E74C3C"
RED_H = "#C0392B"
TEXT_SUB = "#667085"    # 설명 글씨
PLACEHOLDER = "#A0A8B8" # 입력칸 안의 회색 예시 글씨
LINE = "#D5DAE3"        # 테두리
FONT = "Malgun Gothic"


class ModernDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


class ExcelConsolidatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 취합 자동화")
        self.root.geometry("1180x820")
        self.root.minsize(1000, 720)

        self.filepaths = []

        # ================= [왼쪽 메뉴: PDF 도우미와 같은 남색 사이드바] =================
        side = ctk.CTkFrame(root, width=300, corner_radius=0, fg_color=NAVY)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)

        ctk.CTkLabel(side, text="취합 대상 파일", text_color="white", anchor="w",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(fill="x", padx=16, pady=(28, 10))
        ctk.CTkButton(side, text="＋ 엑셀 파일 추가", command=self.add_files).pack(fill="x", padx=16, pady=(2, 2))
        ctk.CTkButton(side, text="－ 엑셀 파일 제거", command=self.remove_files,
                      fg_color=RED, hover_color=RED_H).pack(fill="x", padx=16, pady=(2, 8))

        sort_frame = ctk.CTkFrame(side, fg_color="transparent")
        sort_frame.pack(fill="x", padx=16, pady=2)
        ctk.CTkButton(sort_frame, text="▲ 오름차순", command=self.sort_asc, width=100,
                      fg_color=GRAY, hover_color=GRAY_H).pack(side="left", expand=True, fill="x", padx=(0, 2))
        ctk.CTkButton(sort_frame, text="▼ 내림차순", command=self.sort_desc, width=100,
                      fg_color=GRAY, hover_color=GRAY_H).pack(side="right", expand=True, fill="x", padx=(2, 0))

        ctk.CTkLabel(side, text="※ 파일을 끌어다 놓아 추가하고,\n   목록에서 끌어서 순서를 바꿀 수 있습니다.",
                     font=ctk.CTkFont(size=11), text_color="#AEB8CC", justify="left",
                     anchor="w").pack(fill="x", padx=16, pady=(10, 4))

        box = ctk.CTkFrame(side, corner_radius=10, fg_color=NAVY_L, border_width=1, border_color="#3D4C70")
        box.pack(fill="both", expand=True, padx=16, pady=(4, 16))
        self.listbox = tk.Listbox(box, selectmode=tk.EXTENDED, bd=0, bg=NAVY_L, fg="white",
                                  highlightthickness=0, activestyle="none",
                                  font=(FONT, 11), selectbackground=BLUE, selectforeground="white")
        lb_scroll = ctk.CTkScrollbar(box, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=lb_scroll.set)
        lb_scroll.pack(side="right", fill="y", padx=(0, 2), pady=6)
        self.listbox.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        self.listbox.bind('<Delete>', lambda e: self.remove_files())
        self.listbox.bind('<BackSpace>', lambda e: self.remove_files())

        self.listbox.bind('<Button-1>', self.on_drag_start)
        self.listbox.bind('<B1-Motion>', self.on_drag_motion)
        self.listbox.bind('<ButtonRelease-1>', self.on_drag_release)

        # ================= [오른쪽 본문] =================
        main = ctk.CTkFrame(root, corner_radius=0, fg_color=BG)
        main.pack(side="left", fill="both", expand=True)

        # 맨 아래 상태 표시줄
        bar = ctk.CTkFrame(main, height=36, fg_color="white", corner_radius=0)
        bar.pack(side="bottom", fill="x")
        self.status = ctk.CTkLabel(bar, text="준비됨", text_color=TEXT_SUB, anchor="w")
        self.status.pack(side="left", padx=16)

        ctk.CTkLabel(main, text="엑셀 취합", anchor="w",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(main, text="여러 엑셀 파일을 목록에 담고 순서를 정해 하나의 엑셀 파일로 합칩니다. "
                                "파일을 창 위로 끌어다 놓아도 추가됩니다.",
                     anchor="w", justify="left", text_color=TEXT_SUB).pack(fill="x", padx=28, pady=(2, 14))

        body = ctk.CTkFrame(main, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=28, pady=(0, 14))

        # ---------- 오른쪽: 설정 카드 + 버튼 + 미리보기 ----------
        right = ctk.CTkFrame(body, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True)

        opt_frame = ctk.CTkFrame(right, corner_radius=12, fg_color="white")
        opt_frame.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(opt_frame, text="취합 설정", font=ctk.CTkFont(size=15, weight="bold")).grid(
            row=0, column=0, columnspan=4, sticky=tk.W, padx=15, pady=(12, 4))

        # ---- 제목행 ----
        self.use_header = tk.BooleanVar(value=True)
        ctk.CTkSwitch(opt_frame, text="제목행", variable=self.use_header,
                      command=self.toggle_title_row).grid(row=1, column=0, sticky=tk.W, padx=15, pady=(8, 0))
        self.title_row = ctk.CTkEntry(opt_frame, width=80, placeholder_text="예: 1",
                                      placeholder_text_color=PLACEHOLDER)
        self.title_row.insert(0, "1")
        self.title_row.grid(row=1, column=1, sticky=tk.W, padx=(0, 5), pady=(8, 0))

        ctk.CTkLabel(opt_frame, text="※ 목록의 첫번째 엑셀은 전부 취합하지만 이후 엑셀은 해당 행 다음부터만 취합합니다.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_SUB, justify="left").grid(
            row=2, column=0, columnspan=4, sticky=tk.W, padx=15, pady=(2, 6))

        # ---- 미취합 행 ----
        self.use_stop = tk.BooleanVar(value=False)
        ctk.CTkSwitch(opt_frame, text="미취합 행", variable=self.use_stop,
                      command=self.toggle_stop_row).grid(row=3, column=0, sticky=tk.W, padx=15, pady=(8, 0))
        self.stop_row = ctk.CTkEntry(opt_frame, width=80, placeholder_text="예: 50",
                                     placeholder_text_color=PLACEHOLDER)
        self.stop_row.grid(row=3, column=1, sticky=tk.W, padx=(0, 5), pady=(8, 0))
        self.stop_row.configure(state="disabled")

        ctk.CTkLabel(opt_frame, text="※ 해당 행 이하는 취합하지 않습니다.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_SUB, justify="left").grid(
            row=4, column=0, columnspan=4, sticky=tk.W, padx=15, pady=(2, 6))

        # ---- 시트 선택 ----
        ctk.CTkLabel(opt_frame, text="시트 선택 방식:").grid(row=5, column=0, sticky=tk.W, padx=15, pady=(8, 0))
        self.sheet_type = tk.StringVar(value="index")
        ctk.CTkRadioButton(opt_frame, text="순번입력", variable=self.sheet_type, value="index",
                           command=self.update_sheet_placeholder).grid(
            row=5, column=1, sticky=tk.W, pady=(8, 0))
        ctk.CTkRadioButton(opt_frame, text="시트명입력", variable=self.sheet_type, value="name",
                           command=self.update_sheet_placeholder).grid(
            row=5, column=2, sticky=tk.W, padx=(5, 0), pady=(8, 0))

        self.sheet_val = ctk.CTkEntry(opt_frame, width=170, placeholder_text="예: 1-3, 5, 8",
                                      placeholder_text_color=PLACEHOLDER)
        self.sheet_val.grid(row=5, column=3, sticky=tk.W, padx=5, pady=(8, 0))
        self.sheet_val.insert(0, "1")

        ctk.CTkLabel(opt_frame, text="※ 다중 시트 취합 시 각 파일의 시트 개수는 동일해야 합니다.",
                     font=ctk.CTkFont(size=11), text_color=TEXT_SUB, justify="left").grid(
            row=6, column=0, columnspan=4, sticky=tk.W, padx=15, pady=(4, 6))

        # ---- 기타 ----
        self.keep_format = tk.BooleanVar(value=True)
        ctk.CTkSwitch(opt_frame, text="원본 서식(셀 병합 등) 유지", variable=self.keep_format).grid(
            row=7, column=0, columnspan=2, sticky=tk.W, padx=15, pady=(6, 14))

        self.add_filename = tk.BooleanVar(value=True)
        ctk.CTkSwitch(opt_frame, text="마지막 열에 원본 파일명 기재", variable=self.add_filename).grid(
            row=7, column=2, columnspan=2, sticky=tk.W, padx=10, pady=(6, 14))

        btn_frame = ctk.CTkFrame(right, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkButton(btn_frame, text="🔍 데이터 미리보기", command=self.preview, height=42,
                      fg_color=GRAY, hover_color=GRAY_H).pack(side="left", padx=(0, 6), expand=True, fill="x")
        ctk.CTkButton(btn_frame, text="🚀 취합 실행하기", command=self.consolidate, height=42,
                      font=ctk.CTkFont(size=15, weight="bold"),
                      fg_color=GREEN, hover_color=GREEN_H).pack(side="right", padx=(6, 0), expand=True, fill="x")

        ctk.CTkLabel(right, text="미리보기", font=ctk.CTkFont(size=13, weight="bold"),
                     anchor="w").pack(fill="x", pady=(2, 4))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="white", foreground="#1F2A44", fieldbackground="white",
                        borderwidth=0, font=(FONT, 9), rowheight=25)
        style.configure("Treeview.Heading", background="#E4E8F0", foreground="#1F2A44",
                        borderwidth=1, relief="flat", font=(FONT, 9, "bold"))
        style.map("Treeview", background=[('selected', BLUE)], foreground=[('selected', 'white')])
        style.map("Treeview.Heading", background=[('active', '#D5DAE3')])

        tree_card = ctk.CTkFrame(right, corner_radius=12, fg_color="white")
        tree_card.pack(fill="both", expand=True)
        tree_frame = tk.Frame(tree_card, bg="white")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(tree_frame, yscrollcommand=self.tree_scroll_y.set,
                                 xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        # 화면 구성이 끝난 뒤, 모든 위젯을 '파일 놓는 곳'으로 등록
        self.enable_drop(self.root)

    # ================= [드래그앤드롭 등록] =================
    def enable_drop(self, widget):
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', self.drop_files_dnd)
        except Exception:
            pass
        for child in widget.winfo_children():
            self.enable_drop(child)

    def set_status(self, text):
        self.status.configure(text=text)
        self.root.update_idletasks()

    # ================= [복구 로직] =================
    def repair_excel_with_win32(self, filepath):
        excel = None
        wb = None
        new_filepath = ""
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.AutomationSecurity = 3

            wb = excel.Workbooks.Open(os.path.abspath(filepath), UpdateLinks=0, ReadOnly=True, CorruptLoad=1)

            dir_name, file_name = os.path.split(filepath)
            name, ext = os.path.splitext(file_name)
            new_filepath = os.path.join(dir_name, name + "_복구됨.xlsx")

            if os.path.exists(new_filepath):
                try: os.remove(new_filepath)
                except Exception: pass

            wb.SaveAs(os.path.abspath(new_filepath), FileFormat=51, ConflictResolution=2)

        except Exception as e:
            raise Exception(f"엑셀 파일 자동 복구 실패 (SaveAs 에러): {e}")
        finally:
            if wb:
                try: wb.Close(False)
                except: pass
                del wb
            if excel:
                try: excel.Quit()
                except: pass
                del excel

        for _ in range(15):
            if os.path.exists(new_filepath):
                try:
                    with open(new_filepath, 'rb') as f: pass
                    break
                except Exception:
                    time.sleep(0.2)
            else:
                time.sleep(0.2)

        return new_filepath

    def safe_load_workbook(self, filepath, data_only=False):
        try:
            wb = openpyxl.load_workbook(filepath, data_only=data_only)
            return wb, filepath
        except Exception:
            pass

        dir_name, file_name = os.path.split(filepath)
        name, ext = os.path.splitext(file_name)
        reusable_path = os.path.join(dir_name, name + "_복구됨.xlsx")

        if os.path.exists(reusable_path):
            try:
                wb = openpyxl.load_workbook(reusable_path, data_only=data_only)
                return wb, reusable_path
            except Exception:
                pass

        recovered_path = self.repair_excel_with_win32(filepath)
        wb = openpyxl.load_workbook(recovered_path, data_only=data_only)
        return wb, recovered_path

    # ================= [기능 메서드] =================
    def unlock_file(self, path):
        try:
            if os.path.exists(path):
                os.chmod(path, stat.S_IWRITE)
        except Exception: pass

    def toggle_title_row(self):
        self.title_row.configure(state="normal" if self.use_header.get() else "disabled")

    def toggle_stop_row(self):
        self.stop_row.configure(state="normal" if self.use_stop.get() else "disabled")

    def update_sheet_placeholder(self):
        """순번입력/시트명입력을 바꾸면 입력칸을 비우고, 회색 예시 글씨를 보여줍니다."""
        text = "예: 1-3, 5, 8" if self.sheet_type.get() == "index" else "예: 1월, 2월, 3월"
        try:
            self.sheet_val.delete(0, "end")                      # 입력된 값(기본값 1 등) 지우기
            self.sheet_val.configure(placeholder_text=text)      # 예시 글씨 바꾸기
            self.sheet_val._activate_placeholder()               # 비어 있으니 예시 글씨 다시 표시
        except Exception:
            pass

    def get_row_settings(self):
        """제목행 / 미취합 행 입력값을 검사해서 (제목행 수, 미취합 행 번호)를 돌려줍니다.
        제목행을 쓰지 않으면 0, 미취합 행을 쓰지 않으면 None 입니다."""
        title_limit = 0
        if self.use_header.get():
            t = self.title_row.get().strip()
            if not t.isdigit() or int(t) < 1:
                raise ValueError("제목행은 1 이상의 숫자로 입력해 주세요.")
            title_limit = int(t)

        stop_row = None
        if self.use_stop.get():
            s = self.stop_row.get().strip()
            if not s.isdigit() or int(s) < 1:
                raise ValueError("미취합 행은 1 이상의 숫자로 입력해 주세요.")
            stop_row = int(s)
            if stop_row <= title_limit:
                raise ValueError("미취합 행은 제목행보다 큰 숫자여야 합니다.")
        return title_limit, stop_row

    def _last_data_row(self, ws, start_row, stop_row):
        """start_row부터 '값이 실제로 있는 마지막 행' 번호를 돌려줍니다.
        (중간에 빈 행이 있어도 그 아래에 값이 있으면 거기까지 포함,
         맨 아래의 빈 행들만 제외. 미취합 행 이하는 보지 않음)
        값이 하나도 없으면 start_row - 1 을 돌려줍니다."""
        limit = ws.max_row if stop_row is None else min(ws.max_row, stop_row - 1)
        last = start_row - 1
        if limit < start_row:
            return last
        r = start_row
        for row in ws.iter_rows(min_row=start_row, max_row=limit, values_only=True):
            if any(v is not None and str(v).strip() != "" for v in row):
                last = r
            r += 1
        return last

    def add_files(self):
        files = filedialog.askopenfilenames(title="엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
        for f in files:
            if os.path.basename(f).startswith("~$"): continue
            self.unlock_file(f)
            if f not in self.filepaths:
                self.filepaths.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))

    def remove_files(self):
        selected = self.listbox.curselection()
        if not selected: return
        for i in reversed(selected):
            self.listbox.delete(i)
            del self.filepaths[i]

    def drop_files_dnd(self, event):
        files = self.root.tk.splitlist(event.data)
        for f in files:
            if os.path.basename(f).startswith("~$"): continue
            if f.lower().endswith(('.xlsx', '.xls')):
                self.unlock_file(f)
                if f not in self.filepaths:
                    self.filepaths.append(f)
                    self.listbox.insert(tk.END, os.path.basename(f))
        return event.action

    def on_drag_start(self, event):
        self.drag_start_y = event.y
        self.is_dragging = False
        index = self.listbox.nearest(event.y)

        bbox = self.listbox.bbox(index)
        if bbox and (event.y < bbox[1] or event.y > bbox[1] + bbox[3]):
            self.listbox.selection_clear(0, tk.END)
            self.dragged_indices = ()
            return

        is_shift = event.state & 0x0001
        is_ctrl = event.state & 0x0004
        if is_shift or is_ctrl:
            self.root.after(10, self._record_drag_start, event.y)
            return

        if len(self.listbox.curselection()) > 1 and index in self.listbox.curselection():
            self.drag_start_index = index
            self.dragged_indices = self.listbox.curselection()
            return "break"
        else:
            self.root.after(10, self._record_drag_start, event.y)

    def _record_drag_start(self, y):
        self.drag_start_index = self.listbox.nearest(y)
        self.dragged_indices = self.listbox.curselection()

    def on_drag_motion(self, event):
        if abs(event.y - getattr(self, 'drag_start_y', event.y)) > 3:
            self.is_dragging = True

        if not getattr(self, 'is_dragging', False): return "break"
        if not getattr(self, 'dragged_indices', None): return "break"

        target_index = self.listbox.nearest(event.y)
        if target_index < 0 or target_index >= self.listbox.size(): return "break"
        if target_index == getattr(self, 'drag_start_index', -1): return "break"

        step = 1 if target_index > getattr(self, 'drag_start_index', -1) else -1
        new_indices = [i + step for i in self.dragged_indices]

        if min(new_indices) < 0 or max(new_indices) >= self.listbox.size():
            return "break"

        if step == 1:
            for i in reversed(self.dragged_indices):
                self.filepaths[i], self.filepaths[i+1] = self.filepaths[i+1], self.filepaths[i]
        else:
            for i in self.dragged_indices:
                self.filepaths[i], self.filepaths[i-1] = self.filepaths[i-1], self.filepaths[i]

        self.listbox.delete(0, tk.END)
        for path in self.filepaths:
            self.listbox.insert(tk.END, os.path.basename(path))

        for i in new_indices:
            self.listbox.selection_set(i)

        self.dragged_indices = new_indices
        self.drag_start_index += step
        return "break"

    def on_drag_release(self, event):
        if not getattr(self, 'is_dragging', False):
            is_shift = event.state & 0x0001
            is_ctrl = event.state & 0x0004
            if not (is_shift or is_ctrl):
                index = self.listbox.nearest(event.y)
                bbox = self.listbox.bbox(index)
                if bbox and bbox[1] <= event.y <= bbox[1] + bbox[3]:
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(index)
                    self._record_drag_start(event.y)

    def sort_asc(self):
        combined = list(zip(self.filepaths, self.listbox.get(0, tk.END)))
        combined.sort(key=lambda x: x[1])
        self._update_listbox(combined)

    def sort_desc(self):
        combined = list(zip(self.filepaths, self.listbox.get(0, tk.END)))
        combined.sort(key=lambda x: x[1], reverse=True)
        self._update_listbox(combined)

    def _update_listbox(self, combined):
        self.listbox.delete(0, tk.END)
        self.filepaths.clear()
        for path, name in combined:
            self.filepaths.append(path)
            self.listbox.insert(tk.END, name)

    # ================= [시트 매핑 로직] =================
    def validate_and_map_sheets(self, files):
        if not files: return [], {}

        stype = self.sheet_type.get()
        sval = self.sheet_val.get().strip()

        if not sval:
            if stype == "index":
                raise ValueError("취합할 시트 순번을 입력해 주세요.  (예: 1-3, 5, 8)")
            else:
                raise ValueError("취합할 시트명을 입력해 주세요.  (예: 1월, 2월, 3월)")

        req_indices = []
        req_names = []

        if stype == "index":
            try:
                for part in sval.split(','):
                    part = part.strip()
                    if not part: continue
                    if '-' in part:
                        start, end = part.split('-')
                        req_indices.extend(range(int(start)-1, int(end)))
                    else:
                        req_indices.append(int(part)-1)
            except ValueError:
                raise ValueError("시트 순번 입력이 올바르지 않습니다.  (예: 1-3, 5, 8)")
            req_indices = list(dict.fromkeys(req_indices))
        else:
            req_names = [n.strip() for n in sval.split(',') if n.strip()]
            req_names = list(dict.fromkeys(req_names))

        is_single = (stype == "index" and len(req_indices) == 1) or (stype == "name" and len(req_names) == 1)

        def get_snames(f):
            try:
                wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
                names = wb.sheetnames
                wb.close()
            except Exception:
                wb, _ = self.safe_load_workbook(f, data_only=True)
                names = wb.sheetnames
                wb.close()
            return names

        base_names = get_snames(files[0])
        base_count = len(base_names)
        mapped_indices = []
        output_sheet_names = []

        if stype == "index":
            for idx in req_indices:
                if 0 <= idx < base_count:
                    mapped_indices.append(idx)
                    output_sheet_names.append(base_names[idx])
                else:
                    raise ValueError(f"첫 번째 파일에 시트 순번({idx+1}번)이 존재하지 않습니다. (총: {base_count}개)")
        else:
            for name in req_names:
                if name in base_names:
                    mapped_indices.append(base_names.index(name))
                    output_sheet_names.append(name)
                else:
                    raise ValueError(f"첫 번째 파일에 '{name}' 시트가 존재하지 않습니다.")

        if not mapped_indices:
            raise ValueError("선택된 시트가 없습니다.")

        file_targets = {}
        for f in files:
            names = get_snames(f)
            file_targets[f] = []

            if is_single:
                if stype == "index":
                    idx = mapped_indices[0]
                    if idx < len(names): file_targets[f].append(names[idx])
                    else: raise ValueError(f"'{os.path.basename(f)}' 파일에 {idx+1}번째 시트가 없습니다.")
                else:
                    name = req_names[0]
                    if name in names: file_targets[f].append(name)
                    else: raise ValueError(f"'{os.path.basename(f)}' 파일에 '{name}' 시트가 존재하지 않습니다.")
            else:
                if len(names) != base_count:
                    raise ValueError(f"다중 시트 취합 시 모든 파일의 총 시트 개수가 일치해야 합니다.\n"
                                     f"기준 파일({os.path.basename(files[0])}): {base_count}개\n"
                                     f"대상 파일({os.path.basename(f)}): {len(names)}개")

                for idx, out_name in zip(mapped_indices, output_sheet_names):
                    if stype == "index":
                        file_targets[f].append(names[idx])
                    else:
                        if out_name in names: file_targets[f].append(out_name)
                        else: file_targets[f].append(names[idx])

        return output_sheet_names, file_targets

    # ================= [미리보기 로직] =================
    def get_text_width(self, text):
        if text is None: return 20
        text_str = str(text)
        length = sum(1.8 if ord(c) > 127 else 1 for c in text_str)
        return int(length * 8) + 15

    def preview(self):
        self.tree.delete(*self.tree.get_children())
        if not self.filepaths:
            messagebox.showwarning("안내", "먼저 취합할 엑셀 파일을 추가해 주세요.")
            return

        try:
            title_limit, stop_row = self.get_row_settings()
        except ValueError as e:
            messagebox.showerror("입력 오류", str(e))
            return

        preview_files = self.filepaths[:3]

        try:
            output_sheet_names, file_targets = self.validate_and_map_sheets(preview_files)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            return

        self.set_status("미리보기 만드는 중...")
        try:
            first_file = preview_files[0]
            self.unlock_file(first_file)
            wb_first, _ = self.safe_load_workbook(first_file, data_only=True)
            sample_ws = wb_first[file_targets[first_file][0]]
            max_col = sample_ws.max_column

            cols = ["행", "시트명"] + [str(i) for i in range(1, max_col + 1)]
            if self.add_filename.get():
                cols.append("원본파일명")

            self.tree["columns"] = cols
            self.tree["show"] = "headings"

            col_widths = {col: self.get_text_width(col) for col in cols}
            col_widths["행"] = 40

            wb_first.close()
            target_len = len(cols) - (1 if self.add_filename.get() else 0)

            for file_idx, file in enumerate(preview_files):
                self.unlock_file(file)
                wb, _ = self.safe_load_workbook(file, data_only=True)

                if file_idx > 0:
                    sep_vals = ["", f"=== [파일: {os.path.basename(file)}] ==="] + [""]*(len(cols)-2)
                    self.tree.insert("", "end", values=sep_vals, tags=('file_sep',))

                for out_name, actual_sname in zip(output_sheet_names, file_targets[file]):
                    ws = wb[actual_sname]

                    sep_vals = ["", f"[{out_name}]"] + [""]*(len(cols)-2)
                    self.tree.insert("", "end", values=sep_vals, tags=('sheet_sep',))

                    if self.use_header.get():
                        for r_idx in range(1, title_limit + 1):
                            row_data = list(ws.iter_rows(min_row=r_idx, max_row=r_idx))
                            if row_data:
                                row_vals = [str(r_idx), out_name] + [cell.value if cell.value is not None else "" for cell in row_data[0]]
                            else:
                                row_vals = [str(r_idx), out_name] + [""] * ws.max_column

                            if len(row_vals) < target_len:
                                row_vals += [""] * (target_len - len(row_vals))
                            row_vals = row_vals[:target_len]

                            if self.add_filename.get() and r_idx == title_limit:
                                row_vals.append("원본 파일명")
                            elif self.add_filename.get():
                                row_vals.append("")

                            for col_name, val in zip(cols, row_vals):
                                col_widths[col_name] = max(col_widths[col_name], self.get_text_width(val))

                            self.tree.insert("", "end", values=row_vals)

                    # 제목행 다음 행부터, 값이 있는 마지막 행까지 중 앞의 2행만 보여줌
                    # (중간의 빈 행도 취합 대상이므로 미리보기에도 그대로 표시)
                    start_row = title_limit + 1
                    last_row = self._last_data_row(ws, start_row, stop_row)
                    count = 0

                    if last_row >= start_row:
                        for r_idx, row in enumerate(
                                ws.iter_rows(min_row=start_row, max_row=last_row, values_only=True),
                                start=start_row):
                            is_empty = all(val is None or str(val).strip() == "" for val in row)

                            row_vals = [str(r_idx), out_name] + [val if val is not None else "" for val in row]

                            if len(row_vals) < target_len:
                                row_vals += [""] * (target_len - len(row_vals))
                            row_vals = row_vals[:target_len]

                            if self.add_filename.get():
                                row_vals.append("" if is_empty else os.path.basename(file))

                            for col_name, val in zip(cols, row_vals):
                                col_widths[col_name] = max(col_widths[col_name], self.get_text_width(val))

                            self.tree.insert("", "end", values=row_vals)
                            count += 1
                            if count >= 2:
                                break
                wb.close()

            # 밝은 화면에 어울리는 구분줄 색상
            self.tree.tag_configure('sheet_sep', background="#DCE9F8", foreground="#1F2A44")
            self.tree.tag_configure('file_sep', background="#DDF1E9", foreground="#1F2A44")

            for col in cols:
                self.tree.heading(col, text=col, anchor="w")
                final_width = min(max(col_widths[col], 40), 400)
                self.tree.column(col, anchor="w", width=final_width)

        except Exception as e:
            messagebox.showerror("미리보기 오류", f"미리보기 생성 중 오류가 발생했습니다:\n{e}")
        finally:
            self.set_status("준비됨")

    # 서식 복사 보조 메서드
    def _copy_style(self, src_cell, target_cell):
        if src_cell.has_style:
            target_cell.font = copy(src_cell.font)
            target_cell.border = copy(src_cell.border)
            target_cell.fill = copy(src_cell.fill)
            target_cell.number_format = copy(src_cell.number_format)
            target_cell.protection = copy(src_cell.protection)
            target_cell.alignment = copy(src_cell.alignment)

    # ================= [취합 로직] =================
    def consolidate(self):
        if not self.filepaths:
            messagebox.showwarning("안내", "취합할 파일이 없습니다.")
            return

        try:
            title_limit, stop_row = self.get_row_settings()
        except ValueError as e:
            messagebox.showerror("입력 오류", str(e))
            return

        try:
            output_sheet_names, file_targets = self.validate_and_map_sheets(self.filepaths)
        except Exception as e:
            messagebox.showerror("구조 검증 오류", str(e))
            return

        try:
            out_wb = openpyxl.Workbook()
            out_wb.remove(out_wb.active)

            out_sheets = {}
            sheet_states = {}

            for s_name in output_sheet_names:
                ws = out_wb.create_sheet(title=s_name[:31])
                out_sheets[s_name] = ws

                # A1셀부터 출력
                sheet_states[s_name] = {
                    'current_row': 1,
                    'header_written': False,
                    'start_c': 1
                }

            total = len(self.filepaths)
            for k, file in enumerate(self.filepaths, start=1):
                self.set_status(f"취합 중  {k}/{total}개 ({int(k / total * 100)}%)  -  {os.path.basename(file)}")
                try:
                    self.unlock_file(file)
                    wb, actual_file = self.safe_load_workbook(file, data_only=not self.keep_format.get())

                    for out_name, actual_sname in zip(output_sheet_names, file_targets[file]):
                        ws = wb[actual_sname]
                        out_ws = out_sheets[out_name]
                        state = sheet_states[out_name]

                        current_row = state['current_row']
                        start_c = state['start_c']
                        row_mapping = {}

                        # ① 제목행: 목록의 첫번째 엑셀에서만 한 번 가져옴
                        if self.use_header.get() and not state['header_written']:
                            for r_idx in range(1, title_limit + 1):
                                row_mapping[r_idx] = current_row
                                row_data = list(ws.iter_rows(min_row=r_idx, max_row=r_idx))
                                if row_data:
                                    for c_idx, cell in enumerate(row_data[0], start=start_c):
                                        new_cell = out_ws.cell(row=current_row, column=c_idx, value=cell.value)
                                        self._copy_style(cell, new_cell)

                                    if self.add_filename.get() and r_idx == title_limit:
                                        out_ws.cell(row=current_row, column=start_c + ws.max_column, value="원본 파일명")
                                current_row += 1
                            state['header_written'] = True

                        # ② 데이터 행: 제목행 다음 행부터, 값이 있는 마지막 행까지
                        #    (중간의 빈 행은 지우지 않고 그대로 가져옴 / 미취합 행 이하는 제외)
                        data_start_row = title_limit + 1
                        last_row = self._last_data_row(ws, data_start_row, stop_row)

                        if last_row >= data_start_row:
                            source_r_idx = data_start_row
                            for row in ws.iter_rows(min_row=data_start_row, max_row=last_row):
                                is_empty = all(cell.value is None or str(cell.value).strip() == "" for cell in row)

                                row_mapping[source_r_idx] = current_row
                                for c_idx, cell in enumerate(row, start=start_c):
                                    new_cell = out_ws.cell(row=current_row, column=c_idx, value=cell.value)
                                    self._copy_style(cell, new_cell)
                                # 빈 행에는 파일명을 적지 않아 빈 행 그대로 유지
                                if self.add_filename.get() and not is_empty:
                                    out_ws.cell(row=current_row, column=start_c + ws.max_column, value=os.path.basename(file))
                                current_row += 1
                                source_r_idx += 1

                        if self.keep_format.get():
                            for m_range in ws.merged_cells.ranges:
                                min_col, min_row, max_col, max_row = m_range.bounds
                                if min_row in row_mapping and max_row in row_mapping:
                                    out_ws.merge_cells(
                                        start_row=row_mapping[min_row],
                                        start_column=min_col + (start_c - 1),
                                        end_row=row_mapping[max_row],
                                        end_column=max_col + (start_c - 1)
                                    )

                        state['current_row'] = current_row

                    wb.close()
                except Exception as e:
                    messagebox.showerror("취합 오류", f"파일을 취합하는 중 오류가 발생했습니다.\n{e}")
                    return

            self.set_status("저장 위치를 선택해 주세요...")
            save_path = filedialog.asksaveasfilename(
                title="취합본 저장 위치 선택",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="취합완료.xlsx"
            )

            if save_path:
                self.set_status("저장 중...")
                out_wb.save(save_path)
                messagebox.showinfo("완료", "엑셀 파일 취합이 성공적으로 완료되었습니다!")

        except Exception as e:
            messagebox.showerror("오류", f"프로세스 실행 중 오류가 발생했습니다: {e}")
        finally:
            self.set_status("준비됨")


if __name__ == "__main__":
    try:
        ctk.ThemeManager.theme["CTkFont"]["family"] = FONT
    except Exception:
        pass
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    root = ModernDnD()
    app = ExcelConsolidatorApp(root)
    root.mainloop()