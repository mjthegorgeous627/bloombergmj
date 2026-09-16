"""Button launcher for SAP and Bloomberg Portal automation."""

import os
import json
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from excel_portal_lookup import find_excel_order, _line_kind


BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "automation.log"
POD_PENDING_FILE = BASE_DIR / "pod_pending.json"
STOP_FILE = BASE_DIR / "stop_automation.flag"
RUN_NOW_FILE = BASE_DIR / "run_now.flag"
SESSION_FLAGS = {
    "VL06O": BASE_DIR / "run_now_0.flag",
    "VL10G": BASE_DIR / "run_now_1.flag",
    "ZRMA RLKR": BASE_DIR / "run_now_2.flag",
    "ZRMA Q2": BASE_DIR / "run_now_3.flag",
}

CREATE_NEW_CONSOLE = 0x00000010 if os.name == "nt" else 0


class Launcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MJSuh Automation Launcher")
        self.geometry("1280x780")
        self.minsize(1100, 680)
        self.configure(bg="#f4f6f8")

        self.started_process = None
        self.watchdog_process = None
        self.log_visible = True
        self.sap_tools_visible = False
        self.sap_more_visible = False
        self.portal_advanced_visible = False
        self.portal_detail_visible = False
        self.pod_options_visible = False

        self.status_var = tk.StringVar(value="대기 중")
        self.order_var = tk.StringVar()
        self.new_session_tcode_var = tk.StringVar()

        self.portal_order_var = tk.StringVar()
        self.portal_delivery_var = tk.StringVar()
        self.portal_material_var = tk.StringVar()
        self.portal_qty_var = tk.StringVar(value="1")
        self.portal_serial_var = tk.StringVar()
        self.portal_skip_qr_var = tk.BooleanVar(value=False)
        self.portal_material_only_var = tk.BooleanVar(value=False)

        self.pod_signed_var = tk.StringVar()
        self.pod_datetime_var = tk.StringVar()
        self.pod_remarks_var = tk.StringVar()
        self.pod_pickup_done_var = tk.BooleanVar(value=False)

        self.print_sheet_var = tk.StringVar(value="4-2")
        self.print_date_var = tk.StringVar()
        self.print_afternoon_var = tk.BooleanVar(value=False)

        self._build_ui()
        self.refresh_log()
        self.after(3000, self._tick)

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f6f8")
        style.configure("TLabelframe", background="#f4f6f8")
        style.configure("TLabelframe.Label", background="#f4f6f8", font=("Malgun Gothic", 10, "bold"))
        style.configure("TButton", font=("Malgun Gothic", 10), padding=7)
        style.configure("TLabel", background="#f4f6f8", font=("Malgun Gothic", 9))
        style.configure("TEntry", font=("Malgun Gothic", 10))
        style.configure("Primary.TButton", font=("Malgun Gothic", 11, "bold"), padding=10)
        style.configure("Small.TButton", font=("Malgun Gothic", 9), padding=4)

        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)

        top = ttk.Frame(root)
        top.pack(fill="x")
        ttk.Label(top, text="MJSuh Automation", font=("Malgun Gothic", 17, "bold")).pack(side="left")
        ttk.Label(top, textvariable=self.status_var, font=("Malgun Gothic", 10)).pack(side="right")

        main = ttk.Frame(root)
        main.pack(fill="both", expand=True, pady=(12, 0))

        left = ttk.Frame(main, width=260)
        left.pack(side="left", fill="y", padx=(0, 12))
        left.pack_propagate(False)

        center = ttk.Frame(main, width=440)
        center.pack(side="left", fill="y", padx=(0, 12))
        center.pack_propagate(False)

        self.log_box = ttk.LabelFrame(main, text="최신 로그")
        self.log_box.pack(side="left", fill="both", expand=True)

        self._build_sap_column(left)
        self._build_portal_column(center)
        self._build_log_column()

    def _labeled_entry(self, parent, label, var):
        ttk.Label(parent, text=label).pack(anchor="w", padx=10, pady=(6, 0))
        entry = ttk.Entry(parent, textvariable=var)
        entry.pack(fill="x", padx=10, pady=(1, 3))
        return entry

    def _hint(self, parent, text, padx=10):
        ttk.Label(parent, text=text, wraplength=390, foreground="#475569").pack(
            anchor="w", fill="x", padx=padx, pady=(4, 4)
        )

    def _tiny_toggle(self, parent, command):
        return ttk.Button(parent, text="▼", width=3, style="Small.TButton", command=command)

    def _build_sap_column(self, parent):
        auto_box = ttk.LabelFrame(parent, text="SAP 자동화")
        auto_box.pack(fill="x", pady=(0, 10))
        ttk.Button(auto_box, text="SAP 시작 & 자동루프", style="Primary.TButton", command=self.start_startup).pack(fill="x", padx=10, pady=(10, 7))
        ttk.Button(auto_box, text="전체 즉시 조회", style="Primary.TButton", command=self.trigger_all).pack(fill="x", padx=10, pady=(0, 7))

        self.sap_tools_toggle = self._tiny_toggle(auto_box, self.toggle_sap_tools)
        self.sap_tools_toggle.pack(anchor="center", pady=(0, 6))
        self.sap_tools_box = ttk.Frame(auto_box)
        ttk.Button(self.sap_tools_box, text="자동화 중지", style="Small.TButton", command=self.stop_automation).pack(fill="x", padx=24, pady=(0, 3))
        ttk.Button(self.sap_tools_box, text="한 번만 전체 조회", style="Small.TButton", command=self.run_once).pack(fill="x", padx=24, pady=(0, 3))
        ttk.Button(self.sap_tools_box, text="자동 루프만 시작", style="Small.TButton", command=self.start_main_loop).pack(fill="x", padx=24, pady=(0, 3))
        ttk.Button(self.sap_tools_box, text="워치독 시작 (죽으면 자동재시작)", style="Small.TButton", command=self.start_watchdog).pack(fill="x", padx=24, pady=(0, 3))
        ttk.Button(self.sap_tools_box, text="워치독 중지", style="Small.TButton", command=self.stop_watchdog).pack(fill="x", padx=24, pady=(0, 6))

        order_box = ttk.LabelFrame(parent, text="오더 엑셀 반영")
        order_box.pack(fill="x", pady=(0, 10))
        entry = self._labeled_entry(order_box, "오더번호", self.order_var)
        entry.bind("<Return>", lambda _e: self.run_order())
        ttk.Button(order_box, text="오더 엑셀 반영", style="Primary.TButton", command=self.run_order).pack(fill="x", padx=10, pady=(5, 10))

        dashboard_box = ttk.LabelFrame(parent, text="Vendor Dashboard")
        dashboard_box.pack(fill="x", pady=(0, 10))
        self._hint(dashboard_box, "배송장 엑셀에서 이번 달 Bloomberg Dashboard 복붙용 시트를 만듭니다.")
        ttk.Button(dashboard_box, text="Dashboard 작성/갱신", command=self.run_vendor_dashboard).pack(fill="x", padx=10, pady=(5, 10))

        new_session_box = ttk.LabelFrame(parent, text="새 SAP 창 열기")
        new_session_box.pack(fill="x", pady=(0, 10))
        self._hint(new_session_box, "NWBC '+' 새 탭 버튼이 자동화 실행 중 멈추는 버그가 있어, 대신 이 버튼으로 여세요.")
        self._labeled_entry(new_session_box, "T-code (비우면 빈 세션만)", self.new_session_tcode_var)
        ttk.Button(new_session_box, text="새 SAP 세션 열기", style="Primary.TButton", command=self.open_new_sap_session).pack(fill="x", padx=10, pady=(5, 10))

        self.sap_more_toggle = ttk.Button(parent, text="특정 창 조회 / 엑셀 인쇄 ▼", style="Small.TButton", command=self.toggle_sap_more)
        self.sap_more_toggle.pack(fill="x", pady=(0, 6))
        self.sap_more_box = ttk.Frame(parent)

        session_box = ttk.LabelFrame(self.sap_more_box, text="특정 창만 조회")
        session_box.pack(fill="x", pady=(0, 8))
        for label in SESSION_FLAGS:
            ttk.Button(session_box, text=label, style="Small.TButton", command=lambda name=label: self.trigger_session(name)).pack(fill="x", padx=10, pady=3)
        ttk.Frame(session_box).pack(pady=5)

        print_box = ttk.LabelFrame(self.sap_more_box, text="엑셀 인쇄")
        print_box.pack(fill="x")
        self._labeled_entry(print_box, "시트", self.print_sheet_var)
        self._labeled_entry(print_box, "날짜", self.print_date_var)
        ttk.Checkbutton(print_box, text="오후", variable=self.print_afternoon_var).pack(anchor="w", padx=10, pady=4)
        ttk.Button(print_box, text="시트 인쇄", command=self.run_print).pack(fill="x", padx=10, pady=(4, 10))

    def _build_portal_column(self, parent):
        ttk.Label(parent, text="Portal 입력", font=("Malgun Gothic", 12, "bold")).pack(anchor="w", pady=(0, 6))

        serial_box = ttk.LabelFrame(parent, text="1. Serial 등록 & QR 인쇄")
        serial_box.pack(fill="x", pady=(0, 10))
        self._hint(serial_box, "오더번호와 OBD를 알면 둘 다 입력하세요. OBD를 비우면 오늘자 엑셀 시트에서 Delivery#, Material, Serial을 찾습니다.")
        self._labeled_entry(serial_box, "오더번호", self.portal_order_var)
        self._labeled_entry(serial_box, "OBD / Delivery# (알면 입력)", self.portal_delivery_var)
        self._labeled_entry(serial_box, "Serial Number (엑셀에 없을 때만)", self.portal_serial_var)
        ttk.Button(serial_box, text="Portal 로그인 열기 (자동)", command=self.open_portal_login).pack(fill="x", padx=10, pady=(6, 4))
        ttk.Button(serial_box, text="Portal 로그인 (수동)", command=self.open_portal_login_manual).pack(fill="x", padx=10, pady=(0, 4))
        ttk.Checkbutton(serial_box, text="QR 인쇄 안함", variable=self.portal_skip_qr_var).pack(anchor="w", padx=10, pady=(4, 0))
        ttk.Checkbutton(serial_box, text="Material only / no serial", variable=self.portal_material_only_var).pack(anchor="w", padx=10, pady=(2, 0))
        ttk.Button(serial_box, text="Serial 등록 & QR 인쇄", style="Primary.TButton", command=self.portal_ship_and_print).pack(fill="x", padx=10, pady=(6, 5))
        self.portal_detail_toggle = self._tiny_toggle(serial_box, self.toggle_portal_detail)
        self.portal_detail_toggle.pack(anchor="center", pady=(0, 8))
        self.portal_detail_box = ttk.LabelFrame(serial_box, text="자동 조회 실패 시 직접 입력")
        self._labeled_entry(self.portal_detail_box, "Quantity", self.portal_qty_var)
        self._labeled_entry(self.portal_detail_box, "Material #", self.portal_material_var)

        qr_box = ttk.LabelFrame(parent, text="2. QR 수동 인쇄")
        qr_box.pack(fill="x", pady=(0, 10))
        self._hint(qr_box, "방금 받은 QR 라벨을 다시 인쇄할 때 사용합니다.")
        self._labeled_entry(qr_box, "Delivery# / OBD (특정 라벨 인쇄 시)", self.portal_delivery_var)
        ttk.Button(qr_box, text="최근 ZPL 수동 인쇄", command=self.run_latest_zpl).pack(fill="x", padx=10, pady=(6, 4))
        ttk.Button(qr_box, text="Delivery#로 QR 인쇄", command=self.portal_labels_print).pack(fill="x", padx=10, pady=(0, 10))

        pod_flow = ttk.LabelFrame(parent, text="3. 수동 배송 완료 처리")
        pod_flow.pack(fill="x", pady=(0, 10))
        self._hint(pod_flow, "오더번호를 넣고 미리 채우면 최신 로그에 POD 내용이 뜹니다. 확인 후 최종 저장하세요.")
        ttk.Checkbutton(pod_flow, text="ZRX 회수 완료", variable=self.pod_pickup_done_var).pack(anchor="w", padx=10, pady=(4, 0))
        ttk.Button(pod_flow, text="POD 확인/저장 선택", command=self.run_pod_fill).pack(fill="x", padx=10, pady=(6, 4))
        ttk.Button(pod_flow, text="POD 최종 저장", style="Primary.TButton", command=self.run_pod_update).pack(fill="x", padx=10, pady=(0, 10))

        ttk.Button(parent, text="고급/문제해결 버튼 펼치기", command=self.toggle_portal_advanced).pack(fill="x", pady=(0, 6))
        self.portal_advanced_toggle = parent.winfo_children()[-1]
        self.portal_advanced_box = ttk.LabelFrame(parent, text="고급/문제해결")
        self._build_portal_advanced()

        ttk.Button(parent, text="POD 옵션 펼치기", command=self.toggle_pod_options).pack(fill="x", pady=(0, 6))
        self.pod_options_toggle = parent.winfo_children()[-1]
        self.pod_box = ttk.LabelFrame(parent, text="POD 옵션")
        self._build_pod_options()

    def _build_portal_advanced(self):
        box = self.portal_advanced_box
        ttk.Button(box, text="OBD 확인/자동채움", command=self.fill_portal_from_excel).pack(fill="x", padx=10, pady=(10, 4))
        ttk.Button(box, text="오더 찾아 열기", command=self.portal_open_delivery).pack(fill="x", padx=10, pady=4)
        ttk.Button(box, text="Serial 등록 + Run ShipERP", command=self.portal_register_serial).pack(fill="x", padx=10, pady=4)
        ttk.Button(box, text="Packing 준비만", command=self.portal_pack_prepare).pack(fill="x", padx=10, pady=4)
        ttk.Button(box, text="Packing Post", command=self.portal_pack_post).pack(fill="x", padx=10, pady=4)
        ttk.Button(box, text="QR 라벨 다운로드+인쇄", command=self.portal_labels_print).pack(fill="x", padx=10, pady=(4, 10))

    def _build_pod_options(self):
        box = self.pod_box
        ttk.Label(box, text="아래 3칸은 비우면 엑셀/현재시간 기준으로 자동 사용합니다.").pack(anchor="w", padx=10, pady=(8, 0))
        self._labeled_entry(box, "Signed By override", self.pod_signed_var)
        self._labeled_entry(box, "Date/Time override", self.pod_datetime_var)
        self._labeled_entry(box, "Remarks override", self.pod_remarks_var)
        ttk.Frame(box).pack(pady=4)

    def _build_log_column(self):
        controls = ttk.Frame(self.log_box)
        controls.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Button(controls, text="로그 새로고침", command=self.refresh_log).pack(side="left")
        ttk.Button(controls, text="접기/펴기", command=self.toggle_log).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="로그 파일 열기", command=self.open_log).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="폴더 열기", command=self.open_folder).pack(side="left", padx=(8, 0))

        self.log_text = tk.Text(self.log_box, wrap="none", height=30, font=("Consolas", 9), bg="#ffffff")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _tick(self):
        if self.started_process and self.started_process.poll() is not None:
            self.status_var.set(f"실행 종료: code {self.started_process.returncode}")
            self.started_process = None
        if self.log_visible:
            self.refresh_log(auto=True)
        self.after(5000, self._tick)

    def _run(self, args, new_console=True, remember=False):
        cmd = [sys.executable] + args
        try:
            self._append_launcher_log("실행 요청: " + " ".join(args))
            proc = subprocess.Popen(
                cmd,
                cwd=str(BASE_DIR),
                creationflags=CREATE_NEW_CONSOLE if new_console else 0,
            )
            if remember:
                self.started_process = proc
            self.status_var.set("실행 요청: " + " ".join(args))
            return proc
        except Exception as exc:
            messagebox.showerror("실행 실패", str(exc))
            self.status_var.set("실행 실패")
            return None

    def _append_launcher_log(self, message):
        try:
            with LOG_FILE.open("a", encoding="utf-8") as fh:
                fh.write(f"{self._now_text()} [INFO] [Launcher] {message}\n")
        except Exception:
            pass

    def _now_text(self):
        from datetime import datetime

        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _remove_stop_flag(self):
        try:
            STOP_FILE.unlink()
        except FileNotFoundError:
            pass

    def _portal_common(self, require_delivery=True, require_material=False, require_serial=False):
        order = self.portal_order_var.get().strip() or self.order_var.get().strip()
        delivery = self.portal_delivery_var.get().strip()
        material = self.portal_material_var.get().strip()
        qty = self.portal_qty_var.get().strip() or "1"
        serial = self.portal_serial_var.get().strip()
        if not order:
            self._append_launcher_log("Portal 실행 중단: 오더번호 없음")
            messagebox.showwarning("입력 필요", "오더번호를 입력하세요.")
            return None
        if require_delivery and not delivery:
            messagebox.showwarning("입력 필요", "Delivery# / OBD를 입력하세요.")
            return None
        if require_material and not material:
            messagebox.showwarning("입력 필요", "Material을 입력하세요.")
            return None
        if require_serial and not serial:
            messagebox.showwarning("입력 필요", "Serial Number를 입력하세요.")
            return None
        return {"order": order, "delivery": delivery, "material": material, "qty": qty, "serial": serial}

    def start_startup(self):
        if self.started_process and self.started_process.poll() is None:
            messagebox.showwarning("이미 실행 중", "런처에서 시작한 자동화가 이미 실행 중입니다.")
            return
        self._remove_stop_flag()
        self._run(["startup.py"], remember=True)

    def start_main_loop(self):
        if self.started_process and self.started_process.poll() is None:
            messagebox.showwarning("이미 실행 중", "런처에서 시작한 자동화가 이미 실행 중입니다.")
            return
        self._remove_stop_flag()
        self._run(["main.py"], remember=True)

    def run_once(self):
        self._run(["main.py", "--once"])

    def start_watchdog(self):
        if self.watchdog_process and self.watchdog_process.poll() is None:
            messagebox.showwarning("이미 실행 중", "워치독이 이미 실행 중입니다.")
            return
        self.watchdog_process = self._run(["watchdog.py"])
        self.status_var.set("워치독 시작됨")

    def stop_watchdog(self):
        if self.watchdog_process and self.watchdog_process.poll() is None:
            self.watchdog_process.terminate()
            self.status_var.set("워치독 중지됨")
        else:
            messagebox.showinfo("워치독", "런처에서 시작한 워치독이 없습니다 (직접 연 콘솔창이라면 거기서 닫아주세요).")

    def stop_automation(self):
        STOP_FILE.write_text("", encoding="utf-8")
        if self.started_process and self.started_process.poll() is None:
            if messagebox.askyesno("중지", "런처에서 시작한 프로세스를 바로 종료할까요?\n아니오를 누르면 루프가 stop flag를 보고 종료합니다."):
                self.started_process.terminate()
        self.status_var.set("중지 요청됨")

    def run_order(self):
        order = self.order_var.get().strip()
        if not order:
            messagebox.showwarning("오더번호 필요", "오더번호를 입력하세요.")
            return
        self._run(["order.py", order])

    def run_vendor_dashboard(self):
        self._run(["vendor_dashboard.py"])

    def open_new_sap_session(self):
        tcode = self.new_session_tcode_var.get().strip()
        args = ["open_session.py"] + ([tcode] if tcode else [])
        self._run(args)
        self.new_session_tcode_var.set("")

    def trigger_all(self):
        RUN_NOW_FILE.write_text("", encoding="utf-8")
        self.status_var.set("전체 즉시 조회 요청됨")

    def trigger_session(self, name):
        SESSION_FLAGS[name].write_text("", encoding="utf-8")
        self.status_var.set(f"{name} 즉시 실행 요청됨")

    def run_print(self):
        sheet = self.print_sheet_var.get().strip()
        if not sheet:
            messagebox.showwarning("시트 필요", "예: 4-2")
            return
        args = ["main.py", "--print", sheet]
        date_key = self.print_date_var.get().strip()
        if date_key:
            args.append(date_key)
        if self.print_afternoon_var.get():
            args.append("오후")
        self._run(args)

    def run_latest_zpl(self):
        self._run(["print_zpl_file.py"])

    def open_portal_login(self):
        self._run(["portal_login.py"], remember=False)

    def open_portal_login_manual(self):
        """Just opens the automation Chrome profile to the login page - no
        autofill, no auto-clicking Next, no BUIT wait-loop (see portal_login.py
        --open-only). For logging in by hand while the automated login flow
        is being sorted out; once this Chrome window sits on the real logged-
        in portal page, "Serial 등록 & QR 인쇄" works exactly the same either
        way - it only needs a logged-in tab in this profile, not a login it
        performed itself."""
        self._run(["portal_login.py", "--open-only"], remember=False)

    def portal_open_delivery(self):
        vals = self._portal_common(require_delivery=False)
        if vals:
            self._run(["portal_open_delivery.py", vals["order"]])

    def fill_portal_from_excel(self):
        order = self.portal_order_var.get().strip() or self.order_var.get().strip()
        if not order:
            self._append_launcher_log("OBD 자동채움 실패: 오더번호 없음")
            messagebox.showwarning("입력 필요", "오더번호를 입력하세요.")
            return
        try:
            data = find_excel_order(order)
            delivery = data.get("obd", "")
            material = ""
            serial = ""
            for row in data.get("rows", []):
                if _line_kind(row.get("order_text")) != "delivery":
                    continue
                material = self._digits(row.get("material"))
                serial = self._serial_text(row.get("serial"))
                if material:
                    break
            if not delivery:
                raise RuntimeError("엑셀 A열에서 OBD 줄을 찾지 못했습니다.")
            self.portal_delivery_var.set(delivery)
            self.portal_material_var.set(material)
            if serial and not self.portal_serial_var.get().strip():
                self.portal_serial_var.set(serial)
            self.portal_qty_var.set(self.portal_qty_var.get().strip() or "1")
            self._append_launcher_log(
                f"OBD 자동채움 완료: order={order} delivery={delivery} material={material} serial={serial}"
            )
            messagebox.showinfo("OBD 자동채움 완료", f"Delivery# / OBD: {delivery}\nMaterial: {material}\nSerial: {serial}")
        except Exception as exc:
            self._append_launcher_log(f"OBD 자동채움 실패: order={order} error={exc}")
            messagebox.showerror("OBD 자동채움 실패", str(exc))

    def _digits(self, value):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        text = str(value or "").strip()
        if text.endswith(".0") and text[:-2].replace(".", "", 1).isdigit():
            text = text[:-2]
        return "".join(ch for ch in text if ch.isdigit())

    def _serial_text(self, value):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        text = str(value or "").strip()
        if text.endswith(".0") and text[:-2].replace(".", "", 1).isdigit():
            text = text[:-2]
        return text

    def portal_register_serial(self):
        vals = self._portal_common(require_delivery=True, require_material=True, require_serial=True)
        if vals:
            self._run([
                "portal_register_serial.py",
                "--delivery", vals["delivery"],
                "--order", vals["order"],
                "--material", vals["material"],
                "--qty", vals["qty"],
                "--serial", vals["serial"],
            ])

    def portal_pack_prepare(self):
        vals = self._portal_common(require_delivery=True, require_material=True)
        if vals:
            self._run(["portal_pack_post.py", "--delivery", vals["delivery"], "--material", vals["material"], "--qty", vals["qty"]])

    def portal_pack_post(self):
        vals = self._portal_common(require_delivery=True, require_material=True)
        if not vals:
            return
        if messagebox.askyesno("Packing Post", "Portal에서 Packing Post를 실제로 누를까요?"):
            self._run(["portal_pack_post.py", "--delivery", vals["delivery"], "--material", vals["material"], "--qty", vals["qty"], "--post"])

    def portal_labels_print(self):
        vals = self._portal_common(require_delivery=True)
        if vals:
            self._run(["portal_download_labels.py", "--delivery", vals["delivery"], "--print"])
            self._clear_portal_manual_inputs()

    def portal_ship_and_print(self):
        skip_qr = self.portal_skip_qr_var.get()
        material_only = self.portal_material_only_var.get()
        self._append_launcher_log("배송 처리 버튼 클릭됨" + (" (QR 인쇄 생략)" if skip_qr else ""))
        vals = self._portal_common(require_delivery=False, require_material=False, require_serial=False)
        if not vals:
            return
        self._append_launcher_log(f"배송 처리 확인 대기: order={vals['order']} skip_qr={skip_qr}")
        third_step = "3. QR 라벨 다운로드 + 인쇄" if not skip_qr else "3. QR 라벨 다운로드/인쇄 생략"
        msg = (
            "아래 작업을 한 번에 진행할까요?\n\n"
            "1. Serial 등록 + Run ShipERP\n"
            "2. Packing Post\n"
            f"{third_step}\n\n"
            "Delivery#/Material/Qty는 SAP에서 자동 조회하고,\n"
            "Serial은 엑셀에서 먼저 찾습니다."
        )
        if messagebox.askyesno("배송 처리", msg):
            args = [
                "portal_ship_and_print.py",
                "--order", vals["order"],
            ]
            if vals["delivery"]:
                args.extend(["--delivery", vals["delivery"]])
            if vals["material"]:
                args.extend(["--material", vals["material"]])
            if vals["qty"]:
                args.extend(["--qty", vals["qty"]])
            if vals["serial"]:
                args.extend(["--serial", vals["serial"]])
            if material_only:
                args.append("--material-only")
            if skip_qr:
                args.append("--skip-qr")
            self._run(args)
            self._clear_portal_manual_inputs()
        else:
            self._append_launcher_log("배송 처리 취소됨")

    def _clear_portal_manual_inputs(self):
        self.portal_delivery_var.set("")
        self.portal_material_var.set("")
        self.portal_qty_var.set("1")
        self.portal_serial_var.set("")
        self.portal_material_only_var.set(False)
        self._append_launcher_log("Portal 수동 입력값 자동 삭제 완료")

    def _prepare_pod_payload(self):
        order = self.portal_order_var.get().strip() or self.order_var.get().strip()
        if not order:
            self._append_launcher_log("POD 실행 중단: 오더번호 없음")
            messagebox.showwarning("입력 필요", "오더번호를 입력하세요.")
            return None

        delivery = ""
        try:
            pickup_mode = "collected" if self.pod_pickup_done_var.get() else None
            data = find_excel_order(order, pickup_mode=pickup_mode)
            delivery = data.get("obd", "")
            if delivery:
                self.portal_delivery_var.set(delivery)
            self._append_launcher_log(
                f"POD Excel 조회 완료: order={order} delivery={delivery} "
                f"sheet={data.get('sheet')} rows={data.get('start_row')}-{data.get('end_row')}"
            )
        except Exception as exc:
            self._append_launcher_log(f"POD Excel 조회 실패: order={order} error={exc}")

        if not delivery:
            delivery = self.portal_delivery_var.get().strip()
        if not delivery:
            messagebox.showwarning(
                "입력 필요",
                "Delivery# / OBD를 찾지 못했습니다. OBD 확인/자동채움을 먼저 누르거나 직접 입력하세요.",
            )
            return None

        try:
            from portal_update_pod import prepare_values, write_pending

            values = prepare_values(
                order,
                self.pod_signed_var.get().strip(),
                self.pod_datetime_var.get().strip(),
                self.pod_remarks_var.get().strip(),
                pickup_done=self.pod_pickup_done_var.get(),
            )
            return write_pending(delivery, values)
        except Exception as exc:
            self._append_launcher_log(f"POD 확인값 생성 실패: {exc}")
            messagebox.showerror("POD 확인 실패", str(exc))
            return None

    def _pod_command_from_payload(self, payload, update=False):
        args = [
            "portal_update_pod.py",
            "--delivery", str(payload["delivery"]),
            "--order", str(payload["order"]),
            "--signed-by", str(payload["signed_by"]),
            "--datetime", str(payload["delivery_datetime"]),
            "--remarks", str(payload["remarks"]),
        ]
        if self.pod_pickup_done_var.get():
            args.append("--pickup-done")
        if update:
            args.append("--update")
        return args

    def _ask_pod_action(self, payload):
        dialog = tk.Toplevel(self)
        dialog.title("POD 저장 확인")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        result = {"action": "cancel"}
        body = (
            f"오더번호: {payload.get('order', '')}\n"
            f"OBD넘버: {payload.get('delivery', '')}\n"
            f"배송시간: {payload.get('delivery_datetime', '')}\n"
            f"수신인: {payload.get('signed_by', '')}\n"
            f"배송 Serial: {', '.join(payload.get('delivery_serials', [])) or '-'}\n"
            f"회수 Serial: {', '.join(payload.get('pickup_serials', [])) or '-'}\n"
            f"회수완료 체크: {'예' if payload.get('pickup_done') else '아니오'}\n"
            f"Remarks: {payload.get('remarks', '')}\n\n"
            "이 내용으로 POD 저장하시겠습니까?"
        )
        ttk.Label(dialog, text=body, wraplength=520, justify="left").pack(
            fill="x", padx=18, pady=(16, 12)
        )
        buttons = ttk.Frame(dialog)
        buttons.pack(fill="x", padx=18, pady=(0, 16))

        def choose(action):
            result["action"] = action
            dialog.destroy()

        ttk.Button(buttons, text="확인(저장)", style="Primary.TButton", command=lambda: choose("update")).pack(
            side="left", expand=True, fill="x", padx=(0, 6)
        )
        ttk.Button(buttons, text="저장만 안함", command=lambda: choose("fill")).pack(
            side="left", expand=True, fill="x", padx=6
        )
        ttk.Button(buttons, text="취소", command=lambda: choose("cancel")).pack(
            side="left", expand=True, fill="x", padx=(6, 0)
        )

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose("cancel"))
        self.wait_window(dialog)
        return result["action"]

    def _pod_args(self, update=False):
        order = self.portal_order_var.get().strip() or self.order_var.get().strip()
        if not order:
            self._append_launcher_log("POD 실행 중단: 오더번호 없음")
            messagebox.showwarning("입력 필요", "오더번호를 입력하세요.")
            return None
        delivery = ""
        try:
            pickup_mode = "collected" if self.pod_pickup_done_var.get() else None
            data = find_excel_order(order, pickup_mode=pickup_mode)
            delivery = data.get("obd", "")
            if delivery:
                self.portal_delivery_var.set(delivery)
            self._append_launcher_log(
                f"POD Excel 조회 완료: order={order} delivery={delivery} "
                f"sheet={data.get('sheet')} rows={data.get('start_row')}-{data.get('end_row')}"
            )
        except Exception as exc:
            self._append_launcher_log(f"POD Excel 조회 실패: order={order} error={exc}")

        if not delivery:
            delivery = self.portal_delivery_var.get().strip()
        if not delivery:
            messagebox.showwarning(
                "입력 필요",
                "Delivery# / OBD를 찾지 못했습니다. OBD 확인/자동채움을 먼저 누르거나 직접 입력하세요.",
            )
            return None

        args = ["portal_update_pod.py", "--delivery", delivery, "--order", order]
        signed_by = self.pod_signed_var.get().strip()
        if signed_by:
            args.extend(["--signed-by", signed_by])
        date_time = self.pod_datetime_var.get().strip()
        if date_time:
            args.extend(["--datetime", date_time])
        remarks = self.pod_remarks_var.get().strip()
        if remarks:
            args.extend(["--remarks", remarks])
        if self.pod_pickup_done_var.get():
            args.append("--pickup-done")
        if update:
            args.append("--update")
        return args

    def run_pod_fill(self):
        payload = self._prepare_pod_payload()
        if not payload:
            return

        action = self._ask_pod_action(payload)
        if action == "update":
            self._append_launcher_log("POD 확인 후 저장 진행")
            self._run(self._pod_command_from_payload(payload, update=True))
        elif action == "fill":
            self._append_launcher_log("POD 확인 후 저장 없이 입력만 진행")
            self._run(self._pod_command_from_payload(payload, update=False))
        else:
            self._append_launcher_log("POD 확인 취소됨")

    def run_pod_update(self):
        args = self._pod_args(update=True)
        if not args:
            return
        confirm_msg = "Portal에서 Update POD를 실제로 누를까요?"
        try:
            pending = json.loads(POD_PENDING_FILE.read_text(encoding="utf-8"))
            current_delivery = args[args.index("--delivery") + 1]
            current_order = args[args.index("--order") + 1]
            if pending.get("delivery") == current_delivery and pending.get("order") == current_order:
                args.extend(["--signed-by", pending.get("signed_by", "")])
                args.extend(["--datetime", pending.get("delivery_datetime", "")])
                args.extend(["--remarks", pending.get("remarks", "")])
                confirm_msg = (
                    "아래 내용으로 Update POD를 누를까요?\n\n"
                    f"Delivery#: {pending.get('delivery')}\n"
                    f"Order: {pending.get('order')}\n"
                    f"Status: {pending.get('tracking_status')}\n"
                    f"Time: {pending.get('delivery_datetime')}\n"
                    f"Signed By: {pending.get('signed_by')}\n"
                    f"Delivery Serial: {', '.join(pending.get('delivery_serials', [])) or '-'}\n"
                    f"Pickup Serial: {', '.join(pending.get('pickup_serials', [])) or '-'}\n"
                    f"Pickup Done: {'Yes' if pending.get('pickup_done') else 'No'}\n"
                    f"Remarks: {pending.get('remarks')}"
                )
        except Exception:
            pass
        if messagebox.askyesno("POD 업데이트 확인", confirm_msg):
            self._run(args)

    def toggle_portal_advanced(self):
        self.portal_advanced_visible = not self.portal_advanced_visible
        if self.portal_advanced_visible:
            self.portal_advanced_box.pack(fill="x", pady=(0, 10))
            self.portal_advanced_toggle.configure(text="고급/문제해결 버튼 접기")
        else:
            self.portal_advanced_box.pack_forget()
            self.portal_advanced_toggle.configure(text="고급/문제해결 버튼 펼치기")

    def toggle_sap_tools(self):
        self.sap_tools_visible = not self.sap_tools_visible
        if self.sap_tools_visible:
            self.sap_tools_box.pack(fill="x", pady=(0, 8))
            self.sap_tools_toggle.configure(text="▲")
        else:
            self.sap_tools_box.pack_forget()
            self.sap_tools_toggle.configure(text="▼")

    def toggle_sap_more(self):
        self.sap_more_visible = not self.sap_more_visible
        if self.sap_more_visible:
            self.sap_more_box.pack(fill="x", pady=(0, 10))
            self.sap_more_toggle.configure(text="특정 창 조회 / 엑셀 인쇄 ▲")
        else:
            self.sap_more_box.pack_forget()
            self.sap_more_toggle.configure(text="특정 창 조회 / 엑셀 인쇄 ▼")

    def toggle_portal_detail(self):
        self.portal_detail_visible = not self.portal_detail_visible
        if self.portal_detail_visible:
            self.portal_detail_box.pack(fill="x", padx=10, pady=(0, 10))
            self.portal_detail_toggle.configure(text="▲")
        else:
            self.portal_detail_box.pack_forget()
            self.portal_detail_toggle.configure(text="▼")

    def toggle_pod_options(self):
        self.pod_options_visible = not self.pod_options_visible
        if self.pod_options_visible:
            self.pod_box.pack(fill="x", pady=(0, 10))
            self.pod_options_toggle.configure(text="POD 옵션 접기")
        else:
            self.pod_box.pack_forget()
            self.pod_options_toggle.configure(text="POD 옵션 펼치기")

    def toggle_log(self):
        self.log_visible = not self.log_visible
        if self.log_visible:
            self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
            self.refresh_log()
        else:
            self.log_text.pack_forget()

    def refresh_log(self, auto=False):
        try:
            header = ""
            if POD_PENDING_FILE.exists():
                try:
                    pending = json.loads(POD_PENDING_FILE.read_text(encoding="utf-8"))
                    header = (
                        "[POD 업데이트 전 확인]\n"
                        f"Delivery#: {pending.get('delivery', '')}\n"
                        f"Order: {pending.get('order', '')}\n"
                        f"Tracking Status: {pending.get('tracking_status', '')}\n"
                        f"Delivery Date & Time: {pending.get('delivery_datetime', '')}\n"
                        f"Signed By: {pending.get('signed_by', '')}\n"
                        f"Remarks: {pending.get('remarks', '')}\n"
                        f"Excel: {pending.get('excel_sheet', '')} rows {pending.get('excel_rows', '')}\n"
                        "문제 없으면 'POD 최종 저장'을 누르세요.\n"
                        + "=" * 70
                        + "\n\n"
                    )
                except Exception:
                    header = "[POD 확인값을 읽지 못했습니다.]\n\n"
            if not LOG_FILE.exists():
                content = "automation.log 없음"
            else:
                data = LOG_FILE.read_text(encoding="utf-8", errors="replace").splitlines()
                content = "\n".join(data[-160:])
            content = header + content
            current = self.log_text.get("1.0", "end-1c")
            if current != content:
                self.log_text.delete("1.0", "end")
                self.log_text.insert("1.0", content)
                self.log_text.see("end")
        except Exception as exc:
            if not auto:
                messagebox.showerror("로그 읽기 실패", str(exc))

    def open_log(self):
        if LOG_FILE.exists():
            os.startfile(LOG_FILE)
        else:
            messagebox.showinfo("로그 없음", "automation.log 파일이 아직 없습니다.")

    def open_folder(self):
        os.startfile(BASE_DIR)


if __name__ == "__main__":
    app = Launcher()
    app.mainloop()
