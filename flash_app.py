"""Curve Flash: PySide6 GUI for Excel point-coordinate flash calculations."""
from __future__ import annotations
import sys
import traceback
from pathlib import Path
from dataclasses import dataclass
from math import hypot

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFormLayout, QFrame, QLabel, QLineEdit, QPushButton, QComboBox, QFileDialog,
    QMessageBox, QProgressBar)
from autocad_export import draw_flash_results


@dataclass(frozen=True)
class Point:
    number: int
    northing: float
    easting: float
    row: int


def sheets(path):
    if Path(path).suffix.lower() == ".xls":
        return xlrd.open_workbook(path, on_demand=True).sheet_names()
    if Path(path).suffix.lower() == ".xlsx":
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            return list(wb.sheetnames)
        finally:
            wb.close()
    raise ValueError("فقط فایل‌های XLS و XLSX پشتیبانی می‌شوند.")


def numeric(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        value = value.strip().replace(",", "")
        if not value:
            return None
    try:
        result = float(value)
        return result if abs(result) != float("inf") else None
    except (TypeError, ValueError):
        return None


def read_points(path, sheet_name):
    suffix = Path(path).suffix.lower()
    if suffix == ".xls":
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_name(sheet_name) if sheet_name else book.sheet_by_index(0)
        rows = [tuple(sh.row_values(i)) for i in range(sh.nrows)]
        title = sh.name
    elif suffix == ".xlsx":
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            sh = book[sheet_name] if sheet_name else book[book.sheetnames[0]]
            title = sh.title
            rows = [tuple(row) for row in sh.iter_rows(values_only=True)]
        finally:
            book.close()
    else:
        raise ValueError("فرمت فایل باید XLS یا XLSX باشد.")
    points = []
    for row_no, row in enumerate(rows, 1):
        if len(row) < 3:
            continue
        n, y, x = (numeric(row[i]) for i in range(3))
        if n is None or y is None or x is None or abs(n-round(n)) > 1e-8:
            continue
        points.append(Point(int(round(n)), y, x, row_no))
    if len(points) < 19:
        raise ValueError(f"حداقل ۱۹ نقطه معتبر لازم است؛ تعداد پیدا شده: {len(points)}")
    if len({p.number for p in points}) != len(points):
        raise ValueError("شماره نقاط تکراری است.")
    return title, rows, points


def offset(a, b, p):
    dx, dy = b.easting-a.easting, b.northing-a.northing
    length = hypot(dx, dy)
    if length < 1e-12:
        raise ValueError(f"طول وتر بین نقاط {a.number} و {b.number} صفر است.")
    return (dx*(p.northing-a.northing)-dy*(p.easting-a.easting))/length


def calculate(points, interval=18):
    """Calculate one midpoint flash for each non-overlapping 18-interval chord.

    Example: points 1..19 produce one chord (1, 19), measured only at point 10.
    The next chord starts at point 19, then 37, and so on.
    """
    if len(points) <= interval:
        raise ValueError("تعداد نقاط برای فاصله وتر ۱۸ کافی نیست.")
    results, details = [], []
    for i in range(0, len(points)-interval, interval):
        a, b = points[i], points[i+interval]
        midpoint = points[i + interval // 2]
        signed = offset(a, b, midpoint)
        item = (a.number, b.number, midpoint.number, signed, abs(signed))
        details.append(item)
        results.append((a.number, b.number, midpoint.number, signed, abs(signed), 1))
    return results, details


def export_report(path, title, rows, points, results, details):
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    critical_ids = {r[2] for r in results}
    critical_rows = {p.row for p in points if p.number in critical_ids}
    width = max((len(r) for r in rows), default=3)
    highlight = PatternFill("solid", fgColor="FFF2CC")
    for idx, row in enumerate(rows, 1):
        ws.append(list(row)+[None]*(width-len(row)))
        if idx in critical_rows:
            for c in ws[idx]:
                c.fill = highlight
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    summary = wb.create_sheet("Flash Results")
    summary.append(["Chord start", "Chord end", "Critical point", "Signed flash", "Absolute flash", "Internal points"])
    for r in results:
        summary.append(list(r))
    detail = wb.create_sheet("Midpoint Flash")
    detail.append(["Chord start", "Chord end", "Point", "Signed perpendicular offset", "Absolute offset"])
    for d in details:
        detail.append(list(d))
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for sh in (summary, detail):
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = sh.dimensions
        for c in sh[1]:
            c.fill, c.font = header_fill, header_font
            c.alignment = Alignment(horizontal="center")
        for col in sh.columns:
            letter = get_column_letter(col[0].column)
            sh.column_dimensions[letter].width = min(34, max(12, max(len(str(c.value or "")) for c in col)+2))
        for row in sh.iter_rows(min_row=2):
            for c in row:
                if c.column in (4, 5):
                    c.number_format = "0.000"
    wb.save(target)
    return target


def analyze(source, sheet, destination):
    title, rows, points = read_points(source, sheet)
    results, details = calculate(points, 18)
    export_report(destination, title, rows, points, results, details)
    maximum = max(results, key=lambda r: r[4])
    return {"points":len(points), "chords":len(results), "offsets":len(details),
            "max":maximum[4], "critical":maximum[2], "chord":(maximum[0],maximum[1]),
            "output":str(destination)}


class Worker(QThread):
    success = Signal(dict)
    failure = Signal(str)
    def __init__(self, source, sheet, destination):
        super().__init__()
        self.args = source, sheet, destination
    def run(self):
        try:
            self.success.emit(analyze(*self.args))
        except Exception as exc:
            self.failure.emit(str(exc))


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.last_calculation_ready = False
        self.setWindowTitle("Curve Flash | محاسبه فلش قوس")
        self.resize(780, 480)
        root = QWidget()
        root.setLayoutDirection(Qt.RightToLeft)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(28,24,28,24)
        title = QLabel("محاسبه فلش قوس")
        title.setObjectName("title")
        layout.addWidget(title)
        layout.addWidget(QLabel("انتخاب فایل مختصات Excel و تولید گزارش محاسباتی"))
        panel = QFrame()
        panel.setObjectName("panel")
        form = QFormLayout(panel)
        form.setContentsMargins(18,18,18,18)
        self.source = QLineEdit()
        b1 = QPushButton("انتخاب فایل...")
        b1.clicked.connect(self.choose_source)
        r1 = QHBoxLayout(); r1.addWidget(self.source,1); r1.addWidget(b1)
        form.addRow("فایل ورودی:",r1)
        self.sheet = QComboBox()
        form.addRow("برگه:",self.sheet)
        self.destination = QLineEdit()
        b2 = QPushButton("مسیر خروجی...")
        b2.clicked.connect(self.choose_destination)
        r2 = QHBoxLayout(); r2.addWidget(self.destination,1); r2.addWidget(b2)
        form.addRow("فایل خروجی:",r2)
        layout.addWidget(panel)
        info = QLabel("قالب داده: ستون A شماره نقطه، ستون B مختصات شمالی (Y)، ستون C مختصات شرقی (X). "
                      "وتر از نقطه i تا i+18 تشکیل می‌شود و ۱۷ نقطه داخلی بررسی می‌شوند. "
                      "واحد فلش برابر واحد مختصات ورودی است.")
        info.setWordWrap(True); info.setObjectName("info"); layout.addWidget(info)
        self.run_button = QPushButton("شروع محاسبه و ساخت فایل خروجی")
        self.run_button.setObjectName("run")
        self.run_button.clicked.connect(self.start)
        layout.addWidget(self.run_button)
        self.autocad_button = QPushButton("ارسال نقاط فلش و اندازه‌گذاری به AutoCAD")
        self.autocad_button.setEnabled(False)
        self.autocad_button.clicked.connect(self.send_to_autocad)
        layout.addWidget(self.autocad_button)
        self.progress = QProgressBar(); self.progress.setRange(0,0); self.progress.hide()
        layout.addWidget(self.progress)
        self.status = QLabel("آماده"); self.status.setObjectName("status"); self.status.setWordWrap(True)
        layout.addWidget(self.status); layout.addStretch(1)
        self.setCentralWidget(root)
        self.setStyleSheet("""
          QMainWindow{background:#F2F5F9} QLabel{color:#25384B;font-size:10pt}
          QLabel#title{font-size:23pt;font-weight:700;color:#17365D}
          QFrame#panel,QLabel#status{background:white;border:1px solid #D8E1EA;border-radius:8px;padding:8px}
          QLabel#info{background:#E7EFF7;border-radius:7px;padding:12px}
          QLineEdit,QComboBox{background:white;border:1px solid #C8D3DF;border-radius:5px;padding:8px}
          QPushButton{background:#E4EBF3;border:1px solid #C8D3DF;border-radius:5px;padding:9px 14px}
          QPushButton#run{background:#1F5A91;color:white;border:0;font-weight:bold;padding:13px}
          QPushButton:disabled{background:#AAB8C6;color:white}
        """)
    def choose_source(self):
        path,_=QFileDialog.getOpenFileName(self,"انتخاب فایل Excel",str(Path.cwd()),"Excel (*.xls *.xlsx)")
        if not path: return
        self.source.setText(path)
        try:
            self.sheet.clear(); self.sheet.addItems(sheets(path))
            self.destination.setText(str(Path(path).with_name(Path(path).stem+"_flash_result.xlsx")))
            self.status.setText("فایل انتخاب شد؛ برگه موردنظر را انتخاب کنید.")
        except Exception as exc:
            QMessageBox.critical(self,"خطای فایل",str(exc))
    def choose_destination(self):
        path,_=QFileDialog.getSaveFileName(self,"ذخیره گزارش",self.destination.text() or str(Path.cwd()/"flash_result.xlsx"),"Excel (*.xlsx)")
        if path: self.destination.setText(path if path.lower().endswith(".xlsx") else path+".xlsx")
    def start(self):
        src,dst=self.source.text().strip(),self.destination.text().strip()
        if not src or not Path(src).is_file():
            QMessageBox.warning(self,"ورودی","یک فایل XLS یا XLSX معتبر انتخاب کنید."); return
        if not dst:
            QMessageBox.warning(self,"خروجی","مسیر فایل خروجی را مشخص کنید."); return
        if Path(src).resolve()==Path(dst).resolve():
            QMessageBox.warning(self,"مسیر","فایل خروجی نباید روی فایل ورودی نوشته شود."); return
        self.last_calculation_ready = False
        self.autocad_button.setEnabled(False)
        self.run_button.setEnabled(False); self.progress.show(); self.status.setText("در حال محاسبه...")
        self.worker=Worker(src,self.sheet.currentText(),dst)
        self.worker.success.connect(self.finished); self.worker.failure.connect(self.failed); self.worker.start()
    def finished(self,r):
        self.progress.hide(); self.run_button.setEnabled(True)
        self.last_calculation_ready = True
        self.autocad_button.setEnabled(True)
        self.status.setText(f"محاسبه انجام شد | نقاط: {r['points']} | بازه‌ها: {r['chords']} | "
          f"فلش‌های محاسبه‌شده: {r['offsets']}\nبیشترین فلش: {r['max']:.3f} در نقطه {r['critical']} "
          f"(وتر {r['chord'][0]} تا {r['chord'][1]})\nخروجی: {r['output']}")
        QMessageBox.information(self,"پایان","فایل گزارش Excel ایجاد شد.")
    def failed(self,message):
        self.progress.hide(); self.run_button.setEnabled(True)
        self.last_calculation_ready = False
        self.autocad_button.setEnabled(False)
        self.status.setText("پردازش ناموفق بود.")
        QMessageBox.critical(self,"خطای محاسبه",message)

    def send_to_autocad(self):
        if not self.last_calculation_ready:
            QMessageBox.warning(self, "ارسال به AutoCAD", "ابتدا محاسبه را با موفقیت انجام دهید.")
            return
        source = self.source.text().strip()
        if not source or not Path(source).is_file():
            QMessageBox.warning(self, "فایل ورودی", "فایل Excel ورودی پیدا نشد.")
            return
        answer = QMessageBox.question(
            self,
            "تأیید ترسیم در AutoCAD",
            "قرار است فقط نقطه میانی هر بازه و اندازه فلش آن در نقشه فعال AutoCAD رسم شوند.\\n"
            "پیشنهاد می‌شود پیش از ادامه از فایل DWG یک نسخه پشتیبان بگیرید.\\n\\nادامه می‌دهید؟",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        try:
            _title, _rows, points = read_points(source, self.sheet.currentText())
            results, _details = calculate(points, 18)
            self.autocad_button.setEnabled(False)
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                report = draw_flash_results(points, results)
            finally:
                QApplication.restoreOverrideCursor()
            self.status.setText(
                f"ترسیم در AutoCAD انجام شد | نقشه: {report['drawing']} | "
                f"اندازه‌گذاری‌ها: {report['dimensions']} | لایه: {report['layer']}"
            )
            if report["skipped"]:
                QMessageBox.warning(
                    self, "ترسیم با هشدار",
                    f"{report['dimensions']} اندازه‌گذاری رسم شد و {report['skipped']} مورد به علت داده نامعتبر رد شد."
                )
            else:
                QMessageBox.information(
                    self, "پایان ترسیم",
                    f"{report['dimensions']} نقطه بحرانی و اندازه‌گذاری در نقشه "
                    f"{report['drawing']} روی لایه {report['layer']} رسم شد."
                )
        except Exception as exc:
            print("\\n" + "=" * 72, file=sys.stderr)
            print("[Curve Flash] AutoCAD export failed", file=sys.stderr)
            print(f"Exception: {type(exc).__name__}: {exc}", file=sys.stderr)
            traceback.print_exc(file=sys.stderr)
            print("=" * 72 + "\\n", file=sys.stderr)
            sys.stderr.flush()
            QMessageBox.critical(
                self,
                "خطای اتصال یا ترسیم AutoCAD",
                f"{type(exc).__name__}: {exc}\\n\\nجزئیات کامل خطا در پنجره ترمینال نمایش داده شد."
            )
        finally:
            self.autocad_button.setEnabled(self.last_calculation_ready)


def main():
    app=QApplication(sys.argv)
    window=Window(); window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
