import json
import re
import ctypes
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import sys

if sys.platform.startswith("win"):
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("company.contract.templater")

from PySide6.QtCore import Qt, QDate, QRegularExpression
from PySide6.QtGui import QRegularExpressionValidator
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QFormLayout, QLineEdit, QTextEdit, QPushButton, QMessageBox, QFileDialog,
    QDateEdit, QScrollArea, QCheckBox, QSpinBox, QDoubleSpinBox
)

from docxtpl import DocxTemplate
from pathlib import Path

def get_base_dir() -> Path:
    # Если собрали PyInstaller'ом: base = папка рядом с app.exe
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # Если запускаем как скрипт: base = папка рядом с app.py
    return Path(__file__).resolve().parent

import os, sys
from pathlib import Path

def resource_dir() -> Path:
    # где лежат встроенные ресурсы (schemas/templates/assets)
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

def desktop_dir() -> Path:
    # Windows / Mac / Linux: обычно работает
    return Path.home() / "Desktop"

    p = docs / app_name
    p.mkdir(parents=True, exist_ok=True)
    return p

RES_DIR = resource_dir()
DATA_DIR = desktop_dir()

SCHEMAS_DIR   = RES_DIR / "schemas"
TEMPLATES_DIR = RES_DIR / "templates"
ASSETS_DIR    = RES_DIR / "assets"
OUTPUT_DIR    = DATA_DIR / "Договоры"
OUTPUT_DIR.mkdir(exist_ok=True)

RU_MONTHS_GEN = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля",
    5: "мая", 6: "июня", 7: "июля", 8: "августа",
    9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
}

ACCENT = "#41299e"

APP_QSS = f"""
QWidget {{
    font-size: 10.5pt;
}}

QLabel {{
    color: #222;
}}

QLineEdit, QTextEdit, QDateEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
    border: 1px solid #cfcfe6;
    border-radius: 8px;
    padding: 6px 8px;
    background: #ffffff;
}}

QLineEdit:focus, QTextEdit:focus, QDateEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {{
    border: 2px solid {ACCENT};
}}

QPushButton {{
    background: #41299e;
    color: white;
    border: none;
    border-radius: 10px;
    padding: 8px 12px;
    font-weight: 600;
}}

QPushButton:hover {{
    background: #37228a;
}}

QPushButton:pressed {{
    background: #2d1c73;
}}

QPushButton:disabled {{
    background: #c9c9d9;
    color: #6f6f7a;
}}

QCheckBox::indicator:checked {{
    background: {ACCENT};
    border: 1px solid {ACCENT};
}}

QScrollArea {{
    border: none;
}}
"""

def fmt_ddmmyyyy(d: date) -> str:
    return f"{d.day:02d}.{d.month:02d}.{d.year:04d}"


def fmt_long_ru(d: date) -> str:
    return f"«{d.day:02d}» {RU_MONTHS_GEN[d.month]} {d.year} г."


# --- money to words (RUB) ---
def _morph(n: int, f1: str, f2: str, f5: str) -> str:
    n = abs(n) % 100
    n1 = n % 10
    if 10 < n < 20:
        return f5
    if 1 < n1 < 5:
        return f2
    if n1 == 1:
        return f1
    return f5


_UNITS_M = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_UNITS_F = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_TEENS = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать",
          "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"]
_TENS = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят",
         "шестьдесят", "семьдесят", "восемьдесят", "девяносто"]
_HUNDREDS = ["", "сто", "двести", "триста", "четыреста", "пятьсот",
             "шестьсот", "семьсот", "восемьсот", "девятьсот"]


def _triad_to_words(n: int, female: bool) -> str:
    words = []
    h = n // 100
    t = (n // 10) % 10
    u = n % 10

    if h:
        words.append(_HUNDREDS[h])

    if t == 1:
        words.append(_TEENS[u])
    else:
        if t:
            words.append(_TENS[t])
        if u:
            words.append((_UNITS_F if female else _UNITS_M)[u])

    return " ".join(words).strip()


def int_to_words_ru(n: int) -> str:
    if n == 0:
        return "ноль"

    parts = []
    n_abs = abs(n)

    triads = [
        (0, "", "", "", False),
        (1, "тысяча", "тысячи", "тысяч", True),
        (2, "миллион", "миллиона", "миллионов", False),
        (3, "миллиард", "миллиарда", "миллиардов", False),
    ]

    i = 0
    while n_abs > 0 and i < len(triads):
        tri = n_abs % 1000
        if tri:
            _, f1, f2, f5, female = triads[i]
            w = _triad_to_words(tri, female)
            if i > 0:
                w = f"{w} {_morph(tri, f1, f2, f5)}"
            parts.append(w.strip())
        n_abs //= 1000
        i += 1

    res = " ".join(reversed(parts)).strip()
    if n < 0:
        res = "минус " + res
    return res


def money_rub_to_words(amount: float) -> str:
    rub = int(amount)
    kop = int(round((amount - rub) * 100))
    if kop == 100:
        rub += 1
        kop = 0

    rub_words = int_to_words_ru(rub)
    rub_unit = _morph(rub, "рубль", "рубля", "рублей")
    kop_unit = _morph(kop, "копейка", "копейки", "копеек")

    return f"{rub_words} {rub_unit} {kop:02d} {kop_unit}"


def fmt_money_num_ru_no_kop(v: float) -> str:
    rub = int(round(v))
    return f"{rub:,}".replace(",", " ")


def fio_to_initials(fio_str: str) -> str:
    parts = [p for p in fio_str.replace("  ", " ").split(" ") if p]
    if not parts:
        return ""
    fam = parts[0]
    i = (parts[1][0] + ".") if len(parts) > 1 and parts[1] else ""
    o = (parts[2][0] + ".") if len(parts) > 2 and parts[2] else ""
    if i and o:
        return f"{fam} {i} {o}"
    if i:
        return f"{fam} {i}"
    return fam


@dataclass
class FieldDef:
    key: str
    label: str
    type: str = "text"      # text | multiline | date | int | money
    required: bool = False
    default: object = None
    editable_by_checkbox: bool = False
    checkbox_label: str = "Ввести дату вручную"
    checkbox_default: bool = False


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Шаблонизатор договоров")
        self.resize(920, 700)

        OUTPUT_DIR.mkdir(exist_ok=True)
        TEMPLATES_DIR.mkdir(exist_ok=True)
        SCHEMAS_DIR.mkdir(exist_ok=True)

        self.schema_combo = QComboBox()
        self.schema_combo.addItem("Выберите нужный шаблон…", userData=None)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("Шаблон:"))
        top_row.addWidget(self.schema_combo, 1)

        self.form_widget = QWidget()
        self.form_layout = QFormLayout(self.form_widget)
        self.form_layout.setLabelAlignment(Qt.AlignRight)
        self.form_layout.setFormAlignment(Qt.AlignTop)

        self.schema_combo.currentIndexChanged.connect(self.on_schema_changed)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(self.form_widget)

        self.btn_generate = QPushButton("Сгенерировать договор")
        self.btn_generate.clicked.connect(self.generate)
        self.btn_generate.setEnabled(False)

        root = QVBoxLayout(self)
        root.addLayout(top_row)
        root.addWidget(self.scroll, 1)
        root.addWidget(self.btn_generate)

        self.schemas = self.load_schemas()
        self.widgets_by_key = {}  # key -> (FieldDef, widget)
        self.current_schema_name = None

        if not self.schemas:
            QMessageBox.warning(
                self, "Нет схем",
                f"Не найдены схемы в папке:\n{SCHEMAS_DIR}\n\nСоздай schemas/*.json"
            )
        else:
            for schema_name, schema in self.schemas.items():
                self.schema_combo.addItem(schema["title"], userData=schema_name)
            self.schema_combo.setCurrentIndex(0)

    def load_schemas(self):
        schemas = {}
        for p in SCHEMAS_DIR.glob("*.json"):
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
                title = data.get("title", p.stem)
                fields = [FieldDef(**f) for f in data.get("fields", [])]
                template_docx = TEMPLATES_DIR / f"{p.stem}.docx"
                schemas[p.stem] = {"title": title, "fields": fields, "template": template_docx}
            except Exception as e:
                print(f"Ошибка чтения схемы {p}: {e}")
        return schemas

    def clear_form(self):
        while self.form_layout.rowCount():
            self.form_layout.removeRow(0)
        self.widgets_by_key.clear()

    def on_schema_changed(self):
        is_ok = bool(self.schema_combo.currentData())
        if hasattr(self, "btn_generate"):
            self.btn_generate.setEnabled(is_ok)
        if not is_ok:
            self.clear_form()
            self.current_schema_name = None
            return
        schema_name = self.schema_combo.currentData()
        if not schema_name:
            return
        self.current_schema_name = schema_name
        schema = self.schemas[schema_name]

        self.clear_form()

        for f in schema["fields"]:
            w = self.create_widget_for_field(f)
            self.widgets_by_key[f.key] = (f, w)

            label = QLabel(f.label + (" *" if f.required else ""))
            label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            self.form_layout.addRow(label, w)

    # --- masks & validators for specific keys ---
    def apply_masks_and_validators(self, key: str, w: QLineEdit):
    
    #Важно: если используется inputMask — НЕ ставим regex-валидатор,
    #иначе Qt может блокировать ввод, пока строка не станет полностью валидной.
    #Маска + финальная проверка hasAcceptableInput() — достаточно.
    
        def set_regex(pattern: str):
            rx = QRegularExpression(pattern)
            w.setValidator(QRegularExpressionValidator(rx, w))

    # phone/email в одном поле: валидатор "или телефон, или email"
        if key == "phone_email":
            phone_rx = r'^\+[1-9] \(\d{3}\) \d{3}-\d{2}-\d{2}$'
            email_rx = r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+\.[A-Za-z]{2,}$'
            set_regex(rf'(?:{phone_rx})|(?:{email_rx})')
            w.setPlaceholderText('+7 (999) 123-45-67  или  name@mail.ru')
            return

    # Если позже разнесёшь:
        if key == "phone":
            w.setInputMask(r'+0 (000) 000-00-00;_')
            w.setPlaceholderText('+7 (999) 123-45-67')
            return

        if key == "email":
            set_regex(r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+\.[A-Za-z]{2,}$')
            w.setPlaceholderText('name@mail.ru')
            return

    # Паспорт серия: 4 цифры
        if key == "passport_series":
            w.setInputMask("0000;_")
            return

        if key == "ogrnip":
            w.setInputMask("000000000000000;_")  # 15
            return

    # Паспорт номер: 6 цифр
        if key == "passport_number":
            w.setInputMask("000000;_")
            return

    # Код подразделения: NNN-NNN
        if key == "passport_code":
            w.setInputMask("000-000;_")
            return

    # ИНН: 12 цифр
        if key == "inn":
            w.setInputMask("000000000000;_")
            return

    # Счет: 20 цифр
        if key == "bank_account":
            w.setInputMask("00000000000000000000;_")
            return

    # БИК: 9 цифр
        if key == "bank_bik":
            w.setInputMask("000000000;_")
            return

    # Корр. счет: 20 цифр
        if key == "bank_corr":
            w.setInputMask("00000000000000000000;_")
            return

    # Название курса: без кавычек (тут маски не помогут, оставляем regex)
        if key == "course_name":
            set_regex(r'^[^"«»“”]+$')
            w.setPlaceholderText('Без кавычек')
            return

    def create_widget_for_field(self, f: FieldDef):
        if f.type == "multiline":
            w = QTextEdit()
            w.setFixedHeight(90)
            if isinstance(f.default, str):
                w.setPlainText(f.default)
            return w

        if f.type == "date":
            date_edit = QDateEdit()
            date_edit.setCalendarPopup(True)
            date_edit.setDate(QDate.currentDate())

            if isinstance(f.default, str) and f.default.lower() == "today":
                date_edit.setDate(QDate.currentDate())

            if f.editable_by_checkbox:
                cb = QCheckBox(f.checkbox_label or "Ввести дату вручную")
                cb.setChecked(False)
                date_edit.setEnabled(False)

                def on_toggle(state: bool):
                    date_edit.setEnabled(state)

                cb.toggled.connect(on_toggle)

                box = QWidget()
                lay = QHBoxLayout(box)
                lay.setContentsMargins(0, 0, 0, 0)
                lay.addWidget(date_edit)
                lay.addWidget(cb)
                lay.addStretch(1)

                box._date_edit = date_edit
                box._checkbox = cb
                return box

            return date_edit

        if f.type == "int":
            spin = QSpinBox()
            spin.setRange(0, 10_000_000)
            # default
            if isinstance(f.default, int):
                spin.setValue(f.default)
            elif isinstance(f.default, str) and f.default.lower() == "auto":
                spin.setValue(0)
            else:
                spin.setValue(1)
                # checkbox-lock
            if f.editable_by_checkbox:
                cb = QCheckBox(f.checkbox_label or "Изменить значение")
                cb.setChecked(False)
                spin.setEnabled(False)
                def on_toggle(state: bool):
                    spin.setEnabled(state)
                cb.toggled.connect(on_toggle)
                box = QWidget()
                lay = QHBoxLayout(box)
                lay.setContentsMargins(0, 0, 0, 0)
                lay.addWidget(spin)
                lay.addWidget(cb)
                lay.addStretch(1)
                box._spin = spin
                box._checkbox = cb
                return box
            return spin

        if f.type == "percent":
            spin = QDoubleSpinBox()
            spin.setRange(0, 100)
            spin.setDecimals(2)
            spin.setSingleStep(1.0)
            if isinstance(f.default, (int, float)):
                spin.setValue(float(f.default))
            else:
                spin.setValue(0.0)

            if f.editable_by_checkbox:
                cb = QCheckBox(f.checkbox_label or "Включить")
                checked = bool(getattr(f, "checkbox_default", False))
                cb.setChecked(checked)
                spin.setEnabled(checked)

                def on_toggle(state: bool):
                    spin.setEnabled(state)
                    if not state:
                        spin.setValue(22.0)

                cb.toggled.connect(on_toggle)

                box = QWidget()
                lay = QHBoxLayout(box)
                lay.setContentsMargins(0, 0, 0, 0)
                lay.addWidget(spin)
                lay.addWidget(cb)
                lay.addStretch(1)

                box._dspin = spin
                box._checkbox = cb
                return box
            return spin

        if f.type == "money":
            w = QDoubleSpinBox()
            w.setRange(0, 1_000_000_000)
            w.setDecimals(2)
            w.setSingleStep(100.0)
            if isinstance(f.default, (int, float)):
                w.setValue(float(f.default))
            return w

        # default: text
        w = QLineEdit()
        if isinstance(f.default, str):
            w.setText(f.default)

        # apply key-based masks/validators
        self.apply_masks_and_validators(f.key, w)
        return w

    def read_value(self, f: FieldDef, w):
        if f.type == "multiline":
            return w.toPlainText().strip()

        if f.type == "date":
            qd = w._date_edit.date() if hasattr(w, "_date_edit") else w.date()
            return date(qd.year(), qd.month(), qd.day())

        if f.type == "int":
            if hasattr(w, "_spin"):
                return int(w._spin.value())
            return int(w.value())

        if f.type == "money":
            return float(w.value())

        if f.type == "percent":
            if hasattr(w, "_dspin"):
                return float(w._dspin.value())
            return float(w.value())

        return w.text().strip()

    # --- counters for auto numbering ---
    def counters_path(self) -> Path:
        return OUTPUT_DIR / "_counters.json"

    def next_contract_seq_for_date(self, d: date) -> int:
        p = self.counters_path()
        if p.exists():
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                data = {}
        else:
            data = {}

        key = fmt_ddmmyyyy(d)
        last = int(data.get(key, 0))
        nxt = last + 1
        data[key] = nxt
        p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return nxt


    def mark_invalid(self, widget: QWidget, invalid: bool):
    # простая визуальная подсветка
        if invalid:
            widget.setStyleSheet("border: 2px solid #d9534f; border-radius: 4px;")
        else:
            widget.setStyleSheet("")

    def clear_all_marks(self):
        for _, (_, w) in self.widgets_by_key.items():
            self.mark_invalid(w, False)

    def validate_widgets(self):
        self.clear_all_marks()
        missing = []
        invalid = []

        for key, (f, w) in self.widgets_by_key.items():
            val = self.read_value(f, w)

            # 1) обязательность
            if f.required:
                if val is None:
                    missing.append(f.label)
                    self.mark_invalid(w, True)
                    continue
                if isinstance(val, str) and not val.strip():
                    missing.append(f.label)
                    self.mark_invalid(w, True)
                    continue

            # 2) формат (только на этапе генерации)
            if isinstance(w, QLineEdit):
                txt = w.text()
                # --- Жёсткая валидация заполненности/формата по ключам (на этапе генерации) ---
                def digits_only(s: str) -> str:
                    return "".join(ch for ch in s if ch.isdigit())

# поля строго по длине цифр
                len_rules = {
                    "passport_series": 4,
                    "passport_number": 6,
                    "inn": 12,
                    "bank_account": 20,
                    "bank_bik": 9,
                    "bank_corr": 20,
                    "ogrnip": 15,
                }

                if key in len_rules:
                    d = digits_only(txt)
                    if len(d) != len_rules[key]:
                        invalid.append(f.label)
                        self.mark_invalid(w, True)
                        continue

# формат NNN-NNN
                if key == "passport_code":
                    if not re.fullmatch(r"\d{3}-\d{3}", txt.strip()):
                        invalid.append(f.label)
                        self.mark_invalid(w, True)
                        continue

# название курса — без кавычек
                if key == "course_name":
                    if re.search(r'["«»“”]', txt):
                        invalid.append(f.label)
                        self.mark_invalid(w, True)
                        continue
                

# phone_email (одно поле: телефон ИЛИ email)
                if key == "phone_email":
                    phone_ok = re.fullmatch(r"\+[1-9] \(\d{3}\) \d{3}-\d{2}-\d{2}", txt.strip()) is not None
                    email_ok = re.fullmatch(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+\.[A-Za-z]{2,}", txt.strip()) is not None
                    if txt.strip() and (not phone_ok) and (not email_ok):
                        invalid.append(f.label)
                        self.mark_invalid(w, True)
                        continue
                # Маска: если поле required ИЛИ пользователь что-то вводил — требуем полностью валидный ввод
                if w.inputMask():
                    if (f.required or txt.strip()) and (not w.hasAcceptableInput()):
                        invalid.append(f.label)
                        self.mark_invalid(w, True)
                        continue

                # Regex-валидатор: проверяем только если что-то введено
                if w.validator() is not None:
                    if txt.strip() and (not w.hasAcceptableInput()):
                        invalid.append(f.label)
                        self.mark_invalid(w, True)
                        continue

        if missing or invalid:
            msg = []
            if missing:
                msg.append("Не заполнены обязательные поля:\n- " + "\n- ".join(missing))
            if invalid:
                msg.append("Некорректный формат полей:\n- " + "\n- ".join(invalid))

            QMessageBox.warning(self, "Проверка не пройдена", "\n\n".join(msg))
            return False

        return True

    def build_context(self):
        if not self.validate_widgets():
            return None

        raw = {}
        for key, (f, w) in self.widgets_by_key.items():
            raw[key] = self.read_value(f, w)

        # --- contract numbering & dates ---
        contract_date: date = raw.get("contract_date", date.today())

        seq = raw.get("contract_seq", 0)
        if isinstance(seq, int) and seq == 0:
            seq = self.next_contract_seq_for_date(contract_date)
            raw["contract_seq"] = seq

        contract_number = f"{seq}-{fmt_ddmmyyyy(contract_date)}"
        contract_date_long = fmt_long_ru(contract_date)
        appendix_header = f"от {contract_date_long} №{contract_number}"

        # Base context: all dates as dd.mm.yyyy
        context = {}
        for k, v in raw.items():
            if isinstance(v, date):
                context[k] = fmt_ddmmyyyy(v)
            elif isinstance(v, float):
                context[k] = f"{v:.2f}"
            else:
                context[k] = v

        # computed: main
        context["contract_number"] = contract_number
        context["contract_date_long"] = contract_date_long
        context["appendix_header"] = appendix_header

        # computed: work_end long format
        work_end_date = raw.get("work_end")
        if isinstance(work_end_date, date):
            context["work_end_long"] = fmt_long_ru(work_end_date)
            context["work_end_ddmmyyyy"] = fmt_ddmmyyyy(work_end_date)
        else:
            context["work_end_long"] = ""
            context["work_end_ddmmyyyy"] = ""

        # computed: money full forms (нужно: 50 000 (пятьдесят тысяч) рублей 00 копеек)
        total_val = raw.get("total_sum", 0.0)
        if isinstance(total_val, (int, float)):
            total_val = float(total_val)
            rub = int(total_val)
            kop = int(round((total_val - rub) * 100))
            if kop == 100:
                rub += 1
                kop = 0

            rub_num = f"{rub:,}".replace(",", " ")              # 50 000
            rub_words = int_to_words_ru(rub)                     # пятьдесят тысяч
            rub_unit = _morph(rub, "рубль", "рубля", "рублей")   # рублей
            kop_unit = _morph(kop, "копейка", "копейки", "копеек")

            context["total_sum_num"] = rub_num
            context["total_sum_words"] = f"{rub_words} {rub_unit} {kop:02d} {kop_unit}"
            context["total_sum_full"] = f"{rub_num} ({rub_words}) {rub_unit} {kop:02d} {kop_unit}"
        else:
            context["total_sum_num"] = ""
            context["total_sum_words"] = ""
            context["total_sum_full"] = ""

        adv_val = raw.get("advance_sum", 0.0)
        if isinstance(adv_val, (int, float)):
            adv_val = float(adv_val)
            rub = int(adv_val)
            kop = int(round((adv_val - rub) * 100))
            if kop == 100:
                rub += 1
                kop = 0

            rub_num = f"{rub:,}".replace(",", " ")
            rub_words = int_to_words_ru(rub)
            rub_unit = _morph(rub, "рубль", "рубля", "рублей")
            kop_unit = _morph(kop, "копейка", "копейки", "копеек")

            context["advance_sum_num"] = rub_num
            context["advance_sum_words"] = f"{rub_words} {rub_unit} {kop:02d} {kop_unit}"
            context["advance_sum_full"] = f"{rub_num} ({rub_words}) {rub_unit} {kop:02d} {kop_unit}"
        else:
            context["advance_sum_num"] = ""
            context["advance_sum_words"] = ""
            context["advance_sum_full"] = ""

        # final_part as sum num only
        final_val = raw.get("final_part", None)
        if isinstance(final_val, (int, float)):
            rub = int(round(float(final_val)))
            context["final_part_num"] = f"{rub:,}".replace(",", " ")
        else:
            context["final_part_num"] = ""


        def fmt_money_num_ru(v: float) -> str:
            """Число в RU-формате: 45 000 или 45 000,25"""
            v = round(float(v), 2)
            rub = int(v)
            kop = int(round((v - rub) * 100))
            if kop == 100:
                rub += 1
                kop = 0
            rub_part = f"{rub:,}".replace(",", " ")
            if kop == 0:
                return rub_part
            return f"{rub_part},{kop:02d}"

        def money_full_ru(v: float) -> str:
            """45 000 (сорок пять тысяч) рублей 00 копеек"""
            v = round(float(v), 2)
            rub = int(v)
            kop = int(round((v - rub) * 100))
            if kop == 100:
                rub += 1
                kop = 0

            rub_num = fmt_money_num_ru(v)
            rub_words = int_to_words_ru(rub)
            rub_unit = _morph(rub, "рубль", "рубля", "рублей")
            kop_unit = _morph(kop, "копейка", "копейки", "копеек")

            return f"{rub_num} ({rub_words}) {rub_unit} {kop:02d} {kop_unit}"
        # --- НДС включён в стоимость ---
        vat_rate = raw.get("vat_rate", 0.0)
        try:
            vat_rate = float(vat_rate)
        except Exception:
            vat_rate = 0.0

        total_val = raw.get("total_sum", 0.0)
        total_val = float(total_val) if isinstance(total_val, (int, float)) else 0.0

        if vat_rate > 0 and total_val > 0:
            vat_amount = round(total_val * vat_rate / (100.0 + vat_rate), 2)
            context["vat_full"] = money_full_ru(vat_amount)
        else:
            context["vat_full"] = ""

        
        # initials
        init = str(raw.get("initials", "")).strip()
        fio = str(raw.get("fio", "")).strip()
        context["initials"] = init if init else fio_to_initials(fio)

        # appendices
        app_cnt = raw.get("appendix_count", 3)
        try:
            app_cnt = int(app_cnt)
        except Exception:
            app_cnt = 3
        app_cnt = max(0, min(app_cnt, 20))
        for i in range(1, app_cnt + 1):
            context[f"appendix_{i}_header"] = appendix_header

        # --- VAT / НДС ---
        vat_rate = raw.get("vat_rate", 0.0)
        try:
            vat_rate = float(vat_rate)
        except Exception:
            vat_rate = 0.0

        base_total = raw.get("total_sum", 0.0)
        base_total = float(base_total) if isinstance(base_total, (int, float)) else 0.0
        vat_amount = round(base_total * vat_rate / 100.0, 2)
        total_with_vat = round(base_total + vat_amount, 2)

        # В контекст (для DOCX)
        context["vat_rate"] = f"{vat_rate:.2f}".rstrip("0").rstrip(".")  # "20" или "10.5"
        context["vat_amount_num"] = f"{int(round(vat_amount)):,}".replace(",", " ") if vat_amount.is_integer() else f"{vat_amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
        context["total_with_vat_num"] = f"{int(round(total_with_vat)):,}".replace(",", " ") if total_with_vat.is_integer() else f"{total_with_vat:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")

        return context

    def generate(self):
        if not self.current_schema_name:
            return

        schema = self.schemas[self.current_schema_name]
        template_path: Path = schema["template"]

        if not template_path.exists():
            QMessageBox.critical(
                self, "Нет шаблона DOCX",
                f"Не найден файл:\n{template_path}\n\nОжидается templates/{self.current_schema_name}.docx"
            )
            return

        context = self.build_context()
        if context is None:
            return

        fio_part = str(context.get("fio", "")).strip().replace(" ", "_")
        cn_part = str(context.get("contract_number", self.current_schema_name)).replace(" ", "_")
        default_name = f"{cn_part}_{fio_part}".strip("_")
        if not default_name:
            default_name = self.current_schema_name
        default_path = OUTPUT_DIR / f"{default_name}.docx"

        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить договор", str(default_path), "Word (*.docx)"
        )
        if not save_path:
            return

        try:
            doc = DocxTemplate(str(template_path))
            doc.render(context)
            doc.save(save_path)
            QMessageBox.information(self, "Готово", f"Договор создан:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка генерации", f"Не удалось создать документ.\n\n{e}")


if __name__ == "__main__":
    VERSION = "2026-02-26-PATCH-4"
    app = QApplication([])
    app.setStyleSheet(APP_QSS)
    if sys.platform.startswith("win"):
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("company.contract.templater")
    w = MainWindow()
    icon_path = ASSETS_DIR / "app.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
        w.setWindowIcon(QIcon(str(icon_path)))
    w.setWindowTitle(w.windowTitle())
    w.show()
    app.exec()
