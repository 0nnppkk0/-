from __future__ import annotations

import json
import math
import os
import re
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Dict, List, Optional, Tuple

import pandas as pd


ONT_PATH = r"C:\Users\nnppk\OneDrive\Рабочий стол\ВУЗ\3k\Курсовая (Чуприна)\Онтологически управляемое решение demo v1\document.ont"
EXCEL_PATH = r"C:\Users\nnppk\OneDrive\Рабочий стол\ВУЗ\3k\Курсовая (Чуприна)\Онтологически управляемое решение demo v1\document.xlsx"

OUTPUT_DIR = r"C:\Users\nnppk\OneDrive\Рабочий стол\ВУЗ\3k\Курсовая (Чуприна)\Онтологически управляемое решение demo v1"

PRODUCT_COLUMN = "Продукт"
OUTPUT_XLSX = "опрос_и_анализ.xlsx"


@dataclass
class NormRecord:
    nutrient: str
    minimum: float
    maximum: float


class OntologyParser:
    def __init__(self, ont_path: str):
        self.ont_path = ont_path
        self.nodes: Dict[str, Dict] = {}
        self.relations: List[Dict] = []
        self.name_to_id: Dict[str, str] = {}
        self.children_map: Dict[str, List[str]] = {}
        self.parent_map: Dict[str, str] = {}
        self.order_map: Dict[str, int] = {}
        self.top_categories: List[str] = []
        self.question_ids: Dict[str, str] = {}
        self.variant_map: Dict[str, List[str]] = {}
        self._is_a_names = {"is_a", "isa"}
        self._load_ontology()

    def _load_ontology(self) -> None:
        if not os.path.exists(self.ont_path):
            raise FileNotFoundError(f"Файл онтологии не найден: {self.ont_path}")

        with open(self.ont_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.nodes = {str(node["id"]): node for node in data.get("nodes", [])}
        self.relations = data.get("relations", [])
        self.name_to_id = {
            str(node.get("name", "")).strip(): str(node["id"])
            for node in data.get("nodes", [])
        }

        category_pairs: List[Tuple[int, str]] = []

        for relation in self.relations:
            name = str(relation.get("name", "")).strip()
            src = str(relation.get("source_node_id"))
            dst = str(relation.get("destination_node_id"))

            if name in self._is_a_names:
                self.children_map.setdefault(dst, []).append(src)
                self.parent_map[src] = dst
            elif name == "order":
                order_name = self.get_name(dst).strip()
                self.order_map[src] = self._safe_int(order_name, 999999)
            elif name in {"-1", "-2", "-3"}:
                try:
                    category_pairs.append((int(name), dst))
                except ValueError:
                    pass
            elif name == "start":
                self.question_ids["consume"] = dst
            elif name == "next":
                self.question_ids.setdefault("next_links", []).append((src, dst))
            elif name == "variant":
                self.variant_map.setdefault(src, []).append(dst)

        self.top_categories = [node_id for _, node_id in sorted(category_pairs, key=lambda item: item[0])]
        self._resolve_question_chain()
        self._sort_all_children()

    @staticmethod
    def _safe_int(value: str, default: int) -> int:
        value = str(value).strip()
        return int(value) if value.isdigit() else default

    def _resolve_question_chain(self) -> None:
        yes_id = self.name_to_id.get("Да")
        if not yes_id:
            raise ValueError('В онтологии не найден узел "Да"')

        if "consume" not in self.question_ids:
            raise ValueError('В онтологии не найден стартовый вопрос через связь start')

        next_links = dict(self.question_ids.get("next_links", []))
        frequency_id = next_links.get(yes_id)
        if not frequency_id:
            raise ValueError('Не найден переход next от узла "Да" к вопросу частоты')

        portion_id = next_links.get(frequency_id)
        if not portion_id:
            raise ValueError('Не найден переход next от вопроса частоты к вопросу порции')

        self.question_ids["frequency"] = frequency_id
        self.question_ids["portion"] = portion_id

    def _sort_all_children(self) -> None:
        for parent_id, child_ids in self.children_map.items():
            self.children_map[parent_id] = sorted(
                child_ids,
                key=lambda child_id: (self.order_map.get(child_id, 999999), self.get_name(child_id))
            )

    def get_name(self, node_id: str) -> str:
        return str(self.nodes.get(str(node_id), {}).get("name", "Неизвестно")).strip()

    def get_children(self, node_id: str) -> List[str]:
        return self.children_map.get(str(node_id), [])

    def get_parent(self, node_id: str) -> Optional[str]:
        return self.parent_map.get(str(node_id))

    def has_order(self, node_id: str) -> bool:
        return str(node_id) in self.order_map

    def get_question_text(self, question_key: str, product_name: str) -> str:
        question_id = self.question_ids[question_key]
        return self.get_name(question_id).replace("[]", product_name)

    def get_variants(self, question_key: str) -> List[str]:
        question_id = self.question_ids[question_key]
        variants = self.variant_map.get(question_id, [])
        return [self.get_name(node_id) for node_id in variants]

    def is_terminal_product(self, node_id: str) -> bool:
        return self.get_parent(node_id) is not None and not self.has_order(node_id)

    def get_terminal_products(self) -> List[Tuple[str, str]]:
        result: List[Tuple[str, str]] = []
        visited: set[str] = set()

        def ordered_children(node_id: str) -> List[str]:
            return sorted(
                self.get_children(node_id),
                key=lambda child_id: (self.order_map.get(child_id, 999999), self.get_name(child_id))
            )

        def walk(node_id: str, top_category_id: str) -> None:
            if node_id in visited:
                return
            visited.add(node_id)

            if self.is_terminal_product(node_id):
                result.append((node_id, top_category_id))
                return

            for child_id in ordered_children(node_id):
                walk(child_id, top_category_id)

        for category_id in self.top_categories:
            for child_id in ordered_children(category_id):
                walk(child_id, category_id)

        return result


class NutrientRepository:
    def __init__(self, excel_path: str, product_column: str = PRODUCT_COLUMN):
        self.excel_path = excel_path
        self.product_column = product_column
        self.products_df: pd.DataFrame = pd.DataFrame()
        self.product_lookup: Dict[str, Dict] = {}
        self.norms: Dict[str, NormRecord] = {}
        self._load()

    def _load(self) -> None:
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"Excel файл не найден: {self.excel_path}")

        xls = pd.ExcelFile(self.excel_path)
        sheet_names = xls.sheet_names
        if not sheet_names:
            raise ValueError("В Excel нет листов")

        norms_sheet = None
        data_sheets: List[str] = []

        for sheet in sheet_names:
            if str(sheet).strip().lower() == "нормы":
                norms_sheet = sheet
            else:
                data_sheets.append(sheet)

        if norms_sheet is None:
            raise ValueError('В Excel не найден лист "Нормы"')

        product_frames: List[pd.DataFrame] = []
        for sheet in data_sheets:
            df = pd.read_excel(self.excel_path, sheet_name=sheet)
            if df.empty:
                continue

            df = self._normalize_columns(df)
            product_col = self._detect_product_column(df)
            if not product_col:
                continue

            df = df.rename(columns={product_col: self.product_column})
            df[self.product_column] = df[self.product_column].astype(str).str.strip()
            df = df[df[self.product_column] != ""]
            df = df[df[self.product_column].str.lower() != "nan"]

            if not df.empty:
                product_frames.append(df)

        if not product_frames:
            raise ValueError("Не удалось найти продуктовые таблицы в Excel")

        self.products_df = pd.concat(product_frames, ignore_index=True, sort=False)
        self.products_df = self.products_df.loc[:, ~self.products_df.columns.duplicated()]

        for _, row in self.products_df.iterrows():
            product_name = str(row[self.product_column]).strip()
            if product_name:
                self.product_lookup[self._normalize_product_key(product_name)] = row.to_dict()

        self._load_norms(norms_sheet)

    def _load_norms(self, sheet_name: str) -> None:
        norms_df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)

        if norms_df.shape[1] < 3:
            raise ValueError('На листе "Нормы" должно быть минимум 3 столбца: нутриент, минимум, максимум')

        for _, row in norms_df.iterrows():
            nutrient = str(row.iloc[0]).strip()
            if not nutrient or nutrient.lower() == "nan":
                continue

            minimum = self._to_float(row.iloc[1])
            maximum = self._to_float(row.iloc[2])
            self.norms[nutrient] = NormRecord(
                nutrient=nutrient,
                minimum=minimum,
                maximum=maximum
            )

        if not self.norms:
            raise ValueError('На листе "Нормы" не найдено ни одной валидной строки')

    @staticmethod
    def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
        cols = []
        for idx, col in enumerate(df.columns):
            name = str(col).strip()
            cols.append(name if name and name.lower() != "nan" else f"column_{idx}")
        df = df.copy()
        df.columns = cols
        return df

    @staticmethod
    def _normalize_product_key(name: str) -> str:
        name = str(name).strip().lower().replace("ё", "е")
        return re.sub(r"\s+", " ", name)

    def _detect_product_column(self, df: pd.DataFrame) -> Optional[str]:
        candidates = []
        for col in df.columns:
            values = df[col].dropna().astype(str).str.strip()
            if values.empty:
                continue

            ratio = ((values != "") & (values.str.lower() != "nan")).mean()
            if ratio < 0.5:
                continue

            if any(token in col.lower() for token in ["проду", "name", "food", "назв"]):
                return col

            candidates.append((col, ratio))

        return candidates[0][0] if candidates else None

    @staticmethod
    def _to_float(value) -> float:
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)

        value = str(value).strip().replace(",", ".")
        if not value or value.lower() == "nan":
            return 0.0

        match = re.search(r"-?\d+(?:\.\d+)?", value)
        return float(match.group()) if match else 0.0

    def get_product_data(self, product_name: str) -> Optional[Dict]:
        key = self._normalize_product_key(product_name)
        if key in self.product_lookup:
            return self.product_lookup[key]

        for candidate_key, row in self.product_lookup.items():
            if key == candidate_key or key in candidate_key or candidate_key in key:
                return row

        return None

    def get_norm_names(self) -> List[str]:
        return list(self.norms.keys())


class NutrientAnalyzer:
    def __init__(self, repository: NutrientRepository):
        self.repository = repository

    @staticmethod
    def _to_float(value) -> float:
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)

        text = str(value).strip().replace(",", ".")
        if not text or text.lower() == "nan":
            return 0.0

        match = re.search(r"-?\d+(?:\.\d+)?", text)
        return float(match.group()) if match else 0.0

    def calculate(self, answers: List[Dict]) -> Tuple[pd.DataFrame, bool, List[str]]:
        rows: List[Dict] = []
        nutrient_totals: Dict[str, float] = {
            name: 0.0 for name in self.repository.get_norm_names()
        }

        for answer in answers:
            if answer.get("Употребление") != "Да":
                continue

            product_name = str(answer.get("Продукт", "")).strip()
            portion = self._to_float(answer.get("Порция", 0))
            frequency = self._to_float(answer.get("Частота", 0))
            product_data = self.repository.get_product_data(product_name)

            if not product_data:
                continue

            for nutrient_name, norm in self.repository.norms.items():
                nutrient_per_100g = self._to_float(product_data.get(nutrient_name, 0))
                consumed_amount_per_day = (portion / 100.0) * frequency * nutrient_per_100g / 30.0
                nutrient_totals[nutrient_name] += consumed_amount_per_day

        for nutrient_name, norm in self.repository.norms.items():
            total = nutrient_totals.get(nutrient_name, 0.0)
            status = "Норма"

            if total < norm.minimum:
                status = "Ниже нормы"
            elif norm.maximum > 0 and total > norm.maximum:
                status = "Выше нормы"

            rows.append({
                "Нутриент": nutrient_name,
                "Получено_в_день": round(total, 4),
                "Минимум_в_день": norm.minimum,
                "Максимум_в_день": norm.maximum,
                "Статус": status,
            })

        result_df = pd.DataFrame(rows)
        bad_rows = result_df[result_df["Статус"] != "Норма"]
        balanced = bad_rows.empty
        bad_nutrients = bad_rows["Нутриент"].tolist()

        return result_df, balanced, bad_nutrients


def answer_to_number(value: str) -> float:
    text = str(value).strip().replace(",", ".")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group()) if match else 0.0


def autosize_worksheet(worksheet, dataframe: pd.DataFrame) -> None:
    for idx, col in enumerate(dataframe.columns, start=1):
        values = [str(col)]
        if not dataframe.empty:
            values.extend(dataframe[col].astype(str).tolist())

        max_len = max(len(v) for v in values) if values else 10
        column_letter = worksheet.cell(row=1, column=idx).column_letter
        worksheet.column_dimensions[column_letter].width = min(max(max_len + 2, 12), 40)


class SurveyApp:
    def __init__(self, ont_path: str, excel_path: str, product_column: str = PRODUCT_COLUMN):
        self.ont_path = ont_path
        self.excel_path = excel_path
        self.product_column = product_column

        self.parser = OntologyParser(ont_path)
        self.nutrients = NutrientRepository(excel_path, product_column)
        self.analyzer = NutrientAnalyzer(self.nutrients)

        self.ont_filename = os.path.basename(ont_path)
        self.excel_filename = os.path.basename(excel_path)

        self.products = self.parser.get_terminal_products()
        if not self.products:
            raise ValueError("В онтологии не найдены конечные продукты для опроса")

        self.answers: List[Dict] = []
        self.current_index = 0
        self.current_state = "consume"
        self.current_product_id = ""
        self.current_category_id = ""
        self.temp_frequency = ""

        self.root = tk.Tk()
        self.root.title("Онтологический опросник")
        self.root.geometry("1100x780")
        self.root.resizable(True, True)

        self.setup_ui()

    def setup_ui(self):
        header = tk.Frame(self.root)
        header.pack(fill="x", padx=20, pady=10)

        title = tk.Label(
            header,
            text="Онтологически управляемый опросник",
            font=("Arial", 16, "bold"),
            fg="#1E3A8A"
        )
        title.pack(anchor="w")

        files_info = tk.Label(
            header,
            text=f"ONT: {self.ont_filename} | Excel: {self.excel_filename}",
            font=("Arial", 10),
            fg="gray"
        )
        files_info.pack(anchor="w", pady=(4, 0))

        top_buttons = tk.Frame(header)
        top_buttons.pack(anchor="e", pady=(8, 0))
        ttk.Button(top_buttons, text="Перезапуск", command=self.restart).pack(side="left", padx=5)

        self.survey_frame = ttk.Frame(self.root)
        self.survey_frame.pack(fill="both", expand=True, padx=25, pady=15)

        self.question_lbl = tk.Label(
            self.survey_frame,
            text="",
            wraplength=980,
            font=("Arial", 14),
            justify="center",
            pady=10
        )
        self.question_lbl.pack(pady=20)

        self.options_lb = tk.Listbox(
            self.survey_frame,
            height=10,
            font=("Arial", 12),
            selectmode="single",
            relief="solid",
            bd=2,
            exportselection=False
        )
        self.options_lb.pack(fill="both", expand=True, padx=60, pady=15)
        self.options_lb.bind("<Double-Button-1>", lambda event: self.process_answer())

        btn_frame = ttk.Frame(self.survey_frame)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="Далее", command=self.process_answer).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Пропустить", command=self.skip_product).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Завершить", command=self.finish).pack(side="left", padx=10)

        self.results_frame = ttk.Frame(self.root)

        self.summary_label = tk.Label(
            self.results_frame,
            text="Итоги опроса",
            font=("Arial", 14, "bold"),
            fg="#1E3A8A"
        )
        self.summary_label.pack(anchor="w", padx=10, pady=(10, 5))

        self.answer_table = self._create_tree(
            self.results_frame,
            columns=("Категория", "Продукт", "Употребление", "Частота", "Порция"),
            height=12
        )
        self.answer_table.pack(fill="both", expand=False, padx=10, pady=5)

        self.analysis_label = tk.Label(
            self.results_frame,
            text="Анализ нутриентов",
            font=("Arial", 14, "bold"),
            fg="#1E3A8A"
        )
        self.analysis_label.pack(anchor="w", padx=10, pady=(15, 5))

        self.analysis_table = self._create_tree(
            self.results_frame,
            columns=("Нутриент", "Получено_в_день", "Минимум_в_день", "Максимум_в_день", "Статус"),
            height=14
        )
        self.analysis_table.pack(fill="both", expand=True, padx=10, pady=5)

        self.result_status = tk.Label(
            self.results_frame,
            text="",
            font=("Arial", 11, "bold"),
            justify="left"
        )
        self.result_status.pack(anchor="w", padx=10, pady=10)

        self.status_var = tk.StringVar(value="Опрос готов к запуску")
        tk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w").pack(side="bottom", fill="x")

        self.next_product()

    def _create_tree(self, parent, columns: Tuple[str, ...], height: int = 10):
        frame = ttk.Frame(parent)

        tree = ttk.Treeview(frame, columns=columns, show="headings", height=height)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)

        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=180, anchor="center")

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.tree = tree
        return frame

    def restart(self):
        self.parser = OntologyParser(self.ont_path)
        self.nutrients = NutrientRepository(self.excel_path, self.product_column)
        self.analyzer = NutrientAnalyzer(self.nutrients)

        self.products = self.parser.get_terminal_products()
        self.answers = []
        self.current_index = 0
        self.current_state = "consume"
        self.current_product_id = ""
        self.current_category_id = ""
        self.temp_frequency = ""

        self.results_frame.pack_forget()
        self.survey_frame.pack(fill="both", expand=True, padx=25, pady=15)
        self.next_product()

    def next_product(self):
        if self.current_index >= len(self.products):
            self.finish()
            return

        self.current_product_id, self.current_category_id = self.products[self.current_index]
        product_name = self.parser.get_name(self.current_product_id)
        category_name = self.parser.get_name(self.current_category_id)

        self.current_state = "consume"
        self.question_lbl.config(text=self.parser.get_question_text("consume", product_name))
        self._fill_options(self.parser.get_variants("consume"))

        self.status_var.set(
            f"{self.current_index + 1}/{len(self.products)} | Категория: {category_name} | Продукт: {product_name}"
        )

    def _fill_options(self, options: List[str]) -> None:
        self.options_lb.delete(0, tk.END)
        for option in options:
            self.options_lb.insert(tk.END, option)
        if options:
            self.options_lb.selection_set(0)

    def process_answer(self):
        selection = self.options_lb.curselection()
        if not selection:
            return messagebox.showerror("Ошибка", "Выберите вариант ответа")

        answer = self.options_lb.get(selection[0]).strip()
        product_name = self.parser.get_name(self.current_product_id)
        category_name = self.parser.get_name(self.current_category_id)

        if self.current_state == "consume":
            if answer == "Да":
                self.current_state = "frequency"
                self.question_lbl.config(text=self.parser.get_question_text("frequency", product_name))
                self._fill_options(self.parser.get_variants("frequency"))
            else:
                self.answers.append({
                    "Категория": category_name,
                    "Продукт": product_name,
                    "Употребление": answer,
                    "Частота": "",
                    "Порция": ""
                })
                self.current_index += 1
                self.next_product()

        elif self.current_state == "frequency":
            self.temp_frequency = answer
            self.current_state = "portion"
            self.question_lbl.config(text=self.parser.get_question_text("portion", product_name))
            self._fill_options(self.parser.get_variants("portion"))

        elif self.current_state == "portion":
            row = {
                "Категория": category_name,
                "Продукт": product_name,
                "Употребление": "Да",
                "Частота": answer_to_number(self.temp_frequency),
                "Порция": answer_to_number(answer)
            }

            product_data = self.nutrients.get_product_data(product_name)
            if product_data:
                for key, value in product_data.items():
                    if key != self.product_column:
                        row[str(key)] = value

            self.answers.append(row)
            self.current_index += 1
            self.next_product()

    def skip_product(self):
        product_name = self.parser.get_name(self.current_product_id)
        category_name = self.parser.get_name(self.current_category_id)

        self.answers.append({
            "Категория": category_name,
            "Продукт": product_name,
            "Употребление": "Пропущено",
            "Частота": "",
            "Порция": ""
        })

        self.current_index += 1
        self.next_product()

    def _clear_tree(self, tree_frame):
        tree = tree_frame.tree
        for item in tree.get_children():
            tree.delete(item)
        return tree

    def _fill_tree(self, tree_frame, dataframe: pd.DataFrame):
        tree = self._clear_tree(tree_frame)
        for _, row in dataframe.iterrows():
            values = [row[col] if col in dataframe.columns else "" for col in tree["columns"]]
            tree.insert("", "end", values=values)

    def _save_excel_report(self, answers_df: pd.DataFrame, analysis_df: pd.DataFrame) -> str:
        if not OUTPUT_DIR:
            raise ValueError("Не указана папка сохранения OUTPUT_DIR")

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        output_path = os.path.join(OUTPUT_DIR, OUTPUT_XLSX)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            answers_df.to_excel(writer, index=False, sheet_name="Sheet1")
            analysis_df.to_excel(writer, index=False, sheet_name="Sheet2")

            ws1 = writer.sheets["Sheet1"]
            ws2 = writer.sheets["Sheet2"]

            autosize_worksheet(ws1, answers_df)
            autosize_worksheet(ws2, analysis_df)

        return os.path.abspath(output_path)

    def finish(self):
        if not self.answers:
            return messagebox.showwarning("Внимание", "Пройдите опрос")

        answers_df = pd.DataFrame(self.answers)
        analysis_df, balanced, bad_nutrients = self.analyzer.calculate(self.answers)
        output_path = self._save_excel_report(answers_df, analysis_df)

        answer_view_df = answers_df[["Категория", "Продукт", "Употребление", "Частота", "Порция"]].fillna("")
        analysis_view_df = analysis_df.fillna("")

        self._fill_tree(self.answer_table, answer_view_df)
        self._fill_tree(self.analysis_table, analysis_view_df)

        if balanced:
            status_text = f"Итог: питание сбалансированное\nФайл отчета: {output_path}"
            self.result_status.config(text=status_text, fg="#166534")
        else:
            status_text = (
                f"Итог: питание несбалансированное\n"
                f"Нутриенты вне диапазона: {', '.join(bad_nutrients)}\n"
                f"Файл отчета: {output_path}"
            )
            self.result_status.config(text=status_text, fg="#991B1B")

        self.survey_frame.pack_forget()
        self.results_frame.pack(fill="both", expand=True, padx=25, pady=15)
        messagebox.showinfo("Готово", "Опрос завершен и Excel-отчет сформирован")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    try:
        app = SurveyApp(ONT_PATH, EXCEL_PATH, PRODUCT_COLUMN)
        app.run()
    except Exception as error:
        print(f"Ошибка: {error}")
        input("Нажмите Enter для выхода...")