from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

ONT_PATH = Path(r"C:\Users\welta\OneDrive\Desktop\Универ\Курсовая\Перенос данных в  онтологию\Белковые продукты, Кондитерские изделия, Рыбная продукция.ont")
XLSX_PATH = Path(r"C:\Users\welta\OneDrive\Desktop\Универ\Курсовая\Перенос данных в  онтологию\11.04 вырезка по белковым продуктам и сладостям.xlsx")

OUTPUT_ONT_PATH = Path(r"C:\Users\welta\OneDrive\Desktop\Универ\Курсовая\Перенос данных в  онтологию\УЛ Белковые продукты, Кондитерские изделия, Рыбная продукция.ont")
OUTPUT_XLSX_PATH = Path(r"C:\Users\welta\OneDrive\Desktop\Универ\Курсовая\Перенос данных в  онтологию\УЛ 11.04 вырезка по белковым продуктам и сладостям_updated.xlsx")

def load_ontology(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)

def save_ontology(data: dict, path: Path) -> None:
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def normalize_text(text: object) -> str:
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = text.replace("ё", "е")
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .,;:-\"'()[]{}")

def safe_str(value: object) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()

def next_id(items: List[dict]) -> str:
    max_id = 0
    for item in items:
        try:
            max_id = max(max_id, int(item.get("id", 0)))
        except (TypeError, ValueError):
            continue
    return str(max_id + 1)

def get_namespace(ontology: dict) -> str:
    nodes = ontology.get("nodes", [])
    if nodes:
        return nodes[0].get("namespace", "")
    return ""

def build_node_index(ontology: dict) -> Dict[str, dict]:
    index = {}
    for node in ontology.get("nodes", []):
        norm_name = normalize_text(node.get("name", ""))
        if norm_name:
            index[norm_name] = node
    return index

def relation_exists(ontology: dict, source_id: str, destination_id: str, rel_name: str = "is_a") -> bool:
    for rel in ontology.get("relations", []):
        if (
            rel.get("source_node_id") == source_id
            and rel.get("destination_node_id") == destination_id
            and rel.get("name") == rel_name
        ):
            return True
    return False

def create_node(ontology: dict, name: str, parent_node: dict) -> dict:
    node_id = next_id(ontology["nodes"])

    siblings_count = sum(
        1
        for rel in ontology.get("relations", [])
        if rel.get("destination_node_id") == parent_node["id"] and rel.get("name") == "is_a"
    )

    new_node = {
        "attributes": {},
        "id": node_id,
        "name": name,
        "namespace": parent_node.get("namespace", get_namespace(ontology)),
        "position_x": int(parent_node.get("position_x", 0)) + 220,
        "position_y": int(parent_node.get("position_y", 0)) - 80 + siblings_count * 45,
    }
    ontology["nodes"].append(new_node)
    return new_node

def create_is_a_relation(ontology: dict, child_id: str, parent_id: str) -> dict:
    rel_id = next_id(ontology["relations"])
    relation = {
        "attributes": {},
        "destination_node_id": parent_id,
        "id": rel_id,
        "name": "is_a",
        "namespace": get_namespace(ontology),
        "source_node_id": child_id,
    }
    ontology["relations"].append(relation)
    return relation

def get_product_name_from_row(row: pd.Series) -> str:
    for col in ["Перевод", "Название продукта"]:
        if col in row.index:
            value = safe_str(row[col])
            if value:
                return value
    return ""

def find_exact_node(product_name: str, node_index: Dict[str, dict]) -> Optional[dict]:
    return node_index.get(normalize_text(product_name))

def find_parent_node_for_new_product(product_name: str, node_index: Dict[str, dict]) -> Optional[dict]:
    product_norm = normalize_text(product_name)
    candidates = []

    for node_name_norm, node in node_index.items():
        if product_norm.startswith(node_name_norm + ",") or product_norm.startswith(node_name_norm + " "):
            candidates.append((len(node_name_norm), node))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]

def write_id_to_column_b(df: pd.DataFrame, row_index: int, node_id: str) -> None:
    if len(df.columns) < 2:
        return
    column_b = df.columns[1]
    df.at[row_index, column_b] = node_id

def process_excel_and_ontology(ontology: dict, excel_path: Path) -> Tuple[dict, Dict[str, pd.DataFrame]]:
    sheets = pd.read_excel(excel_path, sheet_name=None, engine="openpyxl")
    node_index = build_node_index(ontology)

    for _, df in sheets.items():
        if df.empty:
            continue

        for row_index, row in df.iterrows():
            product_name = get_product_name_from_row(row)
            if not product_name:
                continue

            exact_node = find_exact_node(product_name, node_index)

            if exact_node is not None:
                write_id_to_column_b(df, row_index, exact_node["id"])
                continue

            parent_node = find_parent_node_for_new_product(product_name, node_index)
            if parent_node is None:
                continue

            new_node = create_node(ontology, product_name, parent_node)
            create_is_a_relation(ontology, new_node["id"], parent_node["id"])

            node_index[normalize_text(product_name)] = new_node
            write_id_to_column_b(df, row_index, new_node["id"])

    return ontology, sheets

def save_excel(sheets: Dict[str, pd.DataFrame], output_path: Path) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

def main() -> None:
    if not ONT_PATH.exists():
        raise FileNotFoundError(f"Не найден файл онтологии: {ONT_PATH}")

    if not XLSX_PATH.exists():
        raise FileNotFoundError(f"Не найден Excel-файл: {XLSX_PATH}")

    ontology = load_ontology(ONT_PATH)
    updated_ontology, updated_sheets = process_excel_and_ontology(ontology, XLSX_PATH)

    save_ontology(updated_ontology, OUTPUT_ONT_PATH)
    save_excel(updated_sheets, OUTPUT_XLSX_PATH)

    print(f"Готово. Обновлённая онтология: {OUTPUT_ONT_PATH}")
    print(f"Готово. Обновлённый Excel: {OUTPUT_XLSX_PATH}")

if __name__ == "__main__":
    main()