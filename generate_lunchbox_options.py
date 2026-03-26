import csv
import json
import re
from collections import defaultdict
from copy import deepcopy
from pathlib import Path


SOURCE_CSV = Path("current-product-options.csv")
OUTPUT_CSV = Path("current-product-options-lunchbox.csv")
OUTPUT_MAP = Path("lunchbox-field-map.json")

ORDINALS = {
    1: "First",
    2: "Second",
    3: "Third",
    4: "Fourth",
    5: "Fifth",
    6: "Sixth",
}


def array_unserialize(value):
    parts = value.split("|")
    result = []
    for item in parts:
        match = re.match(r"<!\[CDATA\[(.*?)\]\]>", item, re.S)
        if match:
            item = match.group(1).replace("%7C", "|")
        result.append(item)
    return result


def array_serialize(value):
    if isinstance(value, list):
        serialized = []
        for item in value:
            item = "" if item is None else str(item)
            if "|" in item:
                item = "<![CDATA[" + item.replace("|", "%7C") + "]]>"
            serialized.append(item)
        return "|".join(serialized)
    if value is None:
        return ""
    return str(value)


def json_compact(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def parse_csv_to_rows(path):
    with path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        return reader.fieldnames, list(reader)


def write_csv_rows(path, headers, rows):
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def parse_clean_data(rows):
    parsed = {}
    for row in rows:
        for key, value in row.items():
            parsed.setdefault(key, [])
            if key.startswith("multiple_"):
                values = array_unserialize(value)
                if key.endswith("checkboxes_options_default_value"):
                    parsed[key].append(values)
                elif key.endswith("options_default_value"):
                    parsed[key].append(values[0] if values else "")
                else:
                    parsed[key].append(values)
            else:
                parsed[key].append(value)

    remove_keys = [index for index, value in enumerate(parsed["sections"]) if value == ""]

    element_keys = {}
    for value in parsed["element_type"]:
        element_keys.setdefault(value, [])
        element_keys[value].append(len(element_keys[value]))

    clean = {}
    for key, values in parsed.items():
        if key in {"element_type", "div_size"}:
            continue

        split = key.split("_")
        if split[0] == "multiple" and len(split) > 1:
            element_key = split[1]
        else:
            element_key = split[0]

        clean.setdefault(key, [])
        for index, value in enumerate(values):
            if element_key in {"sections", "section"}:
                if index not in remove_keys:
                    clean[key].append(value)
            elif element_key == "variations" or (
                element_key in element_keys and index in element_keys[element_key]
            ):
                clean[key].append(value)

    clean["element_type"] = parsed["element_type"]
    clean["div_size"] = parsed["div_size"]
    return clean


def get_key_groups(headers):
    section_keys = []
    type_keys = defaultdict(list)

    for key in headers:
        if key in {"element_type", "div_size"}:
            continue

        if key == "section" or key.startswith("section_") or key == "sections" or key.startswith("sections_"):
            section_keys.append(key)
            continue

        split = key.split("_")
        element_key = split[1] if split[0] == "multiple" and len(split) > 1 else split[0]
        type_keys[element_key].append(key)

    return section_keys, dict(type_keys)


def build_sections_from_clean(clean, headers):
    section_keys, type_keys = get_key_groups(headers)
    sections = []
    section_offsets = []
    section_start_index = 0

    for section_index, _ in enumerate(clean["sections"]):
        section_data = {
            key: clean.get(key, [])[section_index] if section_index < len(clean.get(key, [])) else ""
            for key in section_keys
        }
        sections.append({"data": section_data, "elements": []})
        section_offsets.append(section_start_index)
        section_start_index += int(section_data.get("sections") or 0)

    occurrence_indices = defaultdict(int)
    current_section_index = 0

    for element_position, element_type in enumerate(clean["element_type"]):
        while (
            current_section_index + 1 < len(section_offsets)
            and element_position >= section_offsets[current_section_index + 1]
        ):
            current_section_index += 1

        occurrence_index = occurrence_indices[element_type]
        element_data = {
            key: clean.get(key, [])[occurrence_index]
            if occurrence_index < len(clean.get(key, []))
            else ""
            for key in type_keys.get(element_type, [])
        }
        sections[current_section_index]["elements"].append(
            {
                "type": element_type,
                "div_size": clean["div_size"][element_position],
                "data": element_data,
            }
        )
        occurrence_indices[element_type] += 1

    for section in sections:
        expected_length = int(section["data"].get("sections") or 0)
        actual_length = len(section["elements"])
        if expected_length != actual_length:
            raise ValueError(
                "Section "
                + section["data"].get("sections_internal_name", "<unknown>")
                + " expected "
                + str(expected_length)
                + " elements but found "
                + str(actual_length)
                + "."
            )

    return sections


def build_clean_data_from_sections(sections, headers):
    section_keys, type_keys = get_key_groups(headers)
    clean = {key: [] for key in headers}

    for section in sections:
        for key in section_keys:
            clean[key].append(section["data"].get(key, ""))

    for section in sections:
        for element in section["elements"]:
            clean["element_type"].append(element["type"])
            clean["div_size"].append(element["div_size"])
            for key in type_keys.get(element["type"], []):
                clean[key].append(element["data"].get(key, ""))

    return clean


def write_clean_csv(path, headers, clean):
    row_count = max((len(values) for values in clean.values()), default=0)
    rows = []

    for row_index in range(row_count):
        row = {}
        for header in headers:
            values = clean.get(header, [])
            value = values[row_index] if row_index < len(values) else ""
            row[header] = array_serialize(value)
        rows.append(row)

    write_csv_rows(path, headers, rows)


def new_unique_id(counter):
    counter[0] += 1
    return "lbxsimple260326" + str(counter[0]).zfill(4) + ".00000000"


def build_lunchbox_element(selectbox_template, child_number, unique_counter):
    field_id = new_unique_id(unique_counter)

    selectbox_data = deepcopy(selectbox_template["data"])
    selectbox_data["selectbox_internal_name"] = "Child " + str(child_number) + " Lunchbox Choice"
    selectbox_data["selectbox_header_size"] = "10"
    selectbox_data["selectbox_header_title"] = "Lunchbox Choice"
    selectbox_data["selectbox_header_title_mode"] = ""
    selectbox_data["selectbox_header_title_position"] = ""
    selectbox_data["selectbox_header_title_color"] = ""
    selectbox_data["selectbox_header_subtitle"] = "Select Vegetarian or Non-Vegetarian for this child."
    selectbox_data["selectbox_header_subtitle_position"] = ""
    selectbox_data["selectbox_header_subtitle_color"] = ""
    selectbox_data["selectbox_divider_type"] = "none"
    selectbox_data["selectbox_enabled"] = "1"
    selectbox_data["selectbox_required"] = "1"
    selectbox_data["selectbox_fee"] = ""
    selectbox_data["selectbox_hide_amount"] = ""
    selectbox_data["selectbox_text_before_price"] = ""
    selectbox_data["selectbox_text_after_price"] = ""
    selectbox_data["selectbox_quantity"] = ""
    selectbox_data["selectbox_quantity_min"] = ""
    selectbox_data["selectbox_quantity_max"] = ""
    selectbox_data["selectbox_quantity_step"] = ""
    selectbox_data["selectbox_quantity_default_value"] = ""
    selectbox_data["selectbox_use_url"] = ""
    selectbox_data["selectbox_changes_product_image"] = ""
    selectbox_data["selectbox_placeholder"] = "Select lunchbox choice"
    selectbox_data["multiple_selectbox_options_default_value"] = ""
    selectbox_data["multiple_selectbox_options_title"] = ["Non-Vegetarian", "Vegetarian"]
    selectbox_data["multiple_selectbox_options_image"] = ["", ""]
    selectbox_data["multiple_selectbox_options_imagec"] = ["", ""]
    selectbox_data["multiple_selectbox_options_imagep"] = ["", ""]
    selectbox_data["multiple_selectbox_options_imagel"] = ["", ""]
    selectbox_data["multiple_selectbox_options_value"] = ["Non-Vegetarian", "Vegetarian"]
    selectbox_data["multiple_selectbox_options_price"] = ["0", "0"]
    selectbox_data["multiple_selectbox_options_sale_price"] = ["", ""]
    selectbox_data["multiple_selectbox_options_price_type"] = ["", ""]
    selectbox_data["multiple_selectbox_options_description"] = ["", ""]
    selectbox_data["multiple_selectbox_options_enabled"] = ["1", "1"]
    selectbox_data["multiple_selectbox_options_weight"] = ["", ""]
    selectbox_data["multiple_selectbox_options_url"] = ["", ""]
    selectbox_data["selectbox_uniqid"] = field_id
    selectbox_data["selectbox_clogic"] = json_compact({"toggle": "show", "what": "any", "rules": []})
    selectbox_data["selectbox_logicrules"] = json_compact({"toggle": "show", "rules": []})
    selectbox_data["selectbox_logic"] = ""

    return {
        "type": "selectbox",
        "div_size": "w100",
        "data": selectbox_data,
    }, {
        "field_id": field_id,
        "title": "Lunchbox Choice",
        "options": ["Non-Vegetarian", "Vegetarian"],
        "child_label": ORDINALS[child_number] + " Child Lunchbox Choice",
    }


def main():
    headers, rows = parse_csv_to_rows(SOURCE_CSV)
    clean = parse_clean_data(rows)
    sections = build_sections_from_clean(clean, headers)

    selectbox_template = None
    for section in sections:
        for element in section["elements"]:
            if element["type"] == "selectbox":
                selectbox_template = deepcopy(element)
                break
        if selectbox_template is not None:
            break

    if selectbox_template is None:
        raise ValueError("No selectbox template found in source CSV.")

    output_sections = []
    lunchbox_map = {}
    unique_counter = [0]

    for section in sections:
        updated_section = deepcopy(section)
        section_name = updated_section["data"].get("sections_internal_name", "")
        match = re.match(r"Child (\d+) Section$", section_name)

        if match:
            child_number = int(match.group(1))
            lunchbox_element, map_entry = build_lunchbox_element(
                selectbox_template,
                child_number,
                unique_counter,
            )
            updated_section["elements"].append(lunchbox_element)
            updated_section["data"]["sections"] = str(int(updated_section["data"]["sections"]) + 1)
            lunchbox_map["child_" + str(child_number)] = map_entry

        output_sections.append(updated_section)

    output_clean = build_clean_data_from_sections(output_sections, headers)
    write_clean_csv(OUTPUT_CSV, headers, output_clean)

    with OUTPUT_MAP.open("w", encoding="utf-8") as handle:
        json.dump(lunchbox_map, handle, indent=2)
        handle.write("\n")


if __name__ == "__main__":
    main()
