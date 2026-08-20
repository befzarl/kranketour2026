import json
import re
import subprocess
import tempfile
from contextlib import contextmanager
from pathlib import Path

import pandas as pd

EXCEL_FILE = "Kranke Tour Tool 2026_input.xlsx"
DRAW_SHEET = "Ziehung 2026"
LOTS_SHEET = "Lose"
PARAMETERS_SHEET = "Parameter"
ONLY_DRAWN_LOTS = False
COLORS = [
    "#1f77b4",
    "#2ca02c",
    "#9467bd",
    "#17becf",
    "#4c78a8",
    "#54a24b",
    "#b279a2",
    "#7f7f7f",
    "#bcbd22",
    "#5f9ea0",
    "#3b6fb6",
    "#2e8b57",
    "#6a5acd",
    "#00a6a6",
    "#708090",
    "#8a2be2",
    "#556b2f",
    "#4682b4",
]


def clean_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def valid_coordinate(latitude, longitude):
    return (
        pd.notna(latitude)
        and pd.notna(longitude)
        and -90 <= float(latitude) <= 90
        and -180 <= float(longitude) <= 180
    )


@contextmanager
def readable_workbook_path(excel_path):
    try:
        with excel_path.open("rb"):
            pass
        yield excel_path
        return
    except PermissionError:
        pass

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temporary_file:
        temporary_path = Path(temporary_file.name)

    try:
        subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-Command",
                "Copy-Item -LiteralPath $args[0] -Destination $args[1] -Force",
                str(excel_path),
                str(temporary_path),
            ],
            check=True,
        )
        yield temporary_path
    finally:
        temporary_path.unlink(missing_ok=True)


def lostopf_mapping(parameters):
    mapping = {}
    for _, row in parameters[["Lostopf", "Kürzel"]].dropna().iterrows():
        mapping[str(row["Kürzel"]).strip()] = str(row["Lostopf"]).strip()
    return mapping


def get_lostopf(los_id, prefix_mapping):
    los_id = str(los_id).strip()
    for prefix in sorted(prefix_mapping, key=len, reverse=True):
        if los_id.startswith(f"{prefix}-"):
            return prefix_mapping[prefix]
    return "Unbekannt"


def main():
    excel_path = Path(__file__).parent / EXCEL_FILE
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    with readable_workbook_path(excel_path) as workbook_path:
        draw = pd.read_excel(workbook_path, sheet_name=DRAW_SHEET).dropna(
            subset=["LosID"]
        )
        lots = pd.read_excel(workbook_path, sheet_name=LOTS_SHEET)
        parameters = pd.read_excel(workbook_path, sheet_name=PARAMETERS_SHEET)

    draw["LosID"] = draw["LosID"].astype(str).str.strip()
    lots = lots.dropna(subset=["LosID"]).copy()
    lots["LosID"] = lots["LosID"].astype(str).str.strip()
    lots = lots.drop_duplicates("LosID", keep="first")

    map_input = lots.rename(
        columns={
            "Bonus-punkte": "Punkte Bonus",
            "Aufgaben-zeit (min)": "Zeit-bedarf (min)",
        }
    ).copy()
    if ONLY_DRAWN_LOTS:
        drawn_los_ids = set(draw["LosID"])
        map_input = map_input[map_input["LosID"].isin(drawn_los_ids)].copy()

    records = map_input

    prefix_mapping = lostopf_mapping(parameters)
    lostopf_names = list(prefix_mapping.values())
    color_map = {
        lostopf: COLORS[index % len(COLORS)]
        for index, lostopf in enumerate(lostopf_names)
    }
    color_map["Unbekannt"] = "#7f7f7f"

    coordinates = []
    skipped = 0
    for _, row in records.iterrows():
        task_latitude = row["Breitengrad Aufgabe"]
        task_longitude = row["Längengrad Aufgabe"]
        if not valid_coordinate(task_latitude, task_longitude):
            skipped += 1
            continue
        coordinates.append(
            {
                "id": clean_value(row["LosID"]),
                "aufgabenadresse": clean_value(row["Google Adresse"]),
                "zeitbedarf": clean_value(row["Zeit-bedarf (min)"]),
                "distanz": clean_value(row["Distanz (km)"]),
                "punkte_bahnhof": clean_value(row["Punkte Bahnhof"]),
                "punkte_aufgabe": clean_value(row["Punkte Aufgabe"]),
                "punkte_bonus": clean_value(row["Punkte Bonus"]),
                "aufgabe": clean_value(row["Beschreibung"]),
                "lostopf": get_lostopf(row["LosID"], prefix_mapping),
                "latitude": round(float(task_latitude), 6),
                "longitude": round(float(task_longitude), 6),
            }
        )

    def connection_lines(lostopf):
        lines = []
        category_records = pd.DataFrame(coordinates).query("lostopf == @lostopf")
        for _, group in category_records.groupby(
            lambda index: re.sub(
                r"[a-z]$", "", str(coordinates[index]["id"]), flags=re.IGNORECASE
            )
        ):
            if len(group) > 1:
                lines.append(group[["latitude", "longitude"]].values.tolist())
        return lines

    bhf_pair_lines = connection_lines("Bhf-Paare")
    partner_city_lines = connection_lines("Partnerstädte")

    output_path = excel_path.parent / "coordinates.js"
    output_path.write_text(
        "const lostopfColors = "
        + json.dumps(color_map, ensure_ascii=False, indent=2)
        + ";\n\nconst coordinates = "
        + json.dumps(coordinates, ensure_ascii=False, indent=2)
        + ";\n\nconst bhfPaarLines = "
        + json.dumps(bhf_pair_lines, ensure_ascii=False, indent=2)
        + ";\n\nconst partnerCityLines = "
        + json.dumps(partner_city_lines, ensure_ascii=False, indent=2)
        + ";\n",
        encoding="utf-8",
    )
    print(
        f"Generated {output_path.name} with {len(coordinates)} locations; skipped {skipped} lots without coordinates."
    )


if __name__ == "__main__":
    main()
