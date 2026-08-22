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
STATIONS_SHEET = "Hafas"
PARAMETERS_SHEET = "Parameter"
ONLY_DRAWN_LOTS = True
MUTED_STATIONS = [
    "Münster(Westf)Hbf",
]
COLORS = [
    "#1f77b4",  # Bhf-Paare
    "#2ca02c",  # Ehemalige FV-Bahnhöfe
    "#551fb2",  # Essen+Trinken/Eisdiele
    "#00b4d8",  # Frankfurt & Rhein-Main-Region
    "#023e8a",  # Kleine Hochschulorte
    "#52b788",  # Kranke Ortsnamen
    "#c77dff",  # Hanse/Meer/Küste/Deich
    "#a0522d",  # Mini-Golf in Mini-Städten
    "#94e052",  # Münster Stadt & Umland/Fahrrad
    "#3a86ff",  # Eisenbahnmuseum
    "#1b4332",  # Natur/Wanderbahnhöfe
    "#6b6f2a",  # Partnerstädte
    "#5e60ce",  # Rangierbahnhöfe
    "#e040fb",  # Tatort
    "#d4a373",  # Turm/Insel/Keil/Kopf-Bahnhöfe
    "#7b2d8b",  # Unter Tage
    "#006d77",  # Weltkulturerbe
    "#e07a9e",  # Wunschlos
]


def normalize(value):
    if pd.isna(value):
        return ""
    return re.sub(r"[^a-z0-9]", "", str(value).casefold())


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
                f"Copy-Item -LiteralPath '{excel_path}' -Destination '{temporary_path}' -Force",
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
        stations = pd.read_excel(workbook_path, sheet_name=STATIONS_SHEET)
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

    stations = stations.copy()
    stations["station_key"] = stations["NAME"].map(normalize)
    stations = stations.drop_duplicates("station_key", keep="first")
    records = map_input.copy()
    records["station_key"] = records["Bahnhof"].map(normalize)
    records = records.merge(
        stations[["station_key", "Breitengrad Bahnhof", "Längengrad Bahnhof"]],
        on="station_key",
        how="left",
    )

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
        station_latitude = row["Breitengrad Bahnhof"]
        station_longitude = row["Längengrad Bahnhof"]
        task_latitude = row["Breitengrad Aufgabe"]
        task_longitude = row["Längengrad Aufgabe"]
        has_station = valid_coordinate(station_latitude, station_longitude)
        has_task = valid_coordinate(task_latitude, task_longitude)
        if not has_station and not has_task:
            skipped += 1
            continue
        primary_latitude = station_latitude if has_station else task_latitude
        primary_longitude = station_longitude if has_station else task_longitude
        coordinates.append(
            {
                "id": clean_value(row["LosID"]),
                "bahnhof": clean_value(row["Bahnhof"]),
                "aufgabenadresse": clean_value(row["Google Adresse"]),
                "zeitbedarf": clean_value(row["Zeit-bedarf (min)"]),
                "distanz": clean_value(row["Distanz (km)"]),
                "punkte_bahnhof": clean_value(row["Punkte Bahnhof"]),
                "punkte_aufgabe": clean_value(row["Punkte Aufgabe"]),
                "punkte_bonus": clean_value(row["Punkte Bonus"]),
                "aufgabe": clean_value(row["Beschreibung"]),
                "lostopf": get_lostopf(row["LosID"], prefix_mapping),
                "latitude": round(float(primary_latitude), 6),
                "longitude": round(float(primary_longitude), 6),
                "latitude2": round(float(task_latitude), 6)
                if has_station and has_task
                else None,
                "longitude2": round(float(task_longitude), 6)
                if has_station and has_task
                else None,
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
        + ";\n\nconst mutedStations = "
        + json.dumps(MUTED_STATIONS, ensure_ascii=False, indent=2)
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

    colors_entries = "\n".join(
        f"      {{ name: {json.dumps(name)}, hex: {json.dumps(hex)} }},"
        for name, hex in color_map.items()
        if name != "Unbekannt"
    )
    colors_html_path = excel_path.parent / "colors" / "colors.html"
    template_path = Path(__file__).parent / "colors" / "colors_template.html"
    colors_html = template_path.read_text(encoding="utf-8").replace(
        "__COLORS_PLACEHOLDER__", colors_entries
    )
    colors_html_path.write_text(colors_html, encoding="utf-8")
    print("Generated colors.html")


if __name__ == "__main__":
    main()
