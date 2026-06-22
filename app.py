import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

import streamlit as st


WORKBOOK_PATH = Path(__file__).with_name("Excel-tabel-for-MFT-udmærkelsestegn.-BEN3.xlsx")


@dataclass(frozen=True)
class Option:
    level: int
    requirement: str
    points: int


@dataclass(frozen=True)
class Exercise:
    key: str
    title: str
    subtitle: str
    options: tuple[Option, ...]


def _column_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref).group(0)
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter) - ord("A") + 1
    return index - 1


def _read_xlsx_rows(path: Path) -> list[list[str | None]]:
    """Read simple cell values from the bundled .xlsx without extra Excel dependencies."""
    namespace = {
        "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    with ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", namespace):
                shared_strings.append("".join(t.text or "" for t in item.findall(".//a:t", namespace)))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationship_targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships}
        first_sheet = workbook.find("a:sheets/a:sheet", namespace)
        sheet_id = first_sheet.attrib[f"{{{namespace['r']}}}id"]
        sheet_target = relationship_targets[sheet_id].lstrip("/")
        sheet_path = sheet_target if sheet_target.startswith("xl/") else f"xl/{sheet_target}"

        worksheet = ET.fromstring(archive.read(sheet_path))
        rows: list[list[str | None]] = []
        for row in worksheet.findall("a:sheetData/a:row", namespace):
            row_values: list[str | None] = []
            for cell in row.findall("a:c", namespace):
                column = _column_index(cell.attrib["r"])
                while len(row_values) <= column:
                    row_values.append(None)

                cell_type = cell.attrib.get("t")
                raw_value = cell.find("a:v", namespace)
                inline_value = cell.find("a:is", namespace)
                value: str | None = None
                if cell_type == "s" and raw_value is not None:
                    value = shared_strings[int(raw_value.text)]
                elif cell_type == "inlineStr" and inline_value is not None:
                    value = "".join(t.text or "" for t in inline_value.findall(".//a:t", namespace))
                elif raw_value is not None:
                    value = raw_value.text
                row_values[column] = value
            rows.append(row_values)
        return rows


def _get(row: list[str | None], index: int) -> str:
    if index >= len(row) or row[index] is None:
        return ""
    return str(row[index]).strip()


def _points_from_requirement(value: str) -> int:
    match = re.search(r"\((\d+)p\)", value)
    return int(match.group(1)) if match else 0


def _clean_requirement(value: str) -> str:
    return re.sub(r"\s*\(\d+p\)", "", value).strip()


@st.cache_data(show_spinner=False)
def load_exam_data(path: str) -> tuple[list[Exercise], list[dict[str, int | str]]]:
    rows = _read_xlsx_rows(Path(path))
    headers = rows[0]
    subtitles = rows[1]
    level_rows = [row for row in rows[2:] if _get(row, 0).isdigit()]

    exercise_columns = list(range(1, 9))
    exercises: list[Exercise] = []
    for column in exercise_columns:
        options: list[Option] = [Option(level=0, requirement="Ikke valgt", points=0)]
        for row in level_rows:
            requirement = _get(row, column)
            if requirement and requirement != "-":
                options.append(
                    Option(
                        level=int(_get(row, 0)),
                        requirement=_clean_requirement(requirement),
                        points=_points_from_requirement(requirement),
                    )
                )
        if options:
            title = _get(headers, column)
            title = title.replace("TEST A ", "").replace("TEST B ", "").replace("TEST C ", "")
            exercises.append(
                Exercise(
                    key=f"exercise_{column}",
                    title=title.strip(),
                    subtitle=_get(subtitles, column),
                    options=tuple(options),
                )
            )

    grade_thresholds: list[dict[str, int | str]] = []
    for row in level_rows:
        score_cell = _get(row, 9)
        grade_cell = _get(row, 10)
        if not score_cell or score_cell == "-" or not grade_cell:
            continue
        match = re.match(r"\d+", score_cell)
        if match:
            grade_thresholds.append({"minimum": int(match.group(0)), "grade": grade_cell})

    grade_thresholds.sort(key=lambda item: int(item["minimum"]))
    return exercises, grade_thresholds


def estimate_grade(total: int, thresholds: list[dict[str, int | str]]) -> str:
    grade = "Under 2"
    for threshold in thresholds:
        if total >= int(threshold["minimum"]):
            grade = str(threshold["grade"])
    return grade


def next_grade_target(total: int, thresholds: list[dict[str, int | str]]) -> str:
    current_grade = estimate_grade(total, thresholds)
    for threshold in thresholds:
        minimum = int(threshold["minimum"])
        target_grade = str(threshold["grade"])
        if total < minimum and target_grade != current_grade:
            missing = minimum - total
            return f"{missing} point til {target_grade}"
    return "Maks. nået"


def format_option(option: Option) -> str:
    if option.level == 0:
        return "Ikke valgt - 0 point"
    return f"Niveau {option.level} - {option.requirement} - {option.points} point"


def format_level(option: Option) -> str:
    return "-" if option.level == 0 else str(option.level)


def selected_test_a_mode() -> str:
    return st.segmented_control(
        "Test A type",
        options=["12 min løb", "Biptest"],
        default="12 min løb",
        help="Vælg den Test A-variant hun gennemfører.",
    )


def main() -> None:
    st.set_page_config(
        page_title="MFT Point Tracker",
        layout="centered",
        initial_sidebar_state="collapsed",
    )

    st.markdown(
        """
        <style>
        .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 760px; }
        h1 { margin-bottom: 0.15rem; }
        div[data-testid="stMetric"] {
            background: #f7f7fb;
            border: 1px solid #e6e6ef;
            border-radius: 16px;
            padding: 10px 12px;
        }
        div[data-testid="stMetricValue"] { font-size: 1.65rem; }
        .exercise-card {
            border: 1px solid #e6e6ef;
            border-radius: 16px;
            padding: 14px;
            margin: 14px 0 8px 0;
            background: #fff;
            box-shadow: 0 1px 8px rgba(20, 20, 43, 0.04);
        }
        .exercise-title { font-weight: 700; font-size: 1.05rem; margin-bottom: 0.15rem; }
        .exercise-subtitle { color: #666; font-size: 0.9rem; margin-bottom: 0.55rem; }
        .selected-result {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin-top: 8px;
        }
        .pill {
            background: #eef6ff;
            border: 1px solid #d7eaff;
            border-radius: 999px;
            padding: 5px 10px;
            font-weight: 650;
            font-size: 0.9rem;
        }
        .muted-pill {
            background: #f8f8fb;
            border: 1px solid #e6e6ef;
            border-radius: 999px;
            padding: 5px 10px;
            color: #555;
            font-size: 0.88rem;
        }
        .dashboard-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(145px, 1fr));
            gap: 8px;
            margin: 8px 0 18px 0;
        }
        .dashboard-item {
            border: 1px solid #e6e6ef;
            border-radius: 14px;
            padding: 9px 10px;
            background: white;
        }
        .dashboard-name { font-weight: 700; font-size: 0.9rem; }
        .dashboard-detail { color: #555; font-size: 0.84rem; margin-top: 2px; }
        .level-guide {
            color: #666;
            font-size: 0.82rem;
            margin-top: -4px;
            line-height: 1.35;
        }
        div[role="radiogroup"] {
            gap: 0.35rem;
            flex-wrap: wrap;
        }
        @media (max-width: 640px) {
            .block-container { padding-left: 0.85rem; padding-right: 0.85rem; }
            div[data-testid="column"] { min-width: 30% !important; }
            div[data-testid="stMetricValue"] { font-size: 1.35rem; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("MFT Point Tracker")
    st.caption("Live dashboard til eksamen. Tap niveauet for hver øvelse, så opdateres pointene med det samme.")
    summary = st.container()
    progress_placeholder = st.empty()
    dashboard_placeholder = st.empty()
    st.divider()

    if not WORKBOOK_PATH.exists():
        st.error(f"Kan ikke finde Excel-filen: {WORKBOOK_PATH.name}")
        st.stop()

    exercises, thresholds = load_exam_data(str(WORKBOOK_PATH))
    test_a_mode = selected_test_a_mode()
    active_exercises = [
        exercise
        for exercise in exercises
        if not (
            ("12 min" in exercise.title and test_a_mode == "Biptest")
            or ("Biptest" in exercise.title and test_a_mode == "12 min løb")
        )
    ]

    selections: dict[str, Option] = {}
    total = 0

    control_col, reset_col = st.columns([2, 1])
    with control_col:
        st.caption("Quick select: `-` betyder ikke valgt endnu.")
    with reset_col:
        if st.button("Nulstil", use_container_width=True):
            for exercise in active_exercises:
                st.session_state[exercise.key] = exercise.options[0]

    for exercise in active_exercises:
        st.markdown(
            f"""
            <div class="exercise-card">
                <div class="exercise-title">{exercise.title}</div>
                <div class="exercise-subtitle">{exercise.subtitle}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        level_hint = "  |  ".join(
            f"{format_level(option)}: {option.requirement} ({option.points}p)"
            for option in exercise.options
            if option.level > 0
        )
        st.markdown(f'<div class="level-guide">{html.escape(level_hint)}</div>', unsafe_allow_html=True)
        option = st.radio(
            "Vælg niveau",
            options=exercise.options,
            index=0,
            format_func=format_level,
            horizontal=True,
            key=exercise.key,
            label_visibility="collapsed",
        )
        st.markdown(
            f"""
            <div class="selected-result">
                <span class="pill">{option.points} point</span>
                <span class="muted-pill">{html.escape(option.requirement)}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        selections[exercise.key] = option
        total += option.points

    grade = estimate_grade(total, thresholds)
    next_target = next_grade_target(total, thresholds)
    selected_levels = [option.level for option in selections.values() if option.level > 0]
    lowest_level = min(selected_levels) if selected_levels else "-"
    completed = sum(1 for option in selections.values() if option.level > 0)

    with summary:
        total_col, grade_col, level_col = st.columns(3)
        total_col.metric("Total", f"{total} point")
        grade_col.metric("Karakter", grade)
        level_col.metric("Næste mål", next_target)
    progress_placeholder.progress(min(total / 350, 1.0), text=f"{total} / 350 point")

    dashboard_items = []
    for exercise in active_exercises:
        option = selections[exercise.key]
        level = format_level(option)
        dashboard_items.append(
            f"""
            <div class="dashboard-item">
                <div class="dashboard-name">{html.escape(exercise.title)}</div>
                <div class="dashboard-detail">Niveau {html.escape(level)} · {option.points}p</div>
                <div class="dashboard-detail">{html.escape(option.requirement)}</div>
            </div>
            """
        )
    dashboard_placeholder.markdown(
        f"""
        <div class="dashboard-grid">
            <div class="dashboard-item">
                <div class="dashboard-name">Status</div>
                <div class="dashboard-detail">{completed}/{len(active_exercises)} øvelser valgt</div>
                <div class="dashboard-detail">Laveste niveau: {html.escape(str(lowest_level))}</div>
            </div>
            {''.join(dashboard_items)}
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("Se pointfordeling"):
        for exercise in active_exercises:
            option = selections[exercise.key]
            st.write(f"**{exercise.title}:** {option.points} point ({option.requirement})")

    with st.expander("Karaktergrænser fra Excel-arket"):
        for threshold in thresholds:
            st.write(f"{threshold['minimum']}+ point: karakter {threshold['grade']}")


if __name__ == "__main__":
    main()
