
import streamlit as st
import pandas as pd
import numpy as np
import re
import pickle
from io import BytesIO
from datetime import datetime
from pathlib import Path

APP_VERSION = "v52 Public Safe Session Mode"

st.set_page_config(
    page_title="CAM Smart Assessment Tracker",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# PUBLIC-SAFE MODE:
# Do NOT save uploaded student Excel data to the Streamlit server.
# Each browser/user session gets its own private in-memory database.
# Data disappears when the user session ends or the app restarts.
DATA_DIR = None
DATABASE_FILE = None
LECTURER_FILE = None

MAX_FILES = 6

TS_COLS = ["TS1", "TS2", "TS3", "TS4", "TS5", "TS6"]
QUIZ_COLS = ["Q1", "Q2", "Q3"]
OBT_COLS = ["OBT1", "OBT2", "OBT3"]
MARK_COLS = TS_COLS + QUIZ_COLS + OBT_COLS

# Formula settings based on the provided Excel coursework template.
# TS: six tutorial scores out of 5 each. Total raw 30 converted to 15%.
# Quiz: best 2 quiz marks from Q1-Q3, each out of 20. Average remains out of 20%.
# OBT: best 2 OBT marks from OBT1-OBT3, each out of 15. Average remains out of 15%.
QUIZ_BEST_COUNT = 2
OBT_BEST_COUNT = 2


MAX_MARKS = {
    "TS1": 5, "TS2": 5, "TS3": 5, "TS4": 5, "TS5": 5, "TS6": 5,
    "Q1": 20, "Q2": 20, "Q3": 20,
    "OBT1": 15, "OBT2": 15, "OBT3": 15,
}

GRADE_SCALE = [
    ("A", 80, 100),
    ("A-", 75, 79.99),
    ("B+", 70, 74.99),
    ("B", 65, 69.99),
    ("B-", 60, 64.99),
    ("C+", 55, 59.99),
    ("C", 50, 54.99),
    ("D", 40, 49.99),
    ("F", 0, 39.99),
]


def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()


def normalize_header(x):
    text = clean_text(x).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return text


def get_grade(score):
    try:
        score = float(score)
    except Exception:
        return "F"

    for grade, low, high in GRADE_SCALE:
        if low <= score <= high:
            return grade
    return "F"


def calculate(df):
    df = df.copy()

    required_text_cols = [
        "File_ID", "Display_Name", "Source_File", "Source_Sheet",
        "Course_Code", "Course_Title", "Section",
        "Matric_No", "Name", "Programme", "Student_Status", "Remarks", "Uploaded_At"
    ]

    for col in required_text_cols:
        if col not in df.columns:
            df[col] = ""

    for col in MARK_COLS:
        if col not in df.columns:
            df[col] = np.nan
        df[col] = pd.to_numeric(df[col], errors="coerce")
        df[col] = df[col].round(0)

    # =========================================================
    # FORMULA UPDATED TO MATCH THE ATTACHED EXCEL TEMPLATE
    # =========================================================

    # 1. Tutorial / TS component
    # Excel formula shown: TS weighted mark = SUM(TS1:TS6) * 0.5
    # Raw TS total = 30, weighted contribution = 15%.
    df["TS_Total"] = df[TS_COLS].sum(axis=1, skipna=True)
    df["TS_15"] = df["TS_Total"] * 0.5

    # 2. Quiz component
    # Excel template uses LARGE formula with dropdown value 2.
    # Therefore Quiz_20 = average of the best 2 marks from Q1, Q2, Q3.
    def best_average(row, cols, best_count):
        values = pd.to_numeric(row[cols], errors="coerce").dropna().tolist()

        if not values:
            return np.nan

        values = sorted(values, reverse=True)
        selected_values = values[:min(best_count, len(values))]
        return round(sum(selected_values) / len(selected_values), 1)

    df["Quiz_Total"] = df[QUIZ_COLS].sum(axis=1, skipna=True)
    df["Quiz_20"] = df.apply(lambda row: best_average(row, QUIZ_COLS, QUIZ_BEST_COUNT), axis=1)

    # 3. OBT component
    # Excel template uses LARGE formula with dropdown value 2.
    # Therefore OBT_15 = average of the best 2 marks from OBT1, OBT2, OBT3.
    df["OBT_Total"] = df[OBT_COLS].sum(axis=1, skipna=True)
    df["OBT_15"] = df.apply(lambda row: best_average(row, OBT_COLS, OBT_BEST_COUNT), axis=1)

    # 4. Carry mark
    # Excel formula shown: Carry Mark 50% = TS_15 + Quiz_20 + OBT_15.
    # Carry Mark 100% = Carry Mark 50% * 2.
    df["Assessment_50"] = df[["TS_15", "Quiz_20", "OBT_15"]].sum(axis=1, skipna=True)
    df["Assessment_Percentage"] = df["Assessment_50"] * 2

    # Keep previous output naming but align it with the Excel coursework template.
    df["Carry_Mark_50"] = df["Assessment_50"]
    df["Carry_Mark_100"] = df["Assessment_Percentage"]

    # Grade is based on the 100% carry mark equivalent, same as the coursework template.
    df["Grade"] = df["Assessment_Percentage"].apply(get_grade)

    df["Completed_Components"] = df[MARK_COLS].notna().sum(axis=1)
    df["Completion_Percentage"] = (df["Completed_Components"] / len(MARK_COLS)) * 100

    # At Risk rule retained:
    # Student is At Risk if Carry Mark 50% is below 30/50.
    df["Assessment_Status"] = np.where(
        df["Assessment_50"] < 30,
        "At Risk",
        np.where(df["Completion_Percentage"] < 100, "Incomplete", "Complete")
    )

    df["At_Risk_Reason"] = np.where(
        df["Assessment_50"] < 30,
        "Carry mark below 30/50",
        ""
    )

    return df


def save_database(database):
    """
    Public-safe mode: no server-side persistence.
    The active database is kept only in st.session_state for the current user session.
    """
    return None


def load_database():
    """
    Public-safe mode: start each user/session with an empty database.
    This prevents one user's uploaded Excel file from being visible to other users.
    """
    return {}


def save_lecturer(profile):
    """
    Public-safe mode: lecturer profile remains only in st.session_state.
    """
    return None


def load_lecturer():
    """
    Public-safe mode: no shared lecturer profile is loaded from the server.
    """
    return {}


def detect_section_from_sheet(raw, fallback_filename=""):
    candidates = []

    for r in range(min(raw.shape[0], 25)):
        for c in range(min(raw.shape[1], 12)):
            text = clean_text(raw.iloc[r, c])
            if not text:
                continue

            match = re.search(r"(section|sec|group|grp)\s*[:\-]?\s*([a-zA-Z0-9]+)", text, re.IGNORECASE)
            if match:
                candidates.append(match.group(2))

            if normalize_header(text) in ["section", "sec", "group", "grp"]:
                for offset in [1, 2, 3]:
                    if c + offset < raw.shape[1]:
                        next_text = clean_text(raw.iloc[r, c + offset])
                        if next_text:
                            candidates.append(next_text)

    if candidates:
        section = str(candidates[0]).strip()
        if not section.lower().startswith("section"):
            section = f"Section {section}"
        return section

    match_file = re.search(r"\b(\d{3,4})\b", fallback_filename)
    if match_file:
        return f"Section {match_file.group(1)}"

    return "Section Unknown"


def detect_course_info(raw):
    course_title = ""
    course_code = ""

    for r in range(min(raw.shape[0], 15)):
        for c in range(min(raw.shape[1], 12)):
            text = clean_text(raw.iloc[r, c])
            key = normalize_header(text)

            if "course title" in key:
                for offset in [1, 2, 3]:
                    if c + offset < raw.shape[1]:
                        value = clean_text(raw.iloc[r, c + offset])
                        if value:
                            course_title = value
                            break

            if "course code" in key:
                for offset in [1, 2, 3]:
                    if c + offset < raw.shape[1]:
                        value = clean_text(raw.iloc[r, c + offset])
                        if value:
                            course_code = value
                            break

    return course_code, course_title


def find_student_header_row(raw):
    for r in range(raw.shape[0]):
        row_values = [normalize_header(v) for v in raw.iloc[r].tolist()]
        has_matric = any(("matric" in v or "matrix" in v or "student id" in v or "matric no" in v) for v in row_values)
        has_name = any((v == "name" or "student name" in v or "nama" in v) for v in row_values)

        if has_matric and has_name:
            return r

    return None


def map_columns(header_values):
    col_map = {}

    for idx, val in enumerate(header_values):
        h = normalize_header(val)

        if "matric" in h or "matrix" in h or "student id" in h:
            col_map["Matric_No"] = idx
        elif h == "name" or "student name" in h or "nama" in h:
            col_map["Name"] = idx
        elif h in ["prog", "programme", "program"]:
            col_map["Programme"] = idx
        elif "status" in h:
            col_map["Student_Status"] = idx

        elif h == "ts1":
            col_map["TS1"] = idx
        elif h == "ts2":
            col_map["TS2"] = idx
        elif h == "ts3":
            col_map["TS3"] = idx
        elif h == "ts4":
            col_map["TS4"] = idx
        elif h == "ts5":
            col_map["TS5"] = idx
        elif h == "ts6":
            col_map["TS6"] = idx

        elif h in ["q1", "quiz1", "quiz 1"]:
            col_map["Q1"] = idx
        elif h in ["q2", "quiz2", "quiz 2"]:
            col_map["Q2"] = idx
        elif h in ["q3", "quiz3", "quiz 3"]:
            col_map["Q3"] = idx

        elif h in ["obt1", "obt 1"]:
            col_map["OBT1"] = idx
        elif h in ["obt2", "obt 2"]:
            col_map["OBT2"] = idx
        elif h in ["obt3", "obt 3"]:
            col_map["OBT3"] = idx

    return col_map


def safe_get(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return row.iloc[idx]


def is_valid_student_record(matric, name):
    matric = clean_text(matric)
    name = clean_text(name)

    if not matric or not name:
        return False

    reject_words = [
        "matric no", "matrix no", "student id", "name",
        "total", "average", "avg", "stdv", "absent", "grade",
        "course", "lecturer", "group", "section", "session"
    ]

    if normalize_header(matric) in reject_words or normalize_header(name) in reject_words:
        return False

    if not re.search(r"\d", matric):
        return False

    if not re.search(r"[a-zA-Z]", name):
        return False

    if len(name) < 5:
        return False

    return True


def parse_excel_file(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    all_rows = []
    detection_rows = []

    for sheet_name in xl.sheet_names:
        raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        header_row = find_student_header_row(raw)

        if header_row is None:
            detection_rows.append({
                "Sheet": sheet_name,
                "Status": "Skipped",
                "Reason": "No row found with both Matric No and Name headers",
                "Detected_Students": 0
            })
            continue

        section = detect_section_from_sheet(raw, uploaded_file.name)
        course_code, course_title = detect_course_info(raw)

        header_values = raw.iloc[header_row].tolist()
        col_map = map_columns(header_values)

        if "Matric_No" not in col_map or "Name" not in col_map:
            detection_rows.append({
                "Sheet": sheet_name,
                "Status": "Skipped",
                "Reason": "Matric No or Name column not detected",
                "Detected_Students": 0
            })
            continue

        data = raw.iloc[header_row + 1:].copy()
        detected_count = 0
        skipped_incomplete = 0

        for _, row in data.iterrows():
            matric = clean_text(safe_get(row, col_map.get("Matric_No")))
            name = clean_text(safe_get(row, col_map.get("Name")))

            if not is_valid_student_record(matric, name):
                skipped_incomplete += 1
                continue

            detected_count += 1

            record = {
                "Source_File": uploaded_file.name,
                "Source_Sheet": sheet_name,
                "Course_Code": course_code,
                "Course_Title": course_title,
                "Section": section,
                "Matric_No": matric,
                "Name": name,
                "Programme": clean_text(safe_get(row, col_map.get("Programme"))),
                "Student_Status": clean_text(safe_get(row, col_map.get("Student_Status"))) or "Active",
                "Remarks": "",
                "Uploaded_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            for col in MARK_COLS:
                record[col] = safe_get(row, col_map[col]) if col in col_map else np.nan

            all_rows.append(record)

        detection_rows.append({
            "Sheet": sheet_name,
            "Status": "Read",
            "Reason": "Valid header detected",
            "Detected_Students": detected_count,
            "Skipped_Incomplete_Rows": skipped_incomplete,
            "Header_Row": header_row + 1,
            "Detected_Section": section
        })

    if not all_rows:
        raise ValueError("No valid student records found. The app only accepts rows with complete Matric No and Student Name.")

    df = pd.DataFrame(all_rows)
    df = df.drop_duplicates(subset=["Section", "Matric_No"], keep="last")
    df = calculate(df)

    return df, pd.DataFrame(detection_rows)


def create_file_id(filename, database):
    base_name = Path(filename).stem
    clean_name = re.sub(r"[^A-Za-z0-9_]+", "_", base_name).strip("_")

    if not clean_name:
        clean_name = "Excel_File"

    file_id = clean_name
    counter = 1

    while file_id in database:
        counter += 1
        file_id = f"{clean_name}_{counter}"

    return file_id


def add_file_to_database(uploaded_file, database):
    if len(database) >= MAX_FILES:
        raise ValueError(f"Maximum {MAX_FILES} Excel files can be stored. Please delete one file first.")

    df, report = parse_excel_file(uploaded_file)
    file_id = create_file_id(uploaded_file.name, database)

    display_section = df["Section"].iloc[0] if not df.empty else "Section Unknown"
    display_course = df["Course_Code"].iloc[0] if "Course_Code" in df.columns and not df.empty else ""

    display_name = f"{uploaded_file.name} | {display_section}"
    if display_course:
        display_name = f"{uploaded_file.name} | {display_course} | {display_section}"

    df["File_ID"] = file_id
    df["Display_Name"] = display_name

    database[file_id] = {
        "file_id": file_id,
        "original_filename": uploaded_file.name,
        "display_name": display_name,
        "section": display_section,
        "course_code": display_course,
        "student_count": len(df),
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "data": df,
        "detection_report": report,
    }

    return database, file_id


def get_selected_file(database):
    if not database:
        return None, None

    options = {
        f"{meta['display_name']} ({meta['student_count']} students)": file_id
        for file_id, meta in database.items()
    }

    selected_label = st.sidebar.selectbox("📁 Choose Excel File / Section", list(options.keys()))
    selected_file_id = options[selected_label]

    return selected_file_id, database[selected_file_id]


def update_selected_file(database, file_id, df):
    df = calculate(df)
    database[file_id]["data"] = df
    database[file_id]["student_count"] = len(df)
    database[file_id]["section"] = df["Section"].iloc[0] if not df.empty else database[file_id].get("section", "")
    save_database(database)
    return database


def to_excel_bytes_single(df, lecturer_profile):
    output = BytesIO()
    final_df = calculate(df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Assessment Marks")

        summary = pd.DataFrame({
            "Metric": [
                "Lecturer Name",
                "Position",
                "Total Students",
                "Average Assessment %",
                "Incomplete Students",
                "Students At Risk",
                "Generated At"
            ],
            "Value": [
                lecturer_profile.get("lecturer_name", ""),
                lecturer_profile.get("lecturer_position", ""),
                len(final_df),
                round(final_df["Assessment_Percentage"].mean(), 2),
                int((final_df["Assessment_Status"] == "Incomplete").sum()),
                int((final_df["Assessment_Status"] == "At Risk").sum()),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ]
        })
        summary.to_excel(writer, index=False, sheet_name="Summary")

    return output.getvalue()


def at_risk_warning(df):
    df = calculate(df)
    risk_df = df[df["Assessment_50"] < 30].copy()

    if risk_df.empty:
        return

    st.markdown(
        f"""
        <div class="risk-clean-card">
            <div class="risk-title">⚠️ At Risk Warning</div>
            <div class="risk-subtitle">{len(risk_df)} student(s) have carry mark below 30/50.</div>
            <div class="risk-rule">Rule: Student is marked At Risk when Assessment / Carry Mark is below 30 out of 50.</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    display_cols = ["Matric_No", "Name", "Programme", "Assessment_50", "Assessment_Percentage", "Grade"]

    with st.expander("View At Risk Student List", expanded=False):
        st.dataframe(
            risk_df[display_cols].sort_values("Assessment_50", ascending=True),
            use_container_width=True,
            hide_index=True
        )





def component_performance_stats(df):
    df = calculate(df)

    components = [
        ("Q1", 20), ("Q2", 20), ("Q3", 20),
        ("OBT1", 15), ("OBT2", 15), ("OBT3", 15)
    ]

    rows = []

    for component, max_mark in components:
        marks = pd.to_numeric(df[component], errors="coerce")
        filled_marks = marks.dropna()

        half_mark = max_mark / 2
        below_half = int((filled_marks < half_mark).sum())
        half_or_above = int((filled_marks >= half_mark).sum())
        full_mark = int((filled_marks == max_mark).sum())
        missing = int(marks.isna().sum())

        rows.append({
            "Assessment": component,
            "Max Mark": max_mark,
            "Half Mark": half_mark,
            "Filled Marks": int(filled_marks.count()),
            "Average Mark": round(filled_marks.mean(), 2) if len(filled_marks) else 0,
            "Students Requiring Support": below_half,
            "Support Percentage": round((below_half / filled_marks.count()) * 100, 1) if filled_marks.count() else 0,
            "Students Meeting Expectation": half_or_above,
            "Excellent Performance": full_mark,
            "Missing Marks": missing
        })

    return pd.DataFrame(rows)


def below_half_student_details(df):
    df = calculate(df)

    components = [
        ("Q1", 20), ("Q2", 20), ("Q3", 20),
        ("OBT1", 15), ("OBT2", 15), ("OBT3", 15)
    ]

    records = []

    for component, max_mark in components:
        half_mark = max_mark / 2
        marks = pd.to_numeric(df[component], errors="coerce")

        mask = marks.notna() & (marks < half_mark)

        selected = df.loc[mask, ["Matric_No", "Name", "Programme", "Section", component]].copy()
        selected = selected.rename(columns={component: "Mark"})
        selected["Assessment"] = component
        selected["Max Mark"] = max_mark
        selected["Half Mark"] = half_mark

        records.append(selected)

    if records:
        output = pd.concat(records, ignore_index=True)
        return output[[
            "Assessment", "Matric_No", "Name", "Programme", "Section",
            "Mark", "Max Mark", "Half Mark"
        ]]

    return pd.DataFrame(columns=[
        "Assessment", "Matric_No", "Name", "Programme", "Section",
        "Mark", "Max Mark", "Half Mark"
    ])



def render_support_student_tables(df):
    details_df = below_half_student_details(df)
    stats_df = component_performance_stats(df)

    st.markdown(
        """
        <div class="section-heading">
            <div class="section-kicker">Assessment Monitoring</div>
            <div class="section-title">Quiz and OBT Performance Support Analysis</div>
            <div class="section-subtitle">
                Students Requiring Support scored below half of the maximum mark.
                Only completed marks are analysed; blank marks are not counted.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    summary_cols = st.columns(6)

    for i, (_, row) in enumerate(stats_df.iterrows()):
        assessment = row["Assessment"]
        full_mark = int(row["Excellent Performance"])
        support_required = int(row["Students Requiring Support"])

        with summary_cols[i]:
            st.markdown(f"<div class='component-heading'>{assessment}</div>", unsafe_allow_html=True)
            st.markdown(
                f"""
                <div class="colored-stat-card full-card">
                    <div class="colored-stat-label">Full Marks Achieved</div>
                    <div class="colored-stat-value">{full_mark}</div>
                </div>
                <div class="colored-stat-card support-card">
                    <div class="colored-stat-label">Students Requiring Support</div>
                    <div class="colored-stat-value">{support_required}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown(
        """
        <div class="section-heading compact">
            <div class="section-title">Student Support Details</div>
            <div class="section-subtitle">
                The lists below show matric number and student name for students who scored below half mark.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if details_df.empty:
        st.success("No student scored below half mark for completed Quiz and OBT marks.")
        return

    quiz_details = details_df[details_df["Assessment"].isin(["Q1", "Q2", "Q3"])]
    obt_details = details_df[details_df["Assessment"].isin(["OBT1", "OBT2", "OBT3"])]

    st.markdown("<div class='subsection-title'>Quiz Components</div>", unsafe_allow_html=True)
    if quiz_details.empty:
        st.success("No Quiz marks require support.")
    else:
        for assessment in ["Q1", "Q2", "Q3"]:
            component_df = quiz_details[quiz_details["Assessment"] == assessment]
            if component_df.empty:
                continue

            with st.expander(f"{assessment}: {len(component_df)} student(s) requiring support", expanded=True):
                st.dataframe(
                    component_df[[
                        "Matric_No", "Name", "Programme", "Mark",
                        "Max Mark", "Half Mark"
                    ]],
                    use_container_width=True,
                    hide_index=True
                )

    st.markdown("<div class='subsection-title'>OBT Components</div>", unsafe_allow_html=True)
    if obt_details.empty:
        st.success("No OBT marks require support.")
    else:
        for assessment in ["OBT1", "OBT2", "OBT3"]:
            component_df = obt_details[obt_details["Assessment"] == assessment]
            if component_df.empty:
                continue

            with st.expander(f"{assessment}: {len(component_df)} student(s) requiring support", expanded=True):
                st.dataframe(
                    component_df[[
                        "Matric_No", "Name", "Programme", "Mark",
                        "Max Mark", "Half Mark"
                    ]],
                    use_container_width=True,
                    hide_index=True
                )



def clean_mark_value(value, max_value):
    """
    Clean mark values safely.
    Blank / None / NaN remains NaN.
    Numeric values are rounded to integer and clipped between 0 and max_value.
    """
    if value is None:
        return np.nan

    if isinstance(value, str):
        value = value.strip()
        if value == "" or value.lower() in ["none", "nan", "nat", "<na>"]:
            return np.nan

    try:
        numeric_value = float(value)
    except Exception:
        return np.nan

    if pd.isna(numeric_value):
        return np.nan

    numeric_value = int(round(numeric_value))
    return max(0, min(numeric_value, int(max_value)))


def display_mark_value(value):
    """
    Display missing marks as blank in Bulk Edit.
    """
    value = pd.to_numeric(value, errors="coerce")

    if pd.isna(value):
        return ""

    return str(int(round(float(value))))


def normalize_bulk_editor_df(edited_df, original_df):
    """
    Bulk edit safety layer:
    - Blank marks remain blank.
    - Matric_No and Name are protected.
    - Marks are clipped by maximum allowed mark.
    """
    cleaned = edited_df.copy()

    for protected_col in ["Matric_No", "Name"]:
        if protected_col in original_df.columns and protected_col in cleaned.columns:
            cleaned[protected_col] = original_df[protected_col].values

    for col in MARK_COLS:
        if col in cleaned.columns:
            cleaned[col] = cleaned[col].apply(lambda value: clean_mark_value(value, MAX_MARKS[col]))

    return cleaned



def integer_value(value, max_value):
    value = pd.to_numeric(value, errors="coerce")

    if pd.isna(value):
        return 0

    return int(min(max(round(float(value)), 0), max_value))


def status_color(status):
    if status == "At Risk":
        return "🔴"
    if status == "Incomplete":
        return "🟡"
    return "🟢"


# =========================
# PREMIUM UI CSS
# =========================

st.markdown("""
<style>
    :root {
        --iium-blue: #005BAC;
        --iium-navy: #003366;
        --iium-gold: #D4AF37;
        --soft-bg: #F5F7FA;
        --light-blue: #EAF4FF;
        --danger: #BA1A1A;
        --success: #0B7A3B;
        --sidebar-green: #047857;
        --sidebar-orange: #D97706;
    }

    .main {
        background: linear-gradient(180deg, #F7F9FC 0%, #FFFFFF 100%);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #002A55 0%, #003366 48%, #005BAC 100%);
        border-right: 6px solid #D4AF37;
    }

    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] span {
        color: #FFFFFF !important;
        font-size: 1.05rem !important;
    }

    [data-testid="stSidebar"] .stMarkdown {
        font-size: 1.1rem !important;
    }

    [data-testid="stSidebar"] label {
        color: #FFD966 !important;
        font-weight: 900 !important;
    }

    [data-testid="stSidebar"] div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #002B5B !important;
        border: 3px solid #D4AF37 !important;
        border-radius: 14px !important;
        font-weight: 900 !important;
        min-height: 52px !important;
        box-shadow: 0 6px 18px rgba(0,0,0,0.18) !important;
    }

    [data-testid="stSidebar"] div[data-baseweb="select"] span {
        color: #002B5B !important;
        font-weight: 900 !important;
        font-size: 0.95rem !important;
    }

    div[data-baseweb="popover"] {
        z-index: 999999 !important;
    }

    div[data-baseweb="popover"] * {
        color: #002B5B !important;
        background-color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    ul[role="listbox"] li {
        color: #002B5B !important;
        background-color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    h1, h2, h3 {
        color: #003366;
        font-weight: 900;
    }

    .stButton>button {
        border-radius: 14px;
        background-color: #005BAC;
        color: white;
        border: none;
        font-weight: 900;
        padding: 0.7rem 1.2rem;
        box-shadow: 0 4px 12px rgba(0, 91, 172, 0.18);
    }

    .stButton>button:hover {
        background-color: #003366;
        color: white;
    }

    .stDownloadButton>button {
        border-radius: 14px;
        background-color: #D4AF37;
        color: #003366;
        border: none;
        font-weight: 900;
    }

    .hero-card {
        padding: 1.6rem;
        border-radius: 24px;
        background: linear-gradient(135deg, #003366 0%, #005BAC 72%, #2C7FD3 100%);
        color: white;
        box-shadow: 0 14px 32px rgba(0, 51, 102, 0.22);
        margin-bottom: 1rem;
    }

    .hero-card h1 {
        color: white;
        margin-bottom: 0.2rem;
        font-size: 2.2rem;
        line-height: 1.15;
    }

    .hero-card p {
        color: #EAF4FF;
        margin-bottom: 0;
        font-size: 1rem;
    }

    .lecturer-panel {
        margin-top: 1rem;
        padding: 1rem;
        border-radius: 18px;
        background: rgba(255,255,255,0.16);
        border: 1px solid rgba(255,255,255,0.35);
    }

    .lecturer-panel .label {
        color: #D4AF37;
        font-size: 1rem;
        font-weight: 900;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }

    .lecturer-panel .value {
        color: white;
        font-size: 1.55rem;
        font-weight: 900;
        line-height: 1.3;
    }

    .info-box {
        padding: 1rem;
        border-radius: 18px;
        background: #EAF4FF;
        border: 1px solid #D7E9FF;
        margin-bottom: 1rem;
        color: #003366;
    }

    .side-card {
        margin-top: 1rem;
        padding: 1rem;
        border-radius: 18px;
        background: rgba(255,255,255,0.13);
        border: 1px solid rgba(255,255,255,0.25);
        box-shadow: 0 8px 18px rgba(0,0,0,0.15);
    }

    .side-card.lecturer {
        background: linear-gradient(135deg, rgba(30,58,138,0.92), rgba(37,99,235,0.70));
    }

    .side-card.file {
        background: linear-gradient(135deg, rgba(4,120,87,0.92), rgba(16,185,129,0.72));
    }

    .side-card.alert {
        background: linear-gradient(135deg, rgba(146,64,14,0.94), rgba(245,158,11,0.72));
    }

    .side-card-title {
        color: #FFE082 !important;
        font-weight: 900 !important;
        font-size: 1rem !important;
        margin-bottom: 0.45rem;
    }

    .side-card-value {
        color: white !important;
        font-weight: 900 !important;
        font-size: 1.18rem !important;
        line-height: 1.35;
    }

    .side-card-text {
        color: #FFFFFF !important;
        font-weight: 800 !important;
        font-size: 0.95rem !important;
    }

    .dashboard-card {
        padding: 1.2rem;
        border-radius: 22px;
        background: white;
        border: 1px solid #E5E7EB;
        box-shadow: 0 8px 22px rgba(0,0,0,0.06);
    }

    .dashboard-card .card-label {
        color: #64748B;
        font-size: 0.86rem;
        font-weight: 800;
        text-transform: uppercase;
    }

    .dashboard-card .card-value {
        color: #003366;
        font-size: 2rem;
        font-weight: 900;
    }

    .dashboard-card.gold {
        background: linear-gradient(135deg, #FFF8DC, #FFFFFF);
        border: 1px solid #D4AF37;
    }

    .dashboard-card.blue {
        background: linear-gradient(135deg, #EAF4FF, #FFFFFF);
        border: 1px solid #B7D9FF;
    }

    .dashboard-card.red {
        background: linear-gradient(135deg, #FFEAEA, #FFFFFF);
        border: 1px solid #FFB4AB;
    }


    .warning-panel {
        background: linear-gradient(135deg, #7F1D1D, #DC2626);
        color: white;
        padding: 1rem;
        border-radius: 18px;
        box-shadow: 0 10px 24px rgba(220,38,38,0.22);
        margin: 1rem 0;
    }

    .warning-panel h3 {
        color: white !important;
        margin-bottom: 0.3rem;
    }

    .mini-card {
        padding: 1rem;
        border-radius: 18px;
        background: white;
        border: 1px solid #E5E7EB;
        box-shadow: 0 6px 18px rgba(0,0,0,0.05);
        text-align: center;
    }

    .mini-card .mini-label {
        color: #64748B;
        font-weight: 800;
        font-size: 0.85rem;
    }

    .mini-card .mini-value {
        color: #003366;
        font-weight: 900;
        font-size: 1.7rem;
    }


    .half-card {
        padding: 1.15rem;
        border-radius: 22px;
        color: #003366;
        border: 2px solid rgba(212,175,55,0.55);
        box-shadow: 0 8px 22px rgba(0,0,0,0.07);
        margin-bottom: 0.8rem;
        min-height: 170px;
    }

    .half-card.quiz {
        background: linear-gradient(135deg, #EAF4FF 0%, #FFFFFF 100%);
    }

    .half-card.obt {
        background: linear-gradient(135deg, #FFF8DC 0%, #FFFFFF 100%);
    }

    .half-title {
        font-size: 1.25rem;
        font-weight: 900;
        color: #003366;
    }

    .half-value {
        font-size: 3rem;
        font-weight: 950;
        color: #005BAC;
        line-height: 1.05;
        margin-top: 0.35rem;
    }

    .half-sub {
        font-size: 0.95rem;
        color: #334155;
        font-weight: 800;
        margin-bottom: 0.55rem;
    }

    .half-detail {
        font-size: 0.9rem;
        color: #003366;
        font-weight: 800;
        padding: 0.12rem 0;
    }

    .delete-zone {
        background: linear-gradient(135deg, #FFEAEA, #FFFFFF);
        border: 2px solid #FFB4AB;
        border-radius: 22px;
        padding: 1.1rem;
        margin-top: 1rem;
    }


    .risk-clean-card {
        background: linear-gradient(135deg, #FFF1F2 0%, #FFFFFF 100%);
        border: 2px solid #FDA4AF;
        border-left: 8px solid #DC2626;
        border-radius: 18px;
        padding: 1rem 1.2rem;
        margin: 1rem 0;
        box-shadow: 0 8px 22px rgba(220, 38, 38, 0.10);
    }

    .risk-title {
        color: #991B1B;
        font-size: 1.2rem;
        font-weight: 950;
        margin-bottom: 0.25rem;
    }

    .risk-subtitle {
        color: #7F1D1D;
        font-size: 1rem;
        font-weight: 850;
    }

    .risk-rule {
        color: #9F1239;
        font-size: 0.9rem;
        font-weight: 700;
        margin-top: 0.25rem;
    }


    .below-half-card {
        padding: 1.2rem;
        border-radius: 22px;
        border: 2px solid rgba(220,38,38,0.28);
        box-shadow: 0 8px 22px rgba(0,0,0,0.07);
        margin-bottom: 0.85rem;
        min-height: 185px;
    }

    .below-half-card.quiz {
        background: linear-gradient(135deg, #FFEAEA 0%, #FFFFFF 100%);
    }

    .below-half-card.obt {
        background: linear-gradient(135deg, #FFF8DC 0%, #FFFFFF 100%);
    }

    .below-half-title {
        font-size: 1.25rem;
        font-weight: 950;
        color: #003366;
    }

    .below-half-number {
        font-size: 3rem;
        font-weight: 950;
        color: #DC2626;
        line-height: 1.05;
        margin-top: 0.35rem;
    }

    .below-half-text {
        font-size: 0.95rem;
        color: #7F1D1D;
        font-weight: 850;
        margin-bottom: 0.55rem;
    }

    .below-half-detail {
        font-size: 0.9rem;
        color: #003366;
        font-weight: 800;
        padding: 0.12rem 0;
    }


    .compact-performance-card {
        padding: 0.85rem 1rem;
        border-radius: 18px;
        border: 1.5px solid rgba(220,38,38,0.22);
        box-shadow: 0 5px 16px rgba(0,0,0,0.055);
        margin-bottom: 0.75rem;
        min-height: 138px;
    }

    .compact-performance-card.quiz {
        background: linear-gradient(135deg, #FFF1F2 0%, #FFFFFF 100%);
    }

    .compact-performance-card.obt {
        background: linear-gradient(135deg, #FFF8DC 0%, #FFFFFF 100%);
    }

    .compact-title {
        font-size: 1.05rem;
        font-weight: 950;
        color: #003366;
        margin-bottom: 0.15rem;
    }

    .compact-number {
        font-size: 2.15rem;
        font-weight: 950;
        color: #DC2626;
        line-height: 1;
    }

    .compact-text {
        font-size: 0.8rem;
        color: #7F1D1D;
        font-weight: 850;
        margin-bottom: 0.45rem;
    }

    .compact-grid {
        display: grid;
        grid-template-columns: 1fr auto;
        gap: 0.16rem 0.45rem;
        font-size: 0.78rem;
        color: #003366;
    }

    .compact-grid span {
        font-weight: 750;
        color: #475569;
    }

    .compact-grid b {
        font-weight: 950;
        color: #003366;
    }


    /* ===== BIG MOBILE-FRIENDLY TABS ===== */
    button[data-baseweb="tab"] {
        font-size: 22px !important;
        font-weight: 900 !important;
        min-height: 68px !important;
        padding: 16px 28px !important;
        border-radius: 16px !important;
        background: #F8FAFC !important;
        margin-right: 10px !important;
        border: 2px solid #E2E8F0 !important;
        color: #003366 !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.05) !important;
    }

    button[data-baseweb="tab"] p {
        font-size: 22px !important;
        font-weight: 900 !important;
        color: #003366 !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #003366 0%, #005BAC 100%) !important;
        color: #FFFFFF !important;
        border: 3px solid #D4AF37 !important;
        box-shadow: 0 8px 18px rgba(0, 51, 102, 0.22) !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] p {
        color: #FFFFFF !important;
    }

    button[data-baseweb="tab"]:hover {
        background: #EAF4FF !important;
        transform: translateY(-2px);
        transition: 0.2s ease-in-out;
    }

    div[data-testid="stTabs"] {
        margin-top: 18px !important;
        margin-bottom: 18px !important;
    }

    div[data-testid="stTabs"] [role="tablist"] {
        gap: 8px !important;
        flex-wrap: wrap !important;
    }


    .delete-file-panel {
        background: linear-gradient(135deg, #FFF1F2 0%, #FFFFFF 100%);
        border: 2px solid #FDA4AF;
        border-left: 8px solid #DC2626;
        border-radius: 20px;
        padding: 1.1rem 1.2rem;
        margin: 1rem 0;
        box-shadow: 0 8px 22px rgba(220, 38, 38, 0.10);
    }

    .delete-file-title {
        color: #991B1B;
        font-size: 1.2rem;
        font-weight: 950;
        margin-bottom: 0.3rem;
    }

    .delete-file-text {
        color: #7F1D1D;
        font-size: 0.95rem;
        font-weight: 750;
    }

    .file-info-panel {
        background: linear-gradient(135deg, #EAF4FF 0%, #FFFFFF 100%);
        border: 2px solid #B7D9FF;
        border-left: 8px solid #005BAC;
        border-radius: 20px;
        padding: 1rem 1.2rem;
        margin: 1rem 0;
        box-shadow: 0 8px 22px rgba(0, 91, 172, 0.08);
    }

    .file-info-title {
        color: #003366;
        font-size: 1.08rem;
        font-weight: 950;
    }

    .file-info-text {
        color: #003366;
        font-size: 0.95rem;
        font-weight: 750;
    }


    .support-summary-card {
        padding: 0.85rem;
        border-radius: 16px;
        border: 1.5px solid rgba(220,38,38,0.24);
        background: linear-gradient(135deg, #FFF1F2 0%, #FFFFFF 100%);
        box-shadow: 0 5px 15px rgba(0,0,0,0.05);
        margin-bottom: 0.75rem;
        min-height: 230px;
    }

    .support-title {
        font-size: 1.05rem;
        font-weight: 950;
        color: #003366;
        margin-bottom: 0.35rem;
    }

    .support-label-red,
    .support-label-blue,
    .support-label-green {
        font-size: 0.72rem;
        font-weight: 900;
        text-transform: uppercase;
        letter-spacing: 0.02em;
        margin-top: 0.3rem;
    }

    .support-label-red {
        color: #991B1B;
    }

    .support-label-blue {
        color: #1D4ED8;
    }

    .support-label-green {
        color: #047857;
    }

    .support-number-red,
    .support-number-blue,
    .support-number-green {
        font-size: 1.75rem;
        font-weight: 950;
        line-height: 1;
        margin-bottom: 0.2rem;
    }

    .support-number-red {
        color: #DC2626;
    }

    .support-number-blue {
        color: #2563EB;
    }

    .support-number-green {
        color: #059669;
    }

    .support-detail {
        font-size: 0.78rem;
        color: #003366;
        font-weight: 800;
        padding: 0.08rem 0;
        margin-top: 0.18rem;
    }


    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFC 100%);
        border: 1.5px solid #D7E9FF;
        border-radius: 16px;
        padding: 0.85rem;
        box-shadow: 0 5px 14px rgba(0,0,0,0.05);
        margin-bottom: 0.55rem;
    }

    [data-testid="stMetricLabel"] {
        font-weight: 900 !important;
        color: #003366 !important;
    }

    [data-testid="stMetricValue"] {
        font-weight: 950 !important;
        color: #005BAC !important;
    }


    .colored-stat-card {
        border-radius: 18px;
        padding: 1rem 1.05rem;
        margin-bottom: 0.85rem;
        min-height: 112px;
        box-shadow: 0 8px 22px rgba(0,0,0,0.12);
        border: 1.5px solid rgba(255,255,255,0.35);
    }

    .colored-stat-card.full-card {
        background: linear-gradient(135deg, #047857 0%, #10B981 100%);
    }

    .colored-stat-card.support-card {
        background: linear-gradient(135deg, #B91C1C 0%, #EF4444 100%);
    }

    .colored-stat-label {
        color: #FFFFFF;
        font-size: 0.9rem;
        font-weight: 900;
        letter-spacing: 0.02em;
        margin-bottom: 0.4rem;
    }

    .colored-stat-value {
        color: #FFFFFF;
        font-size: 2.4rem;
        font-weight: 950;
        line-height: 1;
    }

    .colored-stat-card.full-card .colored-stat-value {
        text-shadow: 0 2px 8px rgba(0,0,0,0.18);
    }

    .colored-stat-card.support-card .colored-stat-value {
        text-shadow: 0 2px 8px rgba(0,0,0,0.18);
    }


    .section-heading {
        margin: 1.2rem 0 1rem 0;
        padding: 1.1rem 1.2rem;
        border-radius: 20px;
        background: linear-gradient(135deg, #EAF4FF 0%, #FFFFFF 100%);
        border-left: 8px solid #005BAC;
        box-shadow: 0 8px 22px rgba(0, 91, 172, 0.08);
    }

    .section-heading.compact {
        margin-top: 1.6rem;
        padding: 0.95rem 1.1rem;
    }

    .section-kicker {
        color: #D4AF37;
        font-size: 0.85rem;
        font-weight: 950;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 0.2rem;
    }

    .section-title {
        color: #003366;
        font-size: 1.55rem;
        font-weight: 950;
        line-height: 1.25;
    }

    .section-subtitle {
        color: #475569;
        font-size: 0.95rem;
        font-weight: 650;
        margin-top: 0.25rem;
    }

    .subsection-title {
        color: #003366;
        font-size: 1.25rem;
        font-weight: 950;
        margin: 1.2rem 0 0.55rem 0;
        padding-bottom: 0.35rem;
        border-bottom: 3px solid #D4AF37;
        width: fit-content;
    }

    .component-heading {
        color: #003366;
        font-size: 1.35rem;
        font-weight: 950;
        margin: 0.6rem 0 0.5rem 0;
        letter-spacing: -0.02em;
    }

    .colored-stat-label {
        font-size: 0.82rem !important;
        line-height: 1.2;
    }
@media (max-width: 768px) {
        .hero-card h1 {
            font-size: 1.55rem;
        }

        .lecturer-panel .value {
            font-size: 1.25rem;
        }

        button[data-baseweb="tab"] {
            font-size: 18px !important;
            min-height: 58px !important;
            padding: 12px 18px !important;
        }

        button[data-baseweb="tab"] p {
            font-size: 18px !important;
        }

    }
</style>
""", unsafe_allow_html=True)


# =========================
# SESSION
# =========================

if "database" not in st.session_state:
    st.session_state.database = load_database()

if "lecturer_profile" not in st.session_state:
    st.session_state.lecturer_profile = load_lecturer()


# =========================
# SIDEBAR
# =========================

with st.sidebar:
    st.markdown("## 📊 CAM Manager")
    st.caption("Continuous Assessment Mobile Entry")

    profile = st.session_state.lecturer_profile
    lecturer_name_side = profile.get("lecturer_name", "Lecturer name not set")
    lecturer_position_side = profile.get("lecturer_position", "Position not set")

    st.markdown(
        f"""
        <div class="side-card lecturer">
            <div class="side-card-title">👩‍🏫 Lecturer Profile</div>
            <div class="side-card-value">{lecturer_name_side}</div>
            <div class="side-card-text">{lecturer_position_side}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    selected_file_id = None
    selected_file = None

    if st.session_state.database:
        selected_file_id, selected_file = get_selected_file(st.session_state.database)
        st.session_state.selected_file_id = selected_file_id

        selected_df = calculate(selected_file["data"])
        avg = selected_df["Assessment_Percentage"].mean() if len(selected_df) else 0
        at_risk_count = int((selected_df["Assessment_Status"] == "At Risk").sum())
        incomplete_count = int((selected_df["Assessment_Status"] == "Incomplete").sum())

        st.markdown(
            f"""
            <div class="side-card file">
                <div class="side-card-title">📚 Current File</div>
                <div class="side-card-value">{selected_file.get('section', '')}</div>
                <div class="side-card-text">{selected_file.get('course_code', '')}</div>
                <div class="side-card-text">👥 {len(selected_df)} valid students</div>
                <div class="side-card-text">📈 {avg:.1f}% average</div>
            </div>
            <div class="side-card alert">
                <div class="side-card-title">🔔 Alerts</div>
                <div class="side-card-text">🔴 {at_risk_count} carry mark below 30/50</div>
                <div class="side-card-text">🟡 {incomplete_count} incomplete</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.info("Upload Excel files first.")

    st.divider()

    page = st.radio(
        "Navigation",
        [
            "🏠 Dashboard",
            "👩‍🏫 Lecturer Profile",
            "📤 Upload Excel Files",
            "✍️ Select & Key In Marks",
            "⬇️ Download / Export",
            "🗂️ File Manager"
        ],
        label_visibility="collapsed"
    )


# =========================
# HEADER
# =========================

profile = st.session_state.lecturer_profile
lecturer_name = profile.get("lecturer_name", "")
lecturer_position = profile.get("lecturer_position", "")

st.markdown(
    f"""
    <div class="hero-card">
        <h1>📊 CAM Smart Assessment Tracker</h1>
        <p>Smart continuous assessment tracking for faster mark entry, early support detection, and academic performance monitoring.</p>
        <div class="lecturer-panel">
            <div class="label">Lecturer Information</div>
            <div class="value">{lecturer_name or 'Lecturer name not set'}</div>
            <div class="value">{lecturer_position or 'Position not set'}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


st.info("🔒 Public-safe mode: uploaded Excel files are private to your current browser session and are not saved to the shared server.")


# =========================
# PAGE: LECTURER PROFILE
# =========================

if page == "👩‍🏫 Lecturer Profile":
    st.subheader("👩‍🏫 Lecturer Profile")

    lecturer_name = st.text_input(
        "Lecturer Name",
        value=profile.get("lecturer_name", ""),
        placeholder="Example: SUHAILA BINTI BAHROM"
    )

    position_options = ["Matriculation Lecturer", "Matriculation Teacher"]
    saved_position = profile.get("lecturer_position", "Matriculation Lecturer")
    saved_index = position_options.index(saved_position) if saved_position in position_options else 0

    lecturer_position = st.selectbox(
        "Position / Jawatan",
        position_options,
        index=saved_index
    )

    institution = st.text_input(
        "Institution",
        value=profile.get("institution", "International Islamic University Malaysia"),
        placeholder="International Islamic University Malaysia"
    )

    if st.button("💾 Save Lecturer Profile"):
        st.session_state.lecturer_profile = {
            "lecturer_name": lecturer_name,
            "lecturer_position": lecturer_position,
            "institution": institution,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        save_lecturer(st.session_state.lecturer_profile)
        st.success("Lecturer profile saved successfully.")
        st.rerun()


# =========================
# PAGE: UPLOAD FILES
# =========================

elif page == "📤 Upload Excel Files":
    st.subheader("📤 Upload Excel Files")
    st.markdown(
        """
        <div class="info-box">
        <b>Rules:</b><br>
        1. Store up to <b>6 Excel files per browser session</b>.<br>
        2. Files are kept privately in your current session only.<br>
        3. Other users cannot see your uploaded Excel file.<br>
        4. Data disappears when your session ends or the app restarts.<br>
        5. Only rows with complete <b>Matric No + Student Name</b> are counted.
        </div>
        """,
        unsafe_allow_html=True
    )

    current_count = len(st.session_state.database)
    remaining_slots = MAX_FILES - current_count

    c1, c2, c3 = st.columns(3)
    c1.metric("Stored Excel Files", current_count)
    c2.metric("Remaining Slots", remaining_slots)
    c3.metric("Maximum Files", MAX_FILES)

    if remaining_slots <= 0:
        st.error("Storage is full. Please delete one file in File Manager before uploading another file.")
        st.stop()

    uploaded_files = st.file_uploader(
        f"Upload Excel files. You can still add {remaining_slots} more file(s).",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        if len(uploaded_files) > remaining_slots:
            st.error(f"You selected {len(uploaded_files)} files, but only {remaining_slots} slot(s) are available.")
            st.stop()

        st.write("Selected file(s):")
        for f in uploaded_files:
            st.write(f"- {f.name}")

        if st.button("➕ Save Excel File(s)"):
            success_count = 0

            for f in uploaded_files:
                try:
                    st.session_state.database, file_id = add_file_to_database(f, st.session_state.database)
                    success_count += 1
                    st.success(f"Saved: {f.name}")
                except Exception as e:
                    st.error(f"{f.name}: {e}")

            save_database(st.session_state.database)
            st.success(f"{success_count} file(s) saved successfully.")
            st.rerun()

    st.divider()
    st.subheader("Currently Stored Files")

    if st.session_state.database:
        file_table = []

        for file_id, meta in st.session_state.database.items():
            file_table.append({
                "File ID": file_id,
                "Original Filename": meta["original_filename"],
                "Section": meta["section"],
                "Course Code": meta.get("course_code", ""),
                "Valid Students": meta["student_count"],
                "Uploaded At": meta["uploaded_at"],
            })

        st.dataframe(pd.DataFrame(file_table), use_container_width=True, hide_index=True)
    else:
        st.info("No Excel file stored yet.")


# =========================
# PAGE: KEY IN MARKS
# =========================

elif page == "✍️ Select & Key In Marks":
    st.subheader("✍️ Select File and Key In Marks")

    if not st.session_state.database:
        st.warning("Please upload at least one Excel file first.")
        st.stop()

    file_id = st.session_state.selected_file_id
    meta = st.session_state.database[file_id]
    df = calculate(meta["data"])

    at_risk_warning(df)

    st.markdown(f"### {meta['display_name']}")
    st.caption(f"Valid students detected: {len(df)} | Section: {meta['section']}")

    search = st.text_input("Search student by name or matric number")

    if search:
        df_view = df[
            df["Name"].str.contains(search, case=False, na=False) |
            df["Matric_No"].astype(str).str.contains(search, case=False, na=False)
        ].copy()
    else:
        df_view = df.copy()

    if df_view.empty:
        st.warning("No matching student found.")
        st.stop()

    student_options = (df_view["Matric_No"].astype(str) + " - " + df_view["Name"]).tolist()
    selected_student = st.selectbox("Select Student", student_options)
    selected_matric = selected_student.split(" - ")[0]

    idx = df[df["Matric_No"].astype(str) == selected_matric].index[0]
    student = df.loc[idx]

    st.info(
        f"Student: {student['Name']} | Matric No: {student['Matric_No']} | "
        f"Programme: {student['Programme']} | Status: {student['Student_Status']}"
    )

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Tutorial", "📊 Quiz", "🎯 OBT", "📈 Summary", "⚡ Bulk Edit"])

    with tab1:
        st.markdown("### Tutorial / TS Marks")
        st.caption("Integer marks only. Maximum mark: 5 each")
        cols = st.columns(6)

        for i, col in enumerate(TS_COLS):
            value = integer_value(df.loc[idx, col], 5)
            df.loc[idx, col] = cols[i].number_input(
                col,
                min_value=0,
                max_value=5,
                value=value,
                step=1,
                key=f"{file_id}_{idx}_{col}"
            )

    with tab2:
        st.markdown("### Quiz Marks")
        st.caption("Integer marks only. Maximum mark: 20 each")
        cols = st.columns(3)

        for i, col in enumerate(QUIZ_COLS):
            value = integer_value(df.loc[idx, col], 20)
            df.loc[idx, col] = cols[i].number_input(
                col,
                min_value=0,
                max_value=20,
                value=value,
                step=1,
                key=f"{file_id}_{idx}_{col}"
            )

    with tab3:
        st.markdown("### OBT Marks")
        st.caption("Integer marks only. Maximum mark: 15 each")
        cols = st.columns(3)

        for i, col in enumerate(OBT_COLS):
            value = integer_value(df.loc[idx, col], 15)
            df.loc[idx, col] = cols[i].number_input(
                col,
                min_value=0,
                max_value=15,
                value=value,
                step=1,
                key=f"{file_id}_{idx}_{col}"
            )

    df = calculate(df)

    with tab4:
        updated = df.loc[idx]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("TS 15%", f"{updated['TS_15']:.2f}")
        c2.metric("Quiz 20%", f"{updated['Quiz_20']:.2f}")
        c3.metric("OBT 15%", f"{updated['OBT_15']:.2f}")
        c4.metric("Assessment 50%", f"{updated['Assessment_50']:.2f}")

        c5, c6, c7 = st.columns(3)
        c5.metric("Assessment %", f"{updated['Assessment_Percentage']:.1f}%")
        c6.metric("Grade", updated["Grade"])
        c7.metric("Completion", f"{updated['Completion_Percentage']:.0f}%")

        df.loc[idx, "Remarks"] = st.text_area(
            "Remarks",
            value=str(df.loc[idx, "Remarks"]) if pd.notna(df.loc[idx, "Remarks"]) else "",
            key=f"{file_id}_remarks_{idx}"
        )

    with tab5:
        st.markdown("### Bulk Edit This Excel File")
        st.caption(
            "Paste marks from Excel here. Matric No and Name are locked. Blank cells stay blank. Marks are auto-updated after paste."
        )

        editable_cols = ["Matric_No", "Name", "Programme", "Student_Status"] + MARK_COLS + ["Remarks"]
        bulk_df = df[editable_cols].copy()

        # Display missing marks as blank, not 0 and not None.
        for col in MARK_COLS:
            bulk_df[col] = bulk_df[col].apply(display_mark_value)

        edited = st.data_editor(
            bulk_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key=f"bulk_editor_v30_{file_id}",
            column_config={
                "Matric_No": st.column_config.TextColumn("Matric No", disabled=True),
                "Name": st.column_config.TextColumn("Name", disabled=True),
                "Programme": st.column_config.TextColumn("Programme"),
                "Student_Status": st.column_config.TextColumn("Student Status"),
                **{
                    col: st.column_config.TextColumn(
                        col,
                        help=f"Integer only. Maximum allowed: {MAX_MARKS[col]}. Blank is allowed."
                    )
                    for col in MARK_COLS
                },
                "Remarks": st.column_config.TextColumn("Remarks"),
            }
        )

        # Clean and apply immediately, so calculation uses the latest pasted values.
        cleaned = normalize_bulk_editor_df(edited, df)

        for col in editable_cols:
            if col in df.columns and col in cleaned.columns:
                df[col] = cleaned[col].values

        df = calculate(df)
        st.session_state[f"bulk_cleaned_{file_id}"] = df.copy()

        preview_df = df[editable_cols].copy()
        for col in MARK_COLS:
            preview_df[col] = preview_df[col].apply(display_mark_value)

        with st.expander("Preview updated marks before saving", expanded=False):
            st.dataframe(
                preview_df,
                use_container_width=True,
                hide_index=True
            )

        st.success("Bulk Edit updated in memory. Click 'Save Marks for This Excel File' to store it permanently.")

    if st.button("💾 Save Marks for This Excel File"):
        bulk_key = f"bulk_cleaned_{file_id}"

        if bulk_key in st.session_state:
            df = st.session_state[bulk_key].copy()

        df = calculate(df)
        st.session_state.database = update_selected_file(st.session_state.database, file_id, df)
        save_database(st.session_state.database)

        if bulk_key in st.session_state:
            del st.session_state[bulk_key]

        st.success("Marks saved successfully for selected Excel file.")
        at_risk_warning(df)


# =========================
# PAGE: DASHBOARD
# =========================

elif page == "🏠 Dashboard":
    st.subheader("📈 Dashboard")

    if not st.session_state.database:
        st.warning("Please upload Excel files first.")
        st.stop()

    file_id = st.session_state.selected_file_id
    meta = st.session_state.database[file_id]
    df = calculate(meta["data"])

    at_risk_warning(df)

    st.markdown(f"### {meta['display_name']}")

    total_students = len(df)
    avg_assessment = df["Assessment_Percentage"].mean() if total_students else 0
    incomplete_count = int((df["Assessment_Status"] == "Incomplete").sum())
    at_risk_count = int((df["Assessment_Status"] == "At Risk").sum())
    complete_count = int((df["Assessment_Status"] == "Complete").sum())

    st.markdown(
        f"""
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:20px;">
            <div class="dashboard-card blue">
                <div class="card-label">Valid Students</div>
                <div class="card-value">{total_students}</div>
            </div>
            <div class="dashboard-card gold">
                <div class="card-label">Average Assessment</div>
                <div class="card-value">{avg_assessment:.1f}%</div>
            </div>
            <div class="dashboard-card">
                <div class="card-label">Completed</div>
                <div class="card-value">{complete_count}</div>
            </div>
            <div class="dashboard-card red">
                <div class="card-label">At Risk</div>
                <div class="card-value">{at_risk_count}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Component chart and status summary removed in v18 because the left panel already shows At Risk and Incomplete counts.

    render_support_student_tables(df)

    st.subheader("At Risk Students: Carry Mark Below 30/50")
    risk_df = df[df["Assessment_50"] < 30]

    if risk_df.empty:
        st.success("No students with carry mark below 30/50 for the selected file.")
    else:
        st.error(f"{len(risk_df)} student(s) have carry mark below 30/50.")
        st.dataframe(
            risk_df[["Matric_No", "Name", "Programme", "Assessment_50", "Assessment_Percentage", "Grade", "At_Risk_Reason"]],
            use_container_width=True,
            hide_index=True
        )

    st.subheader("Student Summary")
    summary_df = df[["Matric_No", "Name", "Programme", "TS_15", "Quiz_20", "OBT_15", "Assessment_50", "Assessment_Percentage", "Grade", "Assessment_Status"]].copy()
    summary_df["Status"] = summary_df["Assessment_Status"].apply(status_color)

    st.dataframe(
        summary_df[["Status", "Matric_No", "Name", "Programme", "TS_15", "Quiz_20", "OBT_15", "Assessment_50", "Assessment_Percentage", "Grade", "Assessment_Status"]],
        use_container_width=True,
        hide_index=True
    )


# =========================
# PAGE: DOWNLOAD
# =========================

elif page == "⬇️ Download / Export":
    st.subheader("⬇️ Download / Export")

    if not st.session_state.database:
        st.warning("Please upload Excel files first.")
        st.stop()

    file_id = st.session_state.selected_file_id
    meta = st.session_state.database[file_id]
    df = calculate(meta["data"])

    st.markdown(f"### Selected File: {meta['display_name']}")
    st.dataframe(df, use_container_width=True, hide_index=True)

    profile = st.session_state.lecturer_profile

    st.download_button(
        "⬇️ Download Selected Excel File as CSV",
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{Path(meta['original_filename']).stem}_updated.csv",
        mime="text/csv"
    )

    st.download_button(
        "⬇️ Download Selected Excel File as Excel",
        data=to_excel_bytes_single(df, profile),
        file_name=f"{Path(meta['original_filename']).stem}_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# PAGE: FILE MANAGER
# =========================

elif page == "🗂️ File Manager":
    st.subheader("🗂️ File Manager")

    if not st.session_state.database:
        st.info("No Excel file stored yet.")
        st.stop()

    file_table = []

    for file_id, meta in st.session_state.database.items():
        file_table.append({
            "File ID": file_id,
            "Original Filename": meta["original_filename"],
            "Display Name": meta["display_name"],
            "Section": meta["section"],
            "Course Code": meta.get("course_code", ""),
            "Valid Students": meta["student_count"],
            "Uploaded At": meta["uploaded_at"],
        })

    st.dataframe(pd.DataFrame(file_table), use_container_width=True, hide_index=True)

    st.divider()

    st.markdown(
        """
        <div class="delete-zone">
            <h3 style="color:#7F1D1D;margin-bottom:0.25rem;">🗑️ Delete Uploaded Excel File</h3>
            <p style="color:#7F1D1D;font-weight:700;margin-bottom:0;">
                Select one uploaded Excel file below, then click delete. This will remove only the selected file from your current session.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    delete_options = {
        f"{meta['display_name']}": file_id
        for file_id, meta in st.session_state.database.items()
    }

    selected_delete_label = st.selectbox("Choose uploaded Excel file to delete", list(delete_options.keys()))
    selected_delete_id = delete_options[selected_delete_label]

    selected_meta = st.session_state.database[selected_delete_id]
    st.warning(
        f"You selected: {selected_meta['original_filename']} | {selected_meta['section']} | "
        f"{selected_meta['student_count']} valid students"
    )

    confirm_delete = st.checkbox("I confirm I want to delete this selected Excel file")

    if st.button("🗑️ Delete This Selected Excel File", disabled=not confirm_delete):
        del st.session_state.database[selected_delete_id]
        save_database(st.session_state.database)
        st.success("Selected Excel file deleted successfully.")
        st.rerun()

    st.divider()

    st.markdown(
        """
        <div class="delete-zone">
            <h3 style="color:#7F1D1D;margin-bottom:0.25rem;">⚠️ Delete All Uploaded Excel Files</h3>
            <p style="color:#7F1D1D;font-weight:700;margin-bottom:0;">
                Use this only if you want to clear all files in your current session.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    confirm_delete_all = st.checkbox("I confirm I want to delete ALL uploaded Excel files")

    if st.button("⚠️ Delete ALL Excel Files", disabled=not confirm_delete_all):
        st.session_state.database = {}
        save_database({})
        st.success("All session Excel files deleted successfully.")
        st.rerun()

    st.divider()
    st.subheader("Detection Report for Selected File")

    file_id = st.session_state.selected_file_id
    report = st.session_state.database[file_id].get("detection_report", pd.DataFrame())

    if isinstance(report, pd.DataFrame) and not report.empty:
        st.dataframe(report, use_container_width=True, hide_index=True)
    else:
        st.info("No detection report available.")