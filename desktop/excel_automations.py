"""Excel automation implementations for Workforce Junction.

Implements:
1. Salary File Automation (compare_salary_files + fill_template_file)
2. Enrollment Report Audit (generate_output with SSN matching and plan audit)
3. Pre vs Post Comparison Tool (two-way comparison with highlighted columns)
"""

from __future__ import annotations

import datetime as dt
import os
import re
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Callable, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

if TYPE_CHECKING:
    from api import UI


# ==============================================================================
# Helper Functions & Date Formatting
# ==============================================================================

def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        str(col).replace("\xa0", " ").replace("\u200b", "").strip()
        for col in df.columns
    ]
    return df


def normalize_ssn(ssn: object) -> str:
    if pd.isna(ssn):
        return ""
    return "".join(filter(str.isalnum, str(ssn))).lower()


def parse_date_obj(val: object) -> Optional[dt.date]:
    """Parse any datetime/date string or object into a python date object."""
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (dt.date, dt.datetime)):
        return val if isinstance(val, dt.date) else val.date()
    val_str = str(val).replace("\xa0", " ").replace("\u200b", "").strip()
    if not val_str or val_str.lower() in [
        "", "nan", "nat", "none", "null", "#n/a",
        "new employee", "column missing", "true", "false"
    ]:
        return None
    try:
        if re.search(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}", val_str):
            ts = pd.to_datetime(val_str, errors="coerce")
            if pd.notna(ts):
                return ts.date()
    except Exception:
        pass
    return None


def normalize_text(val: object) -> str:
    """Normalize text and dates for accurate equivalence comparison."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).replace("\xa0", " ").replace("\u200b", "").strip()
    # If it's a date or datetime representation, normalize to MM/DD/YYYY
    d = parse_date_obj(s)
    if d is not None:
        return d.strftime("%m/%d/%Y")
    return s.lower()


def format_excel_dates_and_styling(
    file_path: str,
    yellow_condition: Optional[Callable[[str], bool]] = None,
) -> None:
    """
    Standardizes all date cells in an Excel workbook to native Excel Date format (MM/DD/YYYY).
    This ensures Excel's AutoFilter recognizes them as friendly dates with Year > Month > Day grouping.
    Also applies header highlighting based on yellow_condition.
    """
    wb = load_workbook(file_path)
    ws = wb.active
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    date_col_indices = set()
    for col_idx in range(1, ws.max_column + 1):
        header_val = str(ws.cell(row=1, column=col_idx).value or "").strip()
        header_lower = header_val.lower()
        if "compare" not in header_lower and (
            "date" in header_lower
            or "dob" in header_lower
            or "birth" in header_lower
            or "hire" in header_lower
            or "term" in header_lower
            or "effective" in header_lower
        ):
            date_col_indices.add(col_idx)

    for col_idx in range(1, ws.max_column + 1):
        is_date_col = col_idx in date_col_indices
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = cell.value
            if val is None:
                continue

            parsed_d = None
            if is_date_col:
                parsed_d = parse_date_obj(val)
            elif isinstance(val, str) and re.search(r"^\d{4}-\d{2}-\d{2}", val.strip()):
                parsed_d = parse_date_obj(val)

            if parsed_d is not None:
                cell.value = parsed_d
                cell.number_format = "MM/DD/YYYY"

    if yellow_condition:
        for cell in ws[1]:
            val_str = str(cell.value or "")
            if yellow_condition(val_str):
                cell.fill = yellow

    wb.save(file_path)


def to_number(series: pd.Series) -> pd.Series:
    """Convert '$242,000.00' or '242000' strings to plain numeric strings."""
    return (
        series.astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )


def get_status_column(df: pd.DataFrame) -> str:
    possible = ["Enrollment Status", "Employment Status", "Status"]
    for c in df.columns:
        if str(c).strip().lower() in [p.lower() for p in possible]:
            return c
    raise ValueError(f"Status column not found in columns: {list(df.columns)}")


def is_sp_or_ch(name: str) -> bool:
    low = name.lower()
    return bool(re.search(r"\b(sp|ch)\b", low))


def choose_er_pays_column(cols: list[str]) -> Optional[str]:
    if not cols:
        return None
    filtered = [c for c in cols if not is_sp_or_ch(c)]
    if filtered:
        cols = filtered
    for c in cols:
        if re.search(r"\bee\b", c, re.I) or re.search(r"employee", c, re.I):
            return c
    return cols[0]


def choose_volume_column(cols: list[str]) -> Optional[str]:
    if not cols:
        return None
    filtered = [c for c in cols if not is_sp_or_ch(c)]
    if filtered:
        return filtered[0]
    return cols[0]


# ==============================================================================
# 1. Salary File Automation
# ==============================================================================

def compare_salary_files(pre_file: str, post_file: str, ui: Optional["UI"] = None) -> Tuple[str, pd.DataFrame]:
    """
    Compares Salary File (Pre) with BenJ Work File (Post).
    Pre: normal header at row 1 (index 0)
    Post: header on row 4 (index 3)
    """
    if ui:
        ui.log("Reading Salary File and BenJ Work File...")

    pre = pd.read_excel(pre_file, dtype=str)
    post = pd.read_excel(post_file, dtype=str, header=3)

    pre = clean_column_names(pre)
    post = clean_column_names(post)

    def get_emp_id_col(cols):
        for name in ["Employee ID", "Employee Id", "Emp ID", "Emp Id"]:
            if name in cols:
                return name
        return None

    def get_ssn_col(cols):
        for name in ["SSN", "SSN#", "Social Security Number"]:
            if name in cols:
                return name
        return None

    pre_emp_col = get_emp_id_col(pre.columns)
    post_emp_col = get_emp_id_col(post.columns)
    pre_ssn_col = get_ssn_col(pre.columns)
    post_ssn_col = get_ssn_col(post.columns)

    if not post_emp_col:
        raise ValueError(
            f"Post file must contain Employee ID / Employee Id column.\nFound: {list(post.columns)}"
        )

    # Standardize keys
    post = post.rename(columns={post_emp_col: "Employee ID"})
    post["Employee ID"] = post["Employee ID"].astype(str)

    if post_ssn_col:
        post = post.rename(columns={post_ssn_col: "SSN"})
        post["SSN"] = post["SSN"].astype(str)

    if pre_emp_col:
        pre = pre.rename(columns={pre_emp_col: "Employee ID"})
        pre["Employee ID"] = pre["Employee ID"].astype(str)
    elif pre_ssn_col:
        pre = pre.rename(columns={pre_ssn_col: "SSN"})
        pre["SSN"] = pre["SSN"].astype(str)
    else:
        raise ValueError(
            f"Pre file must contain either Employee ID or SSN column.\nFound: {list(pre.columns)}"
        )

    if "Annual Salary" not in pre.columns:
        raise ValueError(f"Pre file must contain 'Annual Salary' column. Found: {list(pre.columns)}")
    if "Salary" not in post.columns:
        raise ValueError(f"Post file must contain 'Salary' column. Found: {list(post.columns)}")

    PRE_SAL_COL = "Annual Salary"
    POST_SAL_COL = "Salary"
    WORK_STATUS_COL = "Status"
    WORK_SALPERIOD_COL = "Salary Period"

    pre[PRE_SAL_COL] = to_number(pre[PRE_SAL_COL])
    post[POST_SAL_COL] = to_number(post[POST_SAL_COL])

    post_lookup = post.set_index("Employee ID")
    post_lookup_ssn = post.set_index("SSN") if "SSN" in post.columns else None

    comparison_salary = []
    salary_from_benj = []
    status_from_benj = []
    salperiod_from_benj = []
    diff_values = []
    results = []
    comments = []

    for _, row in pre.iterrows():
        pre_sal = row[PRE_SAL_COL]
        post_row = None

        if "Employee ID" in pre.columns:
            emp_id = str(row["Employee ID"])
            if emp_id in post_lookup.index:
                post_row = post_lookup.loc[emp_id]
                if isinstance(post_row, pd.DataFrame):
                    post_row = post_row.iloc[0]

        if post_row is None and "SSN" in pre.columns and post_lookup_ssn is not None:
            ssn_val = str(row["SSN"])
            if ssn_val in post_lookup_ssn.index:
                post_row = post_lookup_ssn.loc[ssn_val]
                if isinstance(post_row, pd.DataFrame):
                    post_row = post_row.iloc[0]

        if post_row is not None:
            post_sal = post_row[POST_SAL_COL]
            if isinstance(post_sal, pd.Series):
                post_sal = post_sal.iloc[0]

            work_status = post_row.get(WORK_STATUS_COL, "")
            work_salperiod = post_row.get(WORK_SALPERIOD_COL, "")
            if isinstance(work_status, pd.Series):
                work_status = work_status.iloc[0]
            if isinstance(work_salperiod, pd.Series):
                work_salperiod = work_salperiod.iloc[0]

            comparison_salary.append(pre_sal)
            salary_from_benj.append(post_sal)
            status_from_benj.append(work_status)
            salperiod_from_benj.append(work_salperiod)

            try:
                pre_num = float(pre_sal)
                post_num = float(post_sal)
                same = round(pre_num, 2) == round(post_num, 2)
                diff_values.append(round(post_num - pre_num, 2))
            except ValueError:
                same = normalize_text(pre_sal) == normalize_text(post_sal)
                diff_values.append("")

            results.append(same)
            comments.append("" if same else "Salary changed")
        else:
            comparison_salary.append("")
            salary_from_benj.append("")
            status_from_benj.append("")
            salperiod_from_benj.append("")
            diff_values.append("")
            results.append(False)
            comments.append("New Employee (not in post file)")

    result_df = pre.copy()
    result_df["Salary From Salary File"] = comparison_salary
    result_df["Salary From WorkFile"] = salary_from_benj
    result_df["Status From WorkFile"] = status_from_benj
    result_df["Salary Period From Benj"] = salperiod_from_benj
    result_df["Diffrence"] = diff_values
    result_df["Results"] = results
    result_df["Comment"] = comments

    downloads = Path.home() / "Downloads"
    downloads.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = str(downloads / f"Salary_Comparison_{timestamp}.xlsx")
    result_df.to_excel(output_file, index=False)

    format_excel_dates_and_styling(
        output_file,
        yellow_condition=lambda val: val in [
            "Salary From Salary File",
            "Salary From WorkFile",
            "Status From WorkFile",
            "Salary Period From Benj",
            "Diffrence",
            "Results",
            "Comment",
        ],
    )

    if ui:
        ui.log(f"✓ Step 1 Complete: Comparison saved to {output_file}", "ok")

    return output_file, result_df


def fill_template_file(comparison_file: str, template_file: str, ui: Optional["UI"] = None) -> str:
    """
    From comparison file:
      - Take rows where Status From WorkFile = Active AND Results = False
      - Writes Employee Id / SSN / Salary / Salary Period into template
    """
    if ui:
        ui.log("Filtering Active + False records and filling template...")

    comparison_df = pd.read_excel(comparison_file, dtype=str)
    comparison_df = clean_column_names(comparison_df)

    active_false = comparison_df[
        (comparison_df["Status From WorkFile"].astype(str).str.strip().str.lower() == "active")
        & (comparison_df["Results"].astype(str).str.strip().str.lower() == "false")
    ].copy()

    if active_false.empty:
        msg = "No records found with Status=Active AND Results=FALSE."
        if ui:
            ui.log(msg, "warn")
        return ""

    template_df = pd.read_excel(template_file, dtype=str)
    template_df = clean_column_names(template_df)

    if "Country ID" in template_df.columns and "SSN" not in template_df.columns:
        template_df = template_df.rename(columns={"Country ID": "SSN"})

    emp_id_col = None
    for name in ["Employee Id", "Employee ID", "Emp ID", "Emp Id"]:
        if name in template_df.columns:
            emp_id_col = name
            break

    ssn_col = None
    for name in ["SSN", "Ssn", "Social Security Number"]:
        if name in template_df.columns:
            ssn_col = name
            break

    sal_col = None
    for name in ["Salary"]:
        if name in template_df.columns:
            sal_col = name
            break

    sal_period_col = None
    for name in ["Salary Period"]:
        if name in template_df.columns:
            sal_period_col = name
            break

    if not sal_col or not sal_period_col:
        raise ValueError(
            f"Template must have at least 'Salary' and 'Salary Period' columns.\nFound: {list(template_df.columns)}"
        )

    rename_map = {}
    if emp_id_col:
        rename_map[emp_id_col] = "Employee Id"
    if ssn_col:
        rename_map[ssn_col] = "SSN"
    rename_map[sal_col] = "Salary"
    rename_map[sal_period_col] = "Salary Period"
    template_df = template_df.rename(columns=rename_map)

    if "Employee Id" not in template_df.columns:
        template_df["Employee Id"] = ""
    if "SSN" not in template_df.columns and "SSN" in comparison_df.columns:
        template_df["SSN"] = ""

    filled_count = 0

    for _, comp_row in active_false.iterrows():
        emp_id_val = str(comp_row.get("Employee ID", "")).strip()
        ssn_val = str(comp_row.get("SSN", "")).strip()

        if not emp_id_val and not ssn_val:
            continue

        salary = str(comp_row.get("Salary From Salary File", "")).strip()
        sal_period = str(comp_row.get("Salary Period From Benj", "")).strip()

        empty_mask = (
            template_df["Employee Id"].astype(str).str.strip().eq("")
            & template_df["SSN"].astype(str).str.strip().eq("")
        )

        if empty_mask.any():
            idx = empty_mask.idxmax()
        else:
            new_row = {col: "" for col in template_df.columns}
            template_df = pd.concat([template_df, pd.DataFrame([new_row])], ignore_index=True)
            idx = template_df.index[-1]

        if emp_id_val:
            template_df.at[idx, "Employee Id"] = emp_id_val
        if ssn_val:
            template_df.at[idx, "SSN"] = ssn_val
        template_df.at[idx, "Salary"] = salary
        template_df.at[idx, "Salary Period"] = sal_period
        filled_count += 1

    output_file = str(
        Path(template_file).parent
        / f"{Path(template_file).stem}_Filled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    template_df.to_excel(output_file, index=False)

    wb = load_workbook(output_file)
    ws = wb.active
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for cell in ws[1]:
        if cell.value in ["Employee Id", "SSN", "Salary", "Salary Period"]:
            col_idx = cell.column
            cell.fill = yellow
            for row_idx in range(2, ws.max_row + 1):
                c = ws.cell(row=row_idx, column=col_idx)
                if c.value and str(c.value).strip():
                    c.fill = yellow

    wb.save(output_file)
    format_excel_dates_and_styling(output_file)

    if ui:
        ui.log(f"✓ Step 2 Complete: {filled_count} record(s) written to {output_file}", "ok")

    return str(output_file)


def run_salary_automation(salary_file: str, work_file: str, template_file: str, ui: "UI") -> None:
    """Executes the full salary automation workflow."""
    ui.log("Starting Salary File Automation...")
    comparison_path, _ = compare_salary_files(salary_file, work_file, ui)
    filled_template_path = fill_template_file(comparison_path, template_file, ui)
    if filled_template_path:
        ui.log("✓ All steps completed successfully!", "ok")
        ui.status("Salary automation finished", "pass")
    else:
        ui.log("✓ Comparison done (no active mismatch records to fill into template).", "info")
        ui.status("Salary automation completed", "pass")


# ==============================================================================
# 2. Enrollment Report Audit
# ==============================================================================

def run_enrollment_audit(pre_file: str, post_file: str, ui: "UI") -> str:
    """
    Audits active enrollment plan columns across Pre and Post enrollment reports.
    """
    ui.log("Reading Pre and Post enrollment reports (header on row 5)...")
    pre = pd.read_excel(pre_file, header=4, dtype=str)
    post = pd.read_excel(post_file, header=4, dtype=str)

    pre = clean_column_names(pre)
    post = clean_column_names(post)

    status_pre = get_status_column(pre)
    status_post = get_status_column(post)

    pre_active = pre[pre[status_pre].str.strip().str.lower() == "active"].copy()
    post_active = post[post[status_post].str.strip().str.lower() == "active"].copy()

    if "Employee SSN" not in pre_active.columns or "Employee SSN" not in post_active.columns:
        raise ValueError("Both files must contain 'Employee SSN' column.")

    pre_active["ssn_norm"] = pre_active["Employee SSN"].map(normalize_ssn)
    post_active["ssn_norm"] = post_active["Employee SSN"].map(normalize_ssn)

    ui.log(f"Active records: Pre={len(pre_active)}, Post={len(post_active)}")

    plan_name_cols = [
        col for col in post_active.columns
        if str(col).strip().lower().endswith("plan name")
    ]

    if not plan_name_cols:
        raise ValueError("No plan columns ending with 'Plan Name' found in Post file")

    excluded_prefixes = ("medical", "dental", "vision")
    plan_name_cols = [p for p in plan_name_cols if not p.strip().lower().startswith(excluded_prefixes)]

    basic_suffixes = ["Plan Name", "Effective Date", "Coverage"]

    for plan_col in plan_name_cols:
        plan_prefix = plan_col.replace("Plan Name", "").strip()
        plan_columns = []

        for suf in basic_suffixes:
            col_name = f"{plan_prefix} {suf}" if suf != "Plan Name" else plan_col
            if col_name in post_active.columns:
                plan_columns.append(col_name)

        er_candidates = [
            c for c in post_active.columns
            if c.lower().startswith(plan_prefix.lower()) and c.lower().strip().endswith("er pays")
        ]
        er_col = choose_er_pays_column(er_candidates)
        if er_col:
            plan_columns.append(er_col)

        vol_candidates = [
            c for c in post_active.columns
            if c.lower().startswith(plan_prefix.lower()) and c.lower().strip().endswith("volume")
        ]
        vol_col = choose_volume_column(vol_candidates)
        if vol_col:
            plan_columns.append(vol_col)

        if not plan_columns:
            continue

        for col in plan_columns:
            original_post_series = post_active[col].copy()
            if col in pre_active.columns:
                pre_lookup = dict(zip(pre_active["ssn_norm"], pre_active[col]))
            else:
                pre_lookup = {}

            pre_values = []
            compare_results = []

            for idx in post_active.index:
                ssn_norm = post_active.at[idx, "ssn_norm"]
                post_val = original_post_series.at[idx]

                if ssn_norm and ssn_norm in pre_lookup:
                    pre_val = pre_lookup[ssn_norm]
                    pre_values.append(pre_val if pd.notna(pre_val) else "#N/A")
                    compare_results.append(normalize_text(post_val) == normalize_text(pre_val))
                else:
                    pre_values.append("#N/A")
                    compare_results.append("#N/A")

            post_active[col] = pre_values

            insert_idx = post_active.columns.get_loc(col) + 1
            post_col_name = f"{col} (Post)"
            if post_col_name in post_active.columns:
                i = 1
                while f"{post_col_name} ({i})" in post_active.columns:
                    i += 1
                post_col_name = f"{post_col_name} ({i})"

            post_active.insert(insert_idx, post_col_name, original_post_series.values)

            compare_col_name = f"{col} Compare"
            if compare_col_name in post_active.columns:
                i = 1
                while f"{compare_col_name} ({i})" in post_active.columns:
                    i += 1
                compare_col_name = f"{compare_col_name} ({i})"

            post_active.insert(insert_idx + 1, compare_col_name, compare_results)

    downloads = Path.home() / "Downloads"
    downloads.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = str(downloads / f"PrePost_Comparison_{timestamp}.xlsx")

    post_active.to_excel(output_file, index=False)

    format_excel_dates_and_styling(
        output_file,
        yellow_condition=lambda val: "(Post)" in val or " Compare" in val,
    )

    ui.log(f"✓ Enrollment Audit completed successfully!\nSaved to: {output_file}", "ok")
    ui.status("Enrollment Audit complete", "pass")
    return output_file


# ==============================================================================
# 3. Pre vs Post Comparison Tool (Bidirectional)
# ==============================================================================

def remove_duplicates_by_employee_name(df: pd.DataFrame) -> pd.DataFrame:
    if "Employee Name" in df.columns:
        return df.drop_duplicates(subset=["Employee Name"]).copy()
    return df


def create_comparison_for_base(
    base_df: pd.DataFrame,
    lookup_df: pd.DataFrame,
    lookup_label: str,
    compare_cols: list[str],
) -> pd.DataFrame:
    result_df = base_df.copy()
    lookup_maps = {}
    for col in compare_cols:
        if col in lookup_df.columns:
            lookup_maps[col] = dict(zip(lookup_df["ssn_norm"], lookup_df[col]))
        else:
            lookup_maps[col] = {}

    for col in compare_cols:
        col_exists_in_base = col in base_df.columns
        insert_idx = result_df.columns.get_loc(col) + 1 if col_exists_in_base else len(result_df.columns)

        post_values = []
        compare_results = []

        lmap = lookup_maps[col]

        for _, row in result_df.iterrows():
            ssn_norm = row.get("ssn_norm", "")
            if ssn_norm and ssn_norm in lmap:
                lookup_val = lmap[ssn_norm]
                is_na = pd.isna(lookup_val) or str(lookup_val).strip() == ""
                post_values.append("Column Missing" if is_na else lookup_val)
                base_val = row[col] if col_exists_in_base else ""
                compare_results.append(normalize_text(base_val) == normalize_text(lookup_val))
            else:
                post_values.append("New Employee")
                compare_results.append("New Employee")

        post_col_name = f"{col} ({lookup_label})"
        compare_col_name = f"{col} Compare"

        result_df.insert(insert_idx, post_col_name, post_values)
        result_df.insert(insert_idx + 1, compare_col_name, compare_results)

    return result_df


def run_comparison_tool(pre_file: str, post_file: str, ui: "UI") -> Tuple[str, str]:
    """
    Produces bidirectional Pre vs Post comparison files (PRE_TO_POST & POST_TO_PRE).
    """
    ui.log("Reading Pre and Post files for bidirectional comparison (header on row 5)...")
    pre = pd.read_excel(pre_file, header=4, dtype=str)
    post = pd.read_excel(post_file, header=4, dtype=str)

    pre = clean_column_names(pre)
    post = clean_column_names(post)

    pre = remove_duplicates_by_employee_name(pre)
    post = remove_duplicates_by_employee_name(post)

    if "Employee SSN" not in pre.columns or "Employee SSN" not in post.columns:
        raise ValueError("Employee SSN column must exist in both files for matching")

    pre["ssn_norm"] = pre["Employee SSN"].map(normalize_ssn)
    post["ssn_norm"] = post["Employee SSN"].map(normalize_ssn)

    status_pre = get_status_column(pre)
    status_post = get_status_column(post)

    if status_pre.strip().lower() == "employment status":
        pre_active = pre[pre[status_pre].str.strip().str.lower() == "active"].copy()
        post_active = post[post[status_post].str.strip().str.lower() == "active"].copy()
        ui.log("Status column is 'Employment Status' — filtering to active records only.")
    else:
        pre_active = pre.copy()
        post_active = post.copy()
        ui.log(f"Status column is '{status_pre}' — no active-only filter applied.", "warn")

    ui.log(f"Records after filter: Pre={len(pre_active)}, Post={len(post_active)}")

    COMPARE_COLS = ["Plan Name", "Coverage", "Effective Date", "Enrollment Status"]

    pre_to_post_df = create_comparison_for_base(pre_active, post_active, "Post", COMPARE_COLS)
    post_to_pre_df = create_comparison_for_base(post_active, pre_active, "Pre", COMPARE_COLS)

    downloads = Path.home() / "Downloads"
    downloads.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pre_stem = Path(pre_file).stem
    post_stem = Path(post_file).stem

    out1 = str(downloads / f"{pre_stem}_vs_{post_stem}_{timestamp}_PRE_TO_POST.xlsx")
    out2 = str(downloads / f"{post_stem}_vs_{pre_stem}_{timestamp}_POST_TO_PRE.xlsx")

    pre_to_post_df.to_excel(out1, index=False)
    post_to_pre_df.to_excel(out2, index=False)

    for fpath in [out1, out2]:
        format_excel_dates_and_styling(
            fpath,
            yellow_condition=lambda val: val.endswith("(Post)") or val.endswith("(Pre)") or "Compare" in val,
        )

    ui.log(f"✓ Pre → Post report saved to: {out1}", "ok")
    ui.log(f"✓ Post → Pre report saved to: {out2}", "ok")
    ui.status("Comparison complete", "pass")

    return out1, out2
