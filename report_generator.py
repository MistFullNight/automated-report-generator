"""
Automated Support Report Generator
----------------------------------
Reads a CSV or Excel support-ticket file and generates a formatted Excel report.

Usage:
    python report_generator.py input/support_tickets.csv output/weekly_support_report.xlsx
"""

from pathlib import Path
import sys
import pandas as pd

REQUIRED_COLUMNS = [
    "Ticket ID", "Created Date", "Category", "Priority", "Status",
    "Assigned To", "Closed Date", "Resolution Hours",
    "SLA Target Hours", "SLA Met"
]

def load_data(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    elif path.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(path)
    else:
        raise ValueError("Input must be .csv, .xlsx, or .xls")

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce")
    df["Closed Date"] = pd.to_datetime(df["Closed Date"], errors="coerce")
    df["Resolution Hours"] = pd.to_numeric(df["Resolution Hours"], errors="coerce")
    return df

def safe_sla_rate(series: pd.Series) -> float:
    finished = series[series.isin(["Yes", "No"])]
    if finished.empty:
        return 0.0
    return (finished == "Yes").mean()

def build_report(df: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    total = len(df)
    closed = int((df["Status"] == "Closed").sum())
    avg_resolution = df.loc[df["Status"] == "Closed", "Resolution Hours"].mean()
    sla_rate = safe_sla_rate(df["SLA Met"])

    status_summary = (
        df.groupby("Status", dropna=False)
          .size()
          .reindex(["Closed", "Open", "In Progress", "Pending"], fill_value=0)
          .rename("Tickets")
          .reset_index()
    )
    status_summary["Share"] = status_summary["Tickets"] / max(total, 1)

    priority_summary = (
        df.groupby("Priority").size()
          .reindex(["Low", "Medium", "High", "Critical"], fill_value=0)
          .rename("Tickets")
          .reset_index()
    )

    agent_summary = (
        df.groupby("Assigned To")
          .agg(
              Assigned=("Ticket ID", "count"),
              Closed=("Status", lambda s: (s == "Closed").sum()),
              Avg_Resolution_Hours=("Resolution Hours", "mean"),
              SLA_Compliance=("SLA Met", safe_sla_rate),
          )
          .reset_index()
    )

    category_rows = []
    for category, group in df.groupby("Category"):
        category_rows.append({
            "Category": category,
            "Tickets": len(group),
            "Closed": int((group["Status"] == "Closed").sum()),
            "SLA Compliance": safe_sla_rate(group["SLA Met"]),
        })
    category_summary = pd.DataFrame(category_rows)

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm") as writer:
        workbook = writer.book

        # Formats
        title_fmt = workbook.add_format({
            "bold": True, "font_size": 18, "font_color": "white",
            "bg_color": "#17365D", "align": "center", "valign": "vcenter"
        })
        header_fmt = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": "#1F4E78",
            "align": "center", "valign": "vcenter", "border": 1
        })
        kpi_label_fmt = workbook.add_format({
            "bold": True, "bg_color": "#D9EAF7", "align": "center", "border": 1
        })
        kpi_value_fmt = workbook.add_format({
            "bold": True, "font_size": 18, "align": "center", "valign": "vcenter", "border": 1
        })
        percent_fmt = workbook.add_format({"num_format": "0.0%"})
        decimal_fmt = workbook.add_format({"num_format": "0.00"})

        # Raw data
        df.to_excel(writer, sheet_name="Raw Data", index=False)
        raw_ws = writer.sheets["Raw Data"]
        raw_ws.freeze_panes(1, 0)
        raw_ws.set_row(0, 22, header_fmt)
        raw_ws.set_column("A:A", 12)
        raw_ws.set_column("B:B", 20)
        raw_ws.set_column("C:F", 16)
        raw_ws.set_column("G:G", 20)
        raw_ws.set_column("H:J", 18)

        # Ticket Summary
        status_summary.to_excel(writer, sheet_name="Ticket Summary", index=False)
        sum_ws = writer.sheets["Ticket Summary"]
        sum_ws.set_row(0, 22, header_fmt)
        sum_ws.set_column("A:A", 18)
        sum_ws.set_column("B:B", 12)
        sum_ws.set_column("C:C", 14, percent_fmt)

        chart = workbook.add_chart({"type": "column"})
        chart.add_series({
            "name": "Tickets",
            "categories": ["Ticket Summary", 1, 0, len(status_summary), 0],
            "values": ["Ticket Summary", 1, 1, len(status_summary), 1],
        })
        chart.set_title({"name": "Ticket Status Distribution"})
        chart.set_legend({"none": True})
        sum_ws.insert_chart("E2", chart)

        # Agent Performance
        agent_summary.to_excel(writer, sheet_name="Agent Performance", index=False)
        agent_ws = writer.sheets["Agent Performance"]
        agent_ws.set_row(0, 22, header_fmt)
        agent_ws.set_column("A:A", 22)
        agent_ws.set_column("B:C", 12)
        agent_ws.set_column("D:D", 22, decimal_fmt)
        agent_ws.set_column("E:E", 18, percent_fmt)

        # Category Analysis
        category_summary.to_excel(writer, sheet_name="Category Analysis", index=False)
        cat_ws = writer.sheets["Category Analysis"]
        cat_ws.set_row(0, 22, header_fmt)
        cat_ws.set_column("A:A", 20)
        cat_ws.set_column("B:C", 12)
        cat_ws.set_column("D:D", 18, percent_fmt)

        # Dashboard
        dash = workbook.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dash
        dash.merge_range("A1:H2", "Weekly Support Operations Dashboard", title_fmt)
        dash.write("A4", "Total Tickets", kpi_label_fmt)
        dash.write("C4", "Closed Tickets", kpi_label_fmt)
        dash.write("E4", "Avg Resolution (hrs)", kpi_label_fmt)
        dash.write("G4", "SLA Compliance", kpi_label_fmt)
        dash.merge_range("A5:B6", total, kpi_value_fmt)
        dash.merge_range("C5:D6", closed, kpi_value_fmt)
        dash.merge_range("E5:F6", 0 if pd.isna(avg_resolution) else float(avg_resolution), kpi_value_fmt)
        dash.merge_range("G5:H6", float(sla_rate), workbook.add_format({
            "bold": True, "font_size": 18, "align": "center",
            "valign": "vcenter", "border": 1, "num_format": "0.0%"
        }))

        dash.write_row("A9", priority_summary.columns.tolist(), header_fmt)
        for row_index, row in enumerate(priority_summary.itertuples(index=False), start=9):
            dash.write(row_index, 0, row.Priority)
            dash.write(row_index, 1, int(row.Tickets))

        pchart = workbook.add_chart({"type": "column"})
        pchart.add_series({
            "name": "Tickets",
            "categories": ["Dashboard", 9, 0, 12, 0],
            "values": ["Dashboard", 9, 1, 12, 1],
        })
        pchart.set_title({"name": "Tickets by Priority"})
        pchart.set_legend({"none": True})
        dash.insert_chart("D9", pchart)

        dash.set_column("A:H", 18)

def main():
    if len(sys.argv) != 3:
        print("Usage: python report_generator.py <input.csv|xlsx> <output.xlsx>")
        raise SystemExit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    data = load_data(input_path)
    build_report(data, output_path)
    print(f"Report created: {output_path}")

if __name__ == "__main__":
    main()
