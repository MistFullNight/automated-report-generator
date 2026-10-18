# Automated Report Generator

A portfolio project that turns raw support-ticket data into a management-ready Excel report.

## What it demonstrates
- Python automation
- Excel reporting
- Data cleaning and validation
- KPI calculation
- Operational dashboards
- Charts and management summaries

## Input
The sample input contains support tickets with:
- Ticket ID
- Created / closed dates
- Category
- Priority
- Status
- Assigned agent
- Resolution hours
- SLA target
- SLA result

## Output
The generated workbook contains:
1. Dashboard
2. Ticket Summary
3. Agent Performance
4. Category Analysis
5. Raw Data

## Run it

```bash
pip install -r requirements.txt
python report_generator.py input/support_tickets.csv output/weekly_support_report.xlsx
```

You can also use an `.xlsx` input file with the same column names.

## Portfolio demo idea
Record a 30–60 second screen demo:
1. Show the raw input file.
2. Run the Python command.
3. Open the generated Excel report.
4. Show the dashboard, agent performance, and category analysis.

## Next improvements
- PDF export
- Date-range filtering
- Interactive UI
- Automatic email delivery
- Configurable SLA rules
