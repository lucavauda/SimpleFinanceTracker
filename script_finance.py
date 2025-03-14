from finance_tracker import FinanceTracker

tracker = FinanceTracker()
data = tracker.load_data("estrattoconto24.csv")
tracker.export_report_with_charts("my_financial_report_24.xlsx")