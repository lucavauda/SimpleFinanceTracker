import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
import io

class FinanceTracker:
    def __init__(self):
        self.data = None
        
    def load_data(self, file_path):
        """Load data from CSV or Excel file"""
        if file_path.endswith('.csv'):
            # Read CSV with European format (comma as decimal separator)
            self.data = pd.read_csv(file_path, decimal=',')
            
            # Convert date columns to datetime with European format
            self.data['Data contabile'] = pd.to_datetime(self.data['Data contabile'], format='%d/%m/%Y')
            self.data['Valuta'] = pd.to_datetime(self.data['Valuta'], format='%d/%m/%Y')
            
        elif file_path.endswith(('.xlsx', '.xls')):
            self.data = pd.read_excel(file_path)
            # Convert date columns if they're not already datetime
            if not pd.api.types.is_datetime64_any_dtype(self.data['Data contabile']):
                self.data['Data contabile'] = pd.to_datetime(self.data['Data contabile'], format='%d/%m/%Y')
                self.data['Valuta'] = pd.to_datetime(self.data['Valuta'], format='%d/%m/%Y')
        else:
            raise ValueError("Unsupported file format. Please use CSV or Excel.")
        
        # Process the data
        self._process_data()
        return self.data
    
    def _process_data(self):
        """Process and clean the data"""
        # Fix Dare and Avere columns to ensure they're numeric
        # Replace empty strings with NaN
        self.data['Dare'] = pd.to_numeric(self.data['Dare'], errors='coerce')
        self.data['Avere'] = pd.to_numeric(self.data['Avere'], errors='coerce')
        
        # Combine Dare and Avere into a single Amount column
        self.data['Amount'] = self.data['Avere'].fillna(0) + self.data['Dare'].fillna(0)
        
        # Extract month and year for analysis
        self.data['Month'] = self.data['Data contabile'].dt.month
        self.data['Year'] = self.data['Data contabile'].dt.year
        self.data['MonthYear'] = self.data['Data contabile'].dt.strftime('%Y-%m')
        
        # Drop unnecessary columns
        if 'Divisa' in self.data.columns and 'Causale' in self.data.columns:
            self.data = self.data.drop(['Divisa', 'Causale'], axis=1)
    
    def get_monthly_summary(self):
        """Get monthly income, expenses, and balance"""
        if self.data is None:
            return "No data loaded. Please load data first."
        
        monthly = self.data.groupby('MonthYear').agg({
            'Amount': 'sum',
            'Data contabile': 'count'
        }).rename(columns={'Data contabile': 'Transactions'})
        
        return monthly.sort_index()
    
    def get_category_summary(self):
        """Get summary by category"""
        if self.data is None:
            return "No data loaded. Please load data first."
        
        return self.data.groupby('Categoria').agg({
            'Amount': 'sum',
            'Data contabile': 'count'
        }).rename(columns={'Data contabile': 'Transactions'}).sort_values('Amount')
    
    def create_monthly_trend_chart(self):
        """Create monthly trend chart and return it as a figure"""
        if self.data is None:
            return None
        
        monthly = self.get_monthly_summary()
        
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.plot(monthly.index, monthly['Amount'], marker='o', linewidth=2)
        ax.set_title('Monthly Balance', fontsize=14)
        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Amount', fontsize=12)
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.tick_params(axis='x', rotation=45)
        plt.tight_layout()
        
        return fig
    
    def create_category_charts(self):
        """Create category breakdown charts and return them as figures"""
        if self.data is None:
            return None, None
        
        # Separate income and expenses
        expenses = self.data[self.data['Amount'] < 0].copy()
        income = self.data[self.data['Amount'] > 0].copy()
        
        # Create expense chart
        expense_by_cat = expenses.groupby('Categoria')['Amount'].sum().abs().sort_values(ascending=False)
        exp_fig, exp_ax = plt.subplots(figsize=(8, 8))
        
        if not expense_by_cat.empty:
            exp_data = expense_by_cat.head(min(5, len(expense_by_cat)))
            exp_ax.pie(exp_data.values, labels=exp_data.index, autopct='%1.1f%%', startangle=90)
            exp_ax.set_title('Top Expense Categories', fontsize=14)
            plt.tight_layout()
        else:
            exp_ax.text(0.5, 0.5, 'No expense data', horizontalalignment='center', verticalalignment='center')
            exp_ax.set_title('Expenses', fontsize=14)
        
        # Create income chart
        income_by_cat = income.groupby('Categoria')['Amount'].sum().sort_values(ascending=False)
        inc_fig, inc_ax = plt.subplots(figsize=(8, 8))
        
        if not income_by_cat.empty:
            inc_data = income_by_cat.head(min(5, len(income_by_cat)))
            inc_ax.pie(inc_data.values, labels=inc_data.index, autopct='%1.1f%%', startangle=90)
            inc_ax.set_title('Top Income Categories', fontsize=14)
            plt.tight_layout()
        else:
            inc_ax.text(0.5, 0.5, 'No income data', horizontalalignment='center', verticalalignment='center')
            inc_ax.set_title('Income', fontsize=14)
        
        return exp_fig, inc_fig
    
    def export_report_with_charts(self, output_path="my_financial_report.xlsx"):
        """Export financial report to Excel with embedded charts"""
        if self.data is None:
            return "No data loaded. Please load data first."
        
        # Create a Pandas Excel writer using the specified filename
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Get the workbook and add worksheets
            workbook = writer.book
            
            # Export transactions
            self.data.to_excel(writer, sheet_name='Transactions', index=False)
            
            # Export monthly summary
            monthly = self.get_monthly_summary()
            monthly.to_excel(writer, sheet_name='Monthly Summary')
            monthly_sheet = writer.sheets['Monthly Summary']
            
            # Export category summary
            cat_summary = self.get_category_summary()
            cat_summary.to_excel(writer, sheet_name='Category Summary')
            
            # Add income vs expense summary
            income = self.data[self.data['Amount'] > 0]['Amount'].sum()
            expense = self.data[self.data['Amount'] < 0]['Amount'].sum()
            balance = income + expense
            
            summary_df = pd.DataFrame({
                'Metric': ['Total Income', 'Total Expenses', 'Balance'],
                'Amount': [income, expense, balance]
            })
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Create Charts sheet
            chart_sheet = workbook.add_worksheet('Charts')
            
            # Add monthly trend chart
            monthly_fig = self.create_monthly_trend_chart()
            if monthly_fig:
                # Save the chart to a in-memory buffer
                imgdata = io.BytesIO()
                monthly_fig.savefig(imgdata, format='png')
                imgdata.seek(0)
                
                # Insert the image
                chart_sheet.insert_image('A1', '', {'image_data': imgdata, 'x_scale': 0.9, 'y_scale': 0.9})
                chart_sheet.set_row(0, 300)  # Set row height
            
            # Add category charts
            expense_fig, income_fig = self.create_category_charts()
            
            if expense_fig:
                # Save expense chart
                exp_imgdata = io.BytesIO()
                expense_fig.savefig(exp_imgdata, format='png')
                exp_imgdata.seek(0)
                
                # Insert the expense chart
                chart_sheet.insert_image('A20', '', {'image_data': exp_imgdata, 'x_scale': 0.8, 'y_scale': 0.8})
                chart_sheet.set_row(20, 300)  # Set row height
            
            if income_fig:
                # Save income chart
                inc_imgdata = io.BytesIO()
                income_fig.savefig(inc_imgdata, format='png')
                inc_imgdata.seek(0)
                
                # Insert the income chart
                chart_sheet.insert_image('J20', '', {'image_data': inc_imgdata, 'x_scale': 0.8, 'y_scale': 0.8})
            
            # Close the figures to free memory
            plt.close('all')
        
        return f"Report with charts exported to {output_path}"
    
    def get_basic_stats(self):
        """Get basic statistics about the financial data"""
        if self.data is None:
            return "No data loaded. Please load data first."
        
        income = self.data[self.data['Amount'] > 0]['Amount'].sum()
        expense = self.data[self.data['Amount'] < 0]['Amount'].sum()
        balance = income + expense
        
        avg_expense = self.data[self.data['Amount'] < 0]['Amount'].mean() if len(self.data[self.data['Amount'] < 0]) > 0 else 0
        avg_income = self.data[self.data['Amount'] > 0]['Amount'].mean() if len(self.data[self.data['Amount'] > 0]) > 0 else 0
        
        top_expense_cat = self.data[self.data['Amount'] < 0].groupby('Categoria')['Amount'].sum().abs().idxmax() if not self.data[self.data['Amount'] < 0].empty else "N/A"
        top_income_cat = self.data[self.data['Amount'] > 0].groupby('Categoria')['Amount'].sum().idxmax() if not self.data[self.data['Amount'] > 0].empty else "N/A"
        
        stats = {
            'Total Income': income,
            'Total Expenses': expense,
            'Balance': balance,
            'Average Expense': avg_expense,
            'Average Income': avg_income,
            'Top Expense Category': top_expense_cat,
            'Top Income Category': top_income_cat,
            'Total Transactions': len(self.data)
        }
        
        return stats

# Example usage
if __name__ == "__main__":
    tracker = FinanceTracker()
    
    # Sample code to test with the data
    # Create a test CSV file with the sample data
    sample_data = """Data contabile,Valuta,Dare,Avere,Divisa,Causale,Descrizione,Categoria,Tag
12/03/2025,10/03/2025,"-100,00",,EUR,VH,Pagamento POS,Ristoranti e bar,
12/03/2025,07/03/2025,"-30,00",,EUR,0U,PAGAMENTO VISA,Arte e Cultura,
11/03/2025,11/03/2025,"-11,50",,EUR,TE,ADDEBITO DIRETTO,Utenze"""
    
    with open("sample_data.csv", "w") as f:
        f.write(sample_data)
    
    # Load the sample data
    data = tracker.load_data("sample_data.csv")
    
    # Print basic statistics
    print("Basic Statistics:")
    for key, value in tracker.get_basic_stats().items():
        print(f"{key}: {value}")
    
    # Export report with charts
    result = tracker.export_report_with_charts("my_financial_report.xlsx")
    print(f"\n{result}")