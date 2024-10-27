import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Load and validate Excel data
def load_loan_data(file):
    try:
        excel_data = pd.ExcelFile(file)
        
        # Check if the sheet 'Loan Data' exists
        if 'Loan Data' not in excel_data.sheet_names:
            st.error("The uploaded file must contain a sheet named 'Loan Data'.")
            return None
        
        # Load and clean the 'Loan Data' sheet
        loan_data = excel_data.parse('Loan Data')
        loan_data.columns = loan_data.iloc[0]  # Set the first row as header
        loan_data = loan_data.drop(0).reset_index(drop=True)  # Drop header row
        loan_data.columns = loan_data.columns.str.strip()  # Strip whitespace from headers
        
        # Required columns check
        required_columns = ['Loan ID', 'Loan Amount ($C)', 'Duration', 'Interest ($C)', 'Late Fee & Interest ($C)', 'Total Payment ($C)']
        if not all(col in loan_data.columns for col in required_columns):
            st.error("The 'Loan Data' sheet must contain all required columns.")
            return None
        
        # Convert relevant columns to numeric, handle errors gracefully
        for col in ['Loan Amount ($C)', 'Interest ($C)', 'Late Fee & Interest ($C)', 'Total Payment ($C)', 'Duration']:
            loan_data[col] = pd.to_numeric(loan_data[col], errors='coerce')
        
        # Drop rows with any NaN values in required columns
        loan_data.dropna(subset=required_columns, inplace=True)
        
        return loan_data[required_columns]
    
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

# Loan Calculator class
class LoanCalculator:
    def __init__(self, loan_amount, interest, late_fee, duration):
        self.loan_amount = loan_amount
        self.interest = interest
        self.late_fee = late_fee
        self.duration = duration

    def total_repayment(self):
        return self.loan_amount + self.interest + self.late_fee

    def calculate_apr(self):
        if self.loan_amount == 0 or self.duration == 0:
            return 0
        daily_interest = self.interest / self.loan_amount
        return (daily_interest * 365 / self.duration) * 100

    def calculate_monthly_payment(self):
        return self.total_repayment() / max((self.duration / 30), 1)

# ROI and Investment input handling
def sync_roi_and_investment():
    if 'initial_investment' not in st.session_state:
        st.session_state['initial_investment'] = 1000.0  # Default investment
    if 'roi_percentage' not in st.session_state:
        st.session_state['roi_percentage'] = 5.0  # Default ROI

    # Event listeners to sync ROI and Initial Investment
    def update_investment():
        st.session_state['initial_investment'] = (st.session_state['roi_percentage'] / 100) * st.session_state['investment_amount']
    
    def update_roi():
        st.session_state['roi_percentage'] = (st.session_state['initial_investment'] / st.session_state['investment_amount']) * 100

    st.sidebar.number_input("Initial Investment ($C)", value=st.session_state['initial_investment'], on_change=update_roi, key='initial_investment')
    st.sidebar.number_input("ROI (%)", value=st.session_state['roi_percentage'], on_change=update_investment, key='roi_percentage')

# Visualization functions
def plot_apr_comparison(apr, loan_id):
    apr_values = [apr, 30, 50, 100]  # Example APR values for comparison
    labels = [loan_id, "Competitor A", "Competitor B", "Industry Average"]
    
    fig, ax = plt.subplots(figsize=(8, 4))
    bars = ax.barh(labels, apr_values, color=['#4CAF50', '#FFC107', '#FF5722', '#2196F3'])
    ax.set_xlabel('APR (%)')
    ax.set_title('APR Comparison with Industry')
    ax.grid(axis='x', linestyle='--', alpha=0.7)
    for bar in bars:
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2, f'{bar.get_width():.2f}%', va='center')
    st.pyplot(fig)

# Streamlit interface
st.title("Payday Loan Dashboard")
st.sidebar.title("Loan Calculator Inputs")

# Sync ROI and investment inputs
sync_roi_and_investment()

# File uploader for loan data
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file:
    loan_data = load_loan_data(uploaded_file)
    if loan_data is not None:
        
        # Show loaded data
        st.write("### Loaded Loan Data")
        st.dataframe(loan_data)

        # Loan ID selection
        loan_ids = loan_data['Loan ID'].unique()
        selected_loan_id = st.sidebar.selectbox("Select Loan ID:", loan_ids)
        
        # Filter data for selected loan
        selected_loan_data = loan_data[loan_data['Loan ID'] == selected_loan_id].iloc[0]
        
        # Extract loan parameters
        loan_amount = selected_loan_data['Loan Amount ($C)']
        interest = selected_loan_data['Interest ($C)']
        late_fee = selected_loan_data['Late Fee & Interest ($C)'] if pd.notnull(selected_loan_data['Late Fee & Interest ($C)']) else 0.0
        duration = selected_loan_data['Duration']
        
        # Calculate metrics
        calculator = LoanCalculator(loan_amount, interest, late_fee, duration)
        total_repayment = calculator.total_repayment()
        apr = calculator.calculate_apr()
        monthly_payment = calculator.calculate_monthly_payment()
        
        # Display loan details
        st.write("### Loan Details")
        st.write(f"**Loan Amount:** ${loan_amount:,.2f}")
        st.write(f"**Interest:** ${interest:,.2f}")
        st.write(f"**Late Fee & Interest:** ${late_fee:,.2f}")
        st.write(f"**Duration:** {duration} days")
        st.write(f"**Total Repayment:** ${total_repayment:,.2f}")
        st.write(f"**Effective APR:** {apr:.2f}%")
        st.write(f"**Estimated Monthly Payment:** ${monthly_payment:,.2f}")
        
        # Display visualizations
        st.subheader("APR Comparison")
        plot_apr_comparison(apr, selected_loan_id)
else:
    st.write("Please upload an Excel file to proceed.")
