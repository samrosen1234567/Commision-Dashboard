#!/usr/bin/env python3
import PySimpleGUI as sg
import pandas as pd
from datetime import datetime

# GUI layout
layout = [
    [sg.Text('Commission Reconciliation App', font=('Helvetica', 16))],
    [sg.Text('Select policies.xlsx:'), sg.Input(key='-POLICIES-'), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
    [sg.Text('Select Paid_policies.xlsx:'), sg.Input(key='-PAID-'), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
    [sg.Button('Reconcile'), sg.Button('Save Results'), sg.Button('Exit')],
    [sg.Text('Monthly Commission Summary:')],
    [sg.Table(values=[], headings=['Month','Expected','Paid','Outstanding'], auto_size_columns=True, key='-SUMMARY-', num_rows=10)],
    [sg.Text('Outstanding New Policies (≤ 90 days old):')],
    [sg.Table(values=[], headings=['PolicyID','SaleDate','ExpectedCommission','DaysSinceSale'], auto_size_columns=True, key='-OUTSTANDING-', num_rows=10)]
]

window = sg.Window('Commission Reconciliation', layout, finalize=True)

results = {}

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == 'Reconcile':
        pol_file = values['-POLICIES-']
        paid_file = values['-PAID-']
        try:
            # Read policies
            policies = pd.read_excel(pol_file, sheet_name="Complete Detail", header=1)
            policies = policies.rename(columns={'MainPolicyNumber':'PolicyID',
                                                'EffectiveDate':'SaleDate',
                                                'PremiumEstimated':'GrossPremium',
                                                'PolicyAndLineTypes':'LineType'})
            policies['SaleDate'] = pd.to_datetime(policies['SaleDate'])
            # Detect commission rate column
            comm_col = next((c for c in policies.columns if 'Rate' in c or c.strip().endswith('%')), None)
            rates = policies[comm_col].astype(str).str.rstrip('%').astype(float)
            policies['CommissionRate'] = rates.div(100) if rates.max()>1 else rates
            policies['ExpectedCommission'] = policies['GrossPremium'] * policies['CommissionRate']
            # Read paid policies
            summary = pd.read_excel(paid_file, sheet_name="Producer Summary", header=None)
            statement_date = pd.to_datetime(summary.iloc[3,2]).normalize()
            paid = pd.read_excel(paid_file, sheet_name="Alex Benedict", header=1).rename(columns={'PolicyNumber':'PolicyID','Pr1$':'CommissionPaid'})
            paid['PaymentDate'] = statement_date
            # Monthly summary
            exp = policies.groupby(policies['SaleDate'].dt.to_period('M').dt.to_timestamp())['ExpectedCommission'].sum()
            pay = paid.groupby(paid['PaymentDate'].dt.to_period('M').dt.to_timestamp())['CommissionPaid'].sum()
            summary_df = pd.concat([exp, pay], axis=1).fillna(0)
            summary_df['Outstanding'] = summary_df['ExpectedCommission'] - summary_df['CommissionPaid']
            summary_df = summary_df.reset_index().rename(columns={'index':'Month','ExpectedCommission':'Expected','CommissionPaid':'Paid'})
            # Outstanding new policies
            unpaid = policies[~policies['PolicyID'].isin(paid['PolicyID'])].copy()
            today = pd.Timestamp.now().normalize()
            unpaid['DaysSinceSale'] = (today - unpaid['SaleDate']).dt.days
            out_90 = unpaid[unpaid['DaysSinceSale']<=90]
            # Update GUI tables
            window['-SUMMARY-'].update(values=summary_df.values.tolist())
            window['-OUTSTANDING-'].update(values=out_90[['PolicyID','SaleDate','ExpectedCommission','DaysSinceSale']].astype(str).values.tolist())
            results['summary'] = summary_df
            results['outstanding'] = out_90
        except Exception as e:
            sg.popup_error(f"Error: {e}")
    if event == 'Save Results' and results:
        fname = sg.popup_get_file('Save summary as', save_as=True, default_extension='.xlsx', file_types=(('Excel','.xlsx'),))
        if fname:
            with pd.ExcelWriter(fname) as writer:
                results['summary'].to_excel(writer, sheet_name='CommissionSummary', index=False)
                results['outstanding'].to_excel(writer, sheet_name='OutstandingNew', index=False)
            sg.popup('Saved results to ' + fname)

window.close()
