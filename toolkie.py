import streamlit as st
import numpy as np
import pandas as pd
import openpyxl
from io import BytesIO

st.title("Forecasting App with Full Calculations")

st.write("## Upload Your Raw Data")
uploaded_file = st.file_uploader("Upload a single Excel file containing all raw data", type=["xlsx", "xls"])

st.write("## Parameters")

historical_horizon_period_start = st.number_input("Historical Horizon Start (Fin_Yr_Wk eg 202430)", value=202530, step=1)
historical_horizon_period_end = st.number_input("Historical Horizon End (Fin_Yr_Wk eg 202430)", value=202533, step=1)
min_acceptable_margin = st.number_input("Minimum Acceptable Margin (e.g. 0.29 for 29%)", value=0.29, step=0.01)
expected_period_start = st.number_input("Expected Period Start (Fin_Yr_Wk eg 202430)", value=202529, step=1)
expected_period_end = st.number_input("Expected Period End (Fin_Yr_Wk eg 202430)", value=202552, step=1)
fin_year = st.number_input("Financial Year (e.g. 2025)", value=2025, step=1)
stock_from_finyear = st.number_input("Stock From Financial Year (e.g. 2025)", value=2025, step=1)
stock_from_week = st.number_input("Stock From Week (e.g. 14)", value=14, step=1)

if st.button("Forecast"):
    if uploaded_file is not None:
        # Read the uploaded Excel file into a DataFrame
        sales_data_df = pd.read_excel(uploaded_file)
        
        # --- Original Code Logic Integration Starts Here ---
        
        sales_data_df['YearWeek'] = sales_data_df['Fin Year'] * 100 + sales_data_df['Week']
        # The code expects 'Fin Year' and 'Week' columns to create YearWeek.
        # If your data schema differs, adjust accordingly.
        if 'Fin Year' not in sales_data_df.columns or 'Week' not in sales_data_df.columns:
            st.error("The uploaded data must contain 'Fin Year' and 'Week' columns.")
            st.stop()
        
        sales_data_df['YearWeek'] = sales_data_df['Fin Year'] * 100 + sales_data_df['Week']
        
        # Get the latest completed week (based on max stock units)
        weeksumsales = sales_data_df.groupby('YearWeek')['Actual Current Stock Units'].sum().reset_index()
        max_stock_row = weeksumsales.loc[weeksumsales['Actual Current Stock Units'].idxmax()]
        max_yearweek = max_stock_row['YearWeek']
        lasted_completed_week = int(max_yearweek)

        # Global thresholds
        required_seaonal_clr_percent = 0.9
        required_quartery_clr_percent = 0.8
        required_8wks_clr_percent = 0.7
        long_horizon_diminishing_return = 0.95
        medium_horizon_diminishing_return = 0.98
        min_availability = 40 
        stock_threshold = 1
        confidence = 0.85
        low_data_confidence = 0.7
        Min_no_of_data_points = 4
        data_reduction_threshold = 0.5
        
        latest_completed_year_week = fin_year*100 + lasted_completed_week
        year_week = fin_year*100 + lasted_completed_week
        stock_from_finyear_week = stock_from_finyear*100 + stock_from_week
        
        # Ensure margin is numeric
        if 'Actual Sales Margin %' not in sales_data_df.columns:
            st.error("The uploaded data must contain 'Actual Sales Margin %' column.")
            st.stop()
        
        sales_data_df['Actual Sales Margin %'] = pd.to_numeric(sales_data_df['Actual Sales Margin %'], errors='coerce')
        sales_data_df['YearWeek'] = pd.to_numeric(sales_data_df['YearWeek'], errors='coerce')

        # Calculate average sales in the historical horizon
        average_sales1 = sales_data_df[
            (sales_data_df['YearWeek'] >= historical_horizon_period_start) &
            (sales_data_df['YearWeek'] <= historical_horizon_period_end)
        ].groupby('SKU ID')['Actual Sales Units'].mean().reset_index()
        average_sales1.columns = ['SKU ID', 'Average Sales']

        sales_data_df2 = sales_data_df.merge(average_sales1, on='SKU ID', how='left')

        # Create HRN column
        sales_data_df2['HRN'] = np.where(
            (sales_data_df2['YearWeek'] >= historical_horizon_period_start) &
            (sales_data_df2['YearWeek'] <= historical_horizon_period_end),
            1,
            0
        )

        # OPP column
        sales_data_df2['OPP'] = np.where(
            ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= sales_data_df2['Average Sales'] * stock_threshold) &
            ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= 2) &
            (sales_data_df2['HRN'] == 1),
            1,
            0
        )

        # GP column
        sales_data_df2['GP'] = np.where(
            ((sales_data_df2['Actual Sales Margin %'] >= (min_acceptable_margin * 0.9)) |
            ((sales_data_df2['Actual Sales Margin %'].isna()) & (sales_data_df2['OPP'] == 1))) &
            (sales_data_df2['HRN'] == 1),
            1,
            0
        )

        # OPP_GP column
        sales_data_df2['OPP_GP'] = sales_data_df2['OPP'] + sales_data_df2['GP'] + sales_data_df2['HRN']

        # Calculate averages
        average_sales_units = sales_data_df2[sales_data_df2['OPP_GP'] == 3].groupby('SKU ID')['Actual Sales Units'].mean().reset_index()
        average_sales_units.columns = ['SKU ID', 'Av Sales U when in stock']

        sales_units_reviewed = sales_data_df2[sales_data_df2['OPP_GP'] == 3].groupby('SKU ID')['Actual Sales Units'].sum().reset_index()
        sales_units_reviewed.columns = ['SKU ID', 'Tot Sales U when in stock']

        sales_units_reviewed_count = sales_data_df2[sales_data_df2['OPP_GP'] == 3].groupby('SKU ID')['Actual Sales Units'].count().reset_index()
        sales_units_reviewed_count.columns = ['SKU ID', 'count of reviewed data point']

        sales_units_total = sales_data_df2[
            (sales_data_df2['HRN'] == 1) &
            ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= 2)
        ].groupby('SKU ID')['Actual Sales Units'].sum().reset_index()
        sales_units_total.columns = ['SKU ID', 'Tot Sales U in horizon']

        tot_average_sales_units = sales_data_df2[
            (sales_data_df2['HRN'] == 1) &
            ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= 2)
        ].groupby('SKU ID')['Actual Sales Units'].mean().reset_index()
        tot_average_sales_units.columns = ['SKU ID', 'Tot Ave Sales U in horizon']

        sales_units_total_count = sales_data_df2[
            (sales_data_df2['HRN'] == 1) &
            ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= 2)
        ].groupby('SKU ID')['Actual Sales Units'].count().reset_index()
        sales_units_total_count.columns = ['SKU ID', 'count of horizon data point']

        merged_df = pd.DataFrame({'SKU ID': sales_data_df2['SKU ID'].unique()})
        merged_df = merged_df.merge(sales_units_total, on='SKU ID', how='left')
        merged_df = merged_df.merge(sales_units_total_count, on='SKU ID', how='left')
        merged_df = merged_df.merge(tot_average_sales_units, on='SKU ID', how='left')
        merged_df = merged_df.merge(sales_units_reviewed, on='SKU ID', how='left')
        merged_df = merged_df.merge(average_sales_units, on='SKU ID', how='left')
        merged_df = merged_df.merge(sales_units_reviewed_count, on='SKU ID', how='left')

        # Use_this_ave_sales_u calculation
        merged_df['Use_this_ave_sales_u'] = np.where(
            (merged_df['count of reviewed data point'] <= Min_no_of_data_points) |
            ((merged_df['count of reviewed data point']/merged_df['count of horizon data point']) <= data_reduction_threshold),
            merged_df['Av Sales U when in stock']*low_data_confidence,
            merged_df['Av Sales U when in stock']*confidence
        )

        # Expected intakes
        expected_filtered_df = sales_data_df[
            (sales_data_df['YearWeek'] >= expected_period_start) &
            (sales_data_df['YearWeek'] <= expected_period_end)
        ]
        expected_intakes = expected_filtered_df.groupby('SKU ID')['Expected Intake Units'].sum().reset_index()
        expected_intakes.columns = ['SKU ID', 'Expected Intake Units']

        # Actual intakes in horizon
        actual_intakes = sales_data_df2[sales_data_df2['HRN'] == 1].groupby('SKU ID')['Actual Intake Units'].sum().reset_index()
        actual_intakes.columns = ['SKU ID', 'Actual Intake Units']

        # Current stock
        current_stock1 = sales_data_df[sales_data_df['YearWeek'] == max_yearweek]
        current_stock = current_stock1.groupby('SKU ID')['Actual Current Stock Units'].sum().reset_index()
        current_stock.columns = ['SKU ID', 'Actual Current Stock Units']

        merged_df = pd.merge(merged_df, actual_intakes, on='SKU ID', how='left')
        merged_df['Actual Intake Units'] = merged_df['Actual Intake Units'].fillna(0, inplace=True)
        merged_df = pd.merge(merged_df, expected_intakes, on='SKU ID', how='left')
        merged_df['Expected Intake Units'] = merged_df['Expected Intake Units'].fillna(0, inplace=True)
        merged_df = pd.merge(merged_df, current_stock, on='SKU ID', how='left')
        merged_df['Actual Current Stock Units'] = merged_df['Actual Current Stock Units'].fillna(0, inplace=True)

        merged_df['Total_Season_Sales'] = round(merged_df["Use_this_ave_sales_u"]*26*long_horizon_diminishing_return)
        merged_df['Total_Season_ideal_intakes'] = round(merged_df["Total_Season_Sales"]/required_seaonal_clr_percent)- merged_df['Actual Current Stock Units']
        merged_df['Total_Qtr_Sales'] = round(merged_df["Use_this_ave_sales_u"]*13*medium_horizon_diminishing_return)
        merged_df['Total_Qtr_ideal_intakes'] = round(merged_df["Total_Qtr_Sales"]/required_quartery_clr_percent)- merged_df['Actual Current Stock Units']
        merged_df['Total_9wks_Sales_once_off_repeat'] = round(merged_df["Use_this_ave_sales_u"]*8*medium_horizon_diminishing_return)
        merged_df['Total_9wks_Sales_once_off_repeat_ideal_intakes'] = round(merged_df["Total_9wks_Sales_once_off_repeat"]/required_8wks_clr_percent)- merged_df['Actual Current Stock Units']

        # Additional product info
        # Ensure these columns exist in the uploaded data or adjust as needed
        required_cols = ['Product ID','Current RSP (incl VAT)','Image 1 URL','product_url','Brand','Department','Category Level 1','Category Level 2','Product','Size']
        missing_cols = [c for c in required_cols if c not in sales_data_df.columns]
        for c in missing_cols:
            sales_data_df[c] = ""

        selected_columns_df = sales_data_df[['SKU ID', 'Product ID','Current RSP (incl VAT)','Image 1 URL', 'product_url','Brand','Department','Category Level 1','Category Level 2','Product','Size']].drop_duplicates(subset='SKU ID')

        selected_columns_df2 = selected_columns_df.merge(merged_df, on='SKU ID', how='left')
        selected_columns_df2.fillna(0, inplace=True)

        df_reordered = selected_columns_df2[['Brand', 'Department', 'Category Level 1', 'Category Level 2', 'SKU ID','Product ID','Product','Size','Current RSP (incl VAT)','Actual Intake Units','Actual Current Stock Units','Expected Intake Units','Tot Ave Sales U in horizon', 'count of horizon data point', 'Tot Sales U when in stock', 'Av Sales U when in stock', 'count of reviewed data point', 'Use_this_ave_sales_u', 'Total_Season_Sales', 'Total_Season_ideal_intakes', 'Total_Qtr_Sales', 'Total_Qtr_ideal_intakes', 'Total_9wks_Sales_once_off_repeat', 'Total_9wks_Sales_once_off_repeat_ideal_intakes','Image 1 URL', 'product_url']]

        # --- Original Code Logic Ends Here ---

        # Convert the result to an Excel file for download
        towrite = BytesIO()
        df_reordered.to_excel(towrite, index=False)
        towrite.seek(0)

        st.success("Forecasting complete! Download your results below:")
        st.download_button(
            label="Download Results",
            data=towrite,
            file_name="forecast_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload a file before forecasting.")
