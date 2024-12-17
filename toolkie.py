import streamlit as st
import numpy as np
import pandas as pd
import openpyxl
from io import BytesIO

# Custom CSS
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
    }
    .big-font {
        font-size: 24px !important;
        color: #1E3D59 !important;
        margin-bottom: 30px !important;
    }
    .section-header {
        font-size: 20px !important;
        color: #17A2B8 !important;
        margin-top: 30px !important;
        margin-bottom: 20px !important;
    }
    .upload-section {
        background-color: #F7F9FC;
        padding: 30px;
        border-radius: 10px;
        margin-bottom: 30px;
    }
    .parameter-section {
        background-color: #F7F9FC;
        padding: 30px;
        border-radius: 10px;
    }
    .stButton>button {
        background-color: #17A2B8;
        color: white;
        padding: 10px 30px;
        border-radius: 5px;
        border: none;
        margin-top: 20px;
    }
    .stButton>button:hover {
        background-color: #138496;
    }
</style>
""", unsafe_allow_html=True)

# App Header
st.markdown('<p class="big-font">📊 Forecasting App with Full Calculations</p>', unsafe_allow_html=True)

# File Upload Section
st.markdown('<p class="section-header">📁 Upload Your Raw Data</p>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Upload a single Excel file containing all raw data", 
        type=["xlsx", "xls"],
        help="Please ensure your file contains all required columns"
    )
    st.markdown('</div>', unsafe_allow_html=True)

# Parameters Section
st.markdown('<p class="section-header">🎯 Parameters</p>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="parameter-section">', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        historical_horizon_period_start = st.number_input(
            "Historical Horizon Start",
            value=202440,
            step=1,
            help="Enter Fin_Yr_Wk (eg 202430)"
        )
        
        historical_horizon_period_end = st.number_input(
            "Historical Horizon End",
            value=202533,
            step=1,
            help="Enter Fin_Yr_Wk (eg 202430)"
        )
        
        min_acceptable_margin = st.number_input(
            "Minimum Acceptable Margin",
            value=0.35,
            step=0.01,
            help="Enter as decimal (e.g. 0.29 for 29%)"
        )
        
        expected_period_start = st.number_input(
            "Expected Period Start",
            value=202529,
            step=1,
            help="Enter Fin_Yr_Wk (eg 202430)"
        )
    
    with col2:
        expected_period_end = st.number_input(
            "Expected Period End",
            value=202552,
            step=1,
            help="Enter Fin_Yr_Wk (eg 202430)"
        )
        
        fin_year = st.number_input(
            "Financial Year",
            value=2025,
            step=1,
            help="Enter year (e.g. 2025)"
        )
        
        stock_from_finyear = st.number_input(
            "Stock From Financial Year",
            value=2025,
            step=1,
            help="Enter year (e.g. 2025)"
        )
        
        stock_from_week = st.number_input(
            "Stock From Week",
            value=14,
            step=1,
            help="Enter week number (e.g. 14)"
        )
    
    st.markdown('</div>', unsafe_allow_html=True)

# Forecast Button
st.markdown('<div style="text-align: center; margin-top: 30px;">', unsafe_allow_html=True)
if st.button("Generate Forecast 🚀"):
    if uploaded_file is not None:
        # Show a spinner while processing
        with st.spinner('Generating forecast... Please wait...'):
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
            weeksumsales = []
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
            
            #4 Convert 'Actual Sales Margin %' to numeric
            sales_data_df['Actual Sales Margin %'] = pd.to_numeric(sales_data_df['Actual Sales Margin %'], errors='coerce')
            sales_data_df['YearWeek'] = pd.to_numeric(sales_data_df['YearWeek'], errors='coerce')
            sales_data_dfuse = sales_data_df #saving for later in case I have to rerun
            #sales_data_df.to_parquet('sales_data_df.parquet', index=False)
            # Step 1: Drop rows where 'Fin Year' is "-"
            sales_data_dfere = sales_data_df[sales_data_df['Fin Year'] != "-"] #this does nothing because sales_data_dfere isn't being used anywhere else


            #5 Calculation 1: Average Sales
            #5.1  Convert YearWeek to integers, so we can sort
            sales_data_df['YearWeek'] = pd.to_numeric(sales_data_df['YearWeek'], errors='coerce')
            #5.2 Calculate the average sales for each product ID I use this to determine "enough" stock, where stock is less than this its deemed that the week didn't have enough stock
            average_sales1 = sales_data_df[
                (sales_data_df['YearWeek'] >= historical_horizon_period_start)&
                (sales_data_df['YearWeek'] <= historical_horizon_period_end)
            ].groupby('SKU ID')['Actual Sales Units'].mean().reset_index()
            average_sales1.columns = ['SKU ID', 'Average Sales']
            
            #5.3 Merge the average sales back to the filtered DataFrame
            sales_data_df2 = []
            sales_data_df2 = sales_data_df.merge(average_sales1, on='SKU ID')

            #6 creates a column that limits the data to the historical horizon chosen "HRN" stands for Sales Horizon
            sales_data_df2['HRN'] = np.where(
                (sales_data_df2['YearWeek'] >= historical_horizon_period_start) & 
                (sales_data_df2['YearWeek'] <= historical_horizon_period_end), 
                1, 
                0
            )
            
            #7 creating a column that returns 1 where the stock threshold is met i.e (opening_stock)''OPP'' > (average_sales*stock_threshold) - this will be used to exclude weeks in which we didn't have enough stock
            sales_data_df2['OPP'] = np.where(
                ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= sales_data_df2['Average Sales'] * stock_threshold) & 
                ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= 2)&
                (sales_data_df2['HRN'] == 1), 
                1, 
                0
            )
            
            #8 creates a column that returns 1 where margin% is above the required threshold, this is used to exclude high anomoly weeks with very low margin
            #ths only downside with this approach is that it completly excludes the weeks below, instead of using elasticity to normalize those sales.
            #Actually I think this is a good idea to completly exclude
            sales_data_df2['GP'] = np.where(
                ((sales_data_df2['Actual Sales Margin %'] >= (min_acceptable_margin * 0.9)) | 
                (((sales_data_df2['Actual Sales Margin %'] == "" ) & (sales_data_df2['OPP'] == 1))) &
                (sales_data_df2['HRN'] == 1)), 
                1, 
                0
            )
            
            #9
            #This is to create a column that sumns the three and if its equal to 3 then that week meets all the 3 criteria

            sales_data_df2['OPP_GP'] = sales_data_df2['OPP'] + sales_data_df2['GP'] + sales_data_df2['HRN']
            
            #10 Final Calculations
            #10.1 Average sales where all three conditions are met
            average_sales_units = sales_data_df2[sales_data_df2['OPP_GP'] == 3].groupby('SKU ID')['Actual Sales Units'].mean().reset_index()
            average_sales_units.columns = ['SKU ID', 'Av Sales U when in stock']
            
            #10.2 Sum of the sales where all three policies are met
            sales_units_reviewed = sales_data_df2[sales_data_df2['OPP_GP'] == 3].groupby('SKU ID')['Actual Sales Units'].sum().reset_index()
            sales_units_reviewed.columns = ['SKU ID', 'Tot Sales U when in stock']
            
            #10.3 Count the data points where all three policies are met
            sales_units_reviewed_count = sales_data_df2[sales_data_df2['OPP_GP'] == 3].groupby('SKU ID')['Actual Sales Units'].count().reset_index()
            sales_units_reviewed_count.columns = ['SKU ID', 'count of reviewed data point']
            
            #10.4 Sum of the sales for the entire historical horizon chosen
            sales_units_total = sales_data_df2[(sales_data_df2['HRN'] == 1) & ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= 2)].groupby('SKU ID')['Actual Sales Units'].sum().reset_index()
            sales_units_total.columns = ['SKU ID', 'Tot Sales U in horizon']

            #10.5 average of the sales for the entire historical horizon chosen
            tot_average_sales_units = sales_data_df2[(sales_data_df2['HRN'] == 1) & ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= 2)].groupby('SKU ID')['Actual Sales Units'].mean().reset_index()
            tot_average_sales_units.columns = ['SKU ID', 'Tot Ave Sales U in horizon']

            #10.6 Count the data points for the entire historical horizon chosen
            sales_units_total_count = sales_data_df2[(sales_data_df2['HRN'] == 1) & ((sales_data_df2['Actual EOW Stock Units']+sales_data_df2['Actual Sales Units']) >= 2)].groupby('SKU ID')['Actual Sales Units'].count().reset_index()
            sales_units_total_count.columns = ['SKU ID', 'count of horizon data point']

            #10.7 Merge the DataFrames
            merged_df = []
            selected_columns_df2 = sales_data_df2[['SKU ID']].drop_duplicates().reset_index(drop=True) #this gives me the original sku ids
            merged_df = selected_columns_df2.merge(sales_units_total, on='SKU ID')
            merged_df = merged_df.merge(sales_units_total_count, on='SKU ID')
            merged_df = merged_df.merge(tot_average_sales_units, on='SKU ID')
            merged_df = merged_df.merge(sales_units_reviewed, on='SKU ID')
            merged_df = merged_df.merge(average_sales_units, on='SKU ID')
            merged_df = merged_df.merge(sales_units_reviewed_count, on='SKU ID')

            #10.7 Display the merged DataFrame
            merged_df_x = merged_df #saving in case of an error later on so I reboot from this point

            #11 Derives the base average weekly sales to be used in forecasting
            merged_df['Use_this_ave_sales_u'] = np.where(
                (merged_df['count of reviewed data point'] <= Min_no_of_data_points) | 
                ((merged_df['count of reviewed data point']/merged_df['count of horizon data point']) <= data_reduction_threshold), 
                merged_df['Av Sales U when in stock']*low_data_confidence, 
                merged_df['Av Sales U when in stock']*confidence
            )
            merged_df.head()

            #12 Gets expected intakes and actual and merges them into the final df
            #12.1 Sum of expected Intakes
            expected_filtered_df = sales_data_df[
                (sales_data_df['YearWeek'] >= expected_period_start) &
                (sales_data_df['YearWeek'] <= expected_period_end)
            ]
            expected_intakes = expected_filtered_df.groupby('SKU ID')['Expected Intake Units'].sum().reset_index()
            expected_intakes.columns = ['SKU ID', 'Expected Intake Units']

            #12.2  Sum of actual Intakes within the selected horizon
            actual_intakes = sales_data_df2[(sales_data_df2['HRN'] == 1)].groupby('SKU ID')['Actual Intake Units'].sum().reset_index()
            actual_intakes.columns = ['SKU ID', 'Actual Intake Units']


            #13
            #CurrentStock
            current_stock1 = []
            current_stock = []
            current_stock1 = sales_data_df[sales_data_df['YearWeek']==max_yearweek] #filters on the last week, because in our business current stock is in the current week
            current_stock = current_stock1.groupby('SKU ID')['Actual Current Stock Units'].sum().reset_index()
            current_stock.columns = ['SKU ID', 'Actual Current Stock Units']

            #14
            # Merge the average sales back to the filtered DataFrame
            merged_df = pd.merge(merged_df, actual_intakes, on='SKU ID',how='left')
            merged_df['Actual Intake Units'].fillna(0, inplace=True)
            merged_df = pd.merge(merged_df,expected_intakes, on='SKU ID',how='left') #the limitation here is that for styles that have not gone live they will be excluded
            merged_df['Expected Intake Units'].fillna(0, inplace=True)
            merged_df = pd.merge(merged_df, current_stock, on='SKU ID',how='left')
            merged_df['Actual Current Stock Units'].fillna(0, inplace=True)
            


            #15
            merged_df['Total_Season_Sales'] = round(merged_df["Use_this_ave_sales_u"]*26*long_horizon_diminishing_return)
            merged_df['Total_Season_ideal_intakes'] = round(merged_df["Total_Season_Sales"]/required_seaonal_clr_percent)- merged_df['Actual Current Stock Units']
            merged_df['Total_Qtr_Sales'] = round(merged_df["Use_this_ave_sales_u"]*13*medium_horizon_diminishing_return)
            merged_df['Total_Qtr_ideal_intakes'] = round(merged_df["Total_Qtr_Sales"]/required_quartery_clr_percent)- merged_df['Actual Current Stock Units']
            merged_df['Total_9wks_Sales_once_off_repeat'] = round(merged_df["Use_this_ave_sales_u"]*8*medium_horizon_diminishing_return)
            merged_df['Total_9wks_Sales_once_off_repeat_ideal_intakes'] = round(merged_df["Total_9wks_Sales_once_off_repeat"]/required_8wks_clr_percent)- merged_df['Actual Current Stock Units']
            
            #16
            # Select only the necessary columns from sales_data_df
            selected_columns_df = sales_data_df[['SKU ID', 'Product ID','Current RSP (incl VAT)','Image 1 URL', 'product_url','Brand','Department','Category Level 1','Category Level 2','Product','Size']].drop_duplicates(subset='SKU ID')

            #17
            # Merge with average_sales_final
            selected_columns_df2 = selected_columns_df.merge(merged_df, on='SKU ID')

            # Replace NA with 0
            selected_columns_df2.fillna(0,inplace=True)
            #selected_columns_df2 ['Ideal_Summer_Intake'] = selected_columns_df2['Ideal_Intakes'] - selected_columns_df2['Actual Current Stock Units']

            #18
            df_reordered = selected_columns_df2[['Brand', 'Department', 'Category Level 1', 'Category Level 2', 'SKU ID','Product ID','Product','Size','Current RSP (incl VAT)','Actual Intake Units','Actual Current Stock Units','Expected Intake Units','Tot Ave Sales U in horizon', 'count of horizon data point', 'Tot Sales U when in stock', 'Av Sales U when in stock', 'count of reviewed data point', 'Use_this_ave_sales_u',   'Total_Season_Sales', 'Total_Season_ideal_intakes', 'Total_Qtr_Sales', 'Total_Qtr_ideal_intakes', 'Total_9wks_Sales_once_off_repeat', 'Total_9wks_Sales_once_off_repeat_ideal_intakes','Image 1 URL', 'product_url']]


            # --- Original Code Logic Ends Here ---

            # Convert the result to an Excel file for download
            towrite = BytesIO()
            df_reordered.to_excel(towrite, index=False)
            towrite.seek(0)

            # Success message and download button
            st.success('✅ Forecast generated successfully!')
            st.download_button(
                label="📥 Download Results",
                data=towrite,
                file_name="forecast_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Display preview of results
            st.markdown("### 📊 Preview of Results")
            st.dataframe(df_reordered.head())

    else:
        st.error('⚠️ Please upload a file before generating the forecast.')

st.markdown('</div>', unsafe_allow_html=True)

# Add footer
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    Made with ❤️ by Your Company
</div>
""", unsafe_allow_html=True)            