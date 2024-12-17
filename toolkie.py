import streamlit as st
import numpy as np
import pandas as pd
import openpyxl
from io import BytesIO

import streamlit as st

# Set page config first
st.set_page_config(
    page_title="Forecasting Toolkie",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern styling
st.markdown("""
    <style>
    /* Main container styling */
    .main {
        padding: 2em;
    }
    
    /* Header styling */
    .header {
        color: #1E88E5;
        font-size: 28px;
        font-weight: 600;
        margin-bottom: 30px;
        padding: 20px 0;
        border-bottom: 2px solid #f0f2f6;
    }
    
    /* Section styling */
    .section-header {
        color: #333;
        font-size: 18px;
        font-weight: 500;
        margin: 20px 0;
        padding: 10px 0;
    }
    
    /* Input container styling */
    .stNumberInput {
        margin: 15px 0;
    }
    
    .stNumberInput > div > div > input {
        border-radius: 5px;
        border: 1px solid #ddd;
        padding: 10px;
    }
    
    /* File uploader styling */
    .uploadedFile {
        border: 2px dashed #1E88E5;
        border-radius: 10px;
        padding: 20px;
        background: #f8f9fa;
    }
    
    /* Button styling */
    .stButton > button {
        width: 100%;
        background-color: #1E88E5;
        color: white;
        padding: 12px 20px;
        border-radius: 5px;
        border: none;
        font-weight: 500;
        transition: all 0.3s;
    }
    
    .stButton > button:hover {
        background-color: #1976D2;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding: 0 24px;
        font-size: 16px;
    }
    </style>
""", unsafe_allow_html=True)

# Layout with tabs
tab1, tab2 = st.tabs(["📤 Upload & Parameters", "📊 Results"])

with tab1:
    # Create two columns
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("### Upload Data")
        uploaded_file = st.file_uploader(
            "Upload Excel file",
            type=["xlsx", "xls"],
            help="Maximum file size: 200MB"
        )
    
    with col2:
        st.markdown("### Forecast Parameters")
        
        with st.container():
            historical_horizon_period_start = st.number_input(
                "Historical Start (YYYYWW)",
                value=202440,
                step=1,
                help="Example: 202430"
            )
            
            historical_horizon_period_end = st.number_input(
                "Historical End (YYYYWW)",
                value=202533,
                step=1
            )
            
            min_acceptable_margin = st.number_input(
                "Minimum Margin",
                value=0.35,
                step=0.01,
                format="%.2f"
            )
            
            expected_period_start = st.number_input(
                "Expected Start (YYYYWW)",
                value=202529,
                step=1
            )
            
            expected_period_end = st.number_input(
                "Expected End (YYYYWW)",
                value=202552,
                step=1
            )

            fin_year = st.number_input(
                "Financial Year",
                value=2025,
                step=1
            )

    # Center the forecast button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.button("Generate Forecast 🚀")

with tab2:
    if 'forecast_generated' not in st.session_state:
        st.info("Generate a forecast to see results here")
    else:
        # Your results display code here
        pass

# Add the rest of your calculation logic here 
#   
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
            # First, get the sum of Total_Season_Sales by Product ID
            #season_sales_sum = df_reordered.groupby('Product ID')['Total_Season_Sales'].sum().reset_index()
            # Get unique product information with max RSP
           # product_info = df_reordered[['Brand', 'Department', 'Category Level 1', 'Category Level 2', 
            #                            'Product ID', 'Product', 'Current RSP (incl VAT)','Image 1 URL', 'product_url']].copy()

            # Get max RSP for each Product ID
            #product_info['Max Current RSP (incl VAT)'] = product_info.groupby('Product ID')['Current RSP (incl VAT)'].transform('max')

            # Drop duplicates to get unique Product IDs with their info
            #product_info = product_info.drop('Current RSP (incl VAT)', axis=1).drop_duplicates(subset='Product ID')

            # Merge the summary with product info
            #summary_df = product_info.merge(season_sales_sum, on='Product ID', how='left')

            # Sort by Total_Season_Sales in descending order
            #summary_df = summary_df.sort_values('Total_Season_Sales', ascending=False)

            # Add this after your calculations but before displaying the results
            def render_images(df):
                return df.to_html(escape=False, formatters=dict( **{
                    'Image 1 URL': lambda x: f'<img src="{x}" width="100" />' if pd.notnull(x) else ''
                }))
                
                # Add custom CSS to make the table scrollable and style it

           
            # --- Original Code Logic Ends Here ---
            # Before creating the Excel file for download:
            #st.markdown("### Preview of Results with Images")
            #html = render_dataframe_with_images(summary_df) I would have loved to have rendered the preview table
            #st.markdown(html, unsafe_allow_html=True)
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

    else:
        st.error('⚠️ Please upload a file before generating the forecast.')

st.markdown('</div>', unsafe_allow_html=True)

# Add footer
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    Developed with ❤️ by Flint Ntshangase
</div>
""", unsafe_allow_html=True)            