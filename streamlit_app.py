import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO


# Function to load and parse a CSV file
def load_csv(file):
    return pd.read_csv(file)

# Function to convert a DataFrame into an Excel file in memory
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.save()
    processed_data = output.getvalue()
    return processed_data


# Streamlit app
def main():
    # Title of the app
    st.title('Monthly Revenue and Forecast Fee Analysis')

    # File uploaders for the CSV files
    price_list_file = st.file_uploader('Upload Price List CSV', type='csv')
    renewals_file = st.file_uploader('Upload Renewals CSV', type='csv')
    forecast_file = st.file_uploader('Upload Forecast CSV', type='csv')

    if price_list_file is not None and renewals_file is not None and forecast_file is not None:
        # Read the CSV files
        price_list_df = load_csv(price_list_file)
        renewals_df = load_csv(renewals_file)
        forecast_df = load_csv(forecast_file)

        # Process Renewals
        combined_df = renewals_df.merge(price_list_df, on='Licence')
        combined_df['renewal_date'] = pd.to_datetime(combined_df['renewal_date'])
        combined_df['Month-Year'] = combined_df['renewal_date'].dt.strftime('%Y-%m')
        monthly_revenue = combined_df.groupby('Month-Year')['Price'].sum().reset_index()
        monthly_revenue['Cumulative Revenue'] = monthly_revenue['Price'].cumsum()

        #-----------------------------------------------------------------
        # Process Forecast
        # Process Forecast data for revenue
        # Process Forecast
        forecast_df['Close Date'] = pd.to_datetime(forecast_df['Close Date'])
        forecast_df['Month-Year'] = forecast_df['Close Date'].dt.strftime('%Y-%m')
        forecast_df['Probability'] = forecast_df['Probability'].str.rstrip('%').astype('float') / 100
        forecast_df['Estimated Value'] = pd.to_numeric(forecast_df['Estimated Value'], errors='coerce')
        forecast_df['Forecast Fee'] = forecast_df['Estimated Value'] * forecast_df['Probability']
        forecast_fee_by_month = forecast_df.groupby('Month-Year')['Forecast Fee'].sum().reset_index()

        # Plotting the forecast revenue chart
        fig, ax = plt.subplots()
        ax.bar(forecast_fee_by_month['Month-Year'], forecast_fee_by_month['Forecast Fee'], color='purple')
        ax.set_xlabel('Month-Year')
        ax.set_ylabel('Forecast Revenue')
        ax.set_title('Monthly Forecast Revenue')
        plt.xticks(rotation=90)
        st.pyplot(fig)
        st.write(forecast_df)
        #----------------------------------------------------------

        # Combine Renewals and Forecast data
        combined_monthly_data = pd.merge(monthly_revenue, forecast_fee_by_month, on='Month-Year', how='outer').fillna(0)
        combined_monthly_data['Total'] = combined_monthly_data['Price'] + combined_monthly_data['Forecast Fee']
        combined_monthly_data['Cumulative Total'] = combined_monthly_data['Total'].cumsum()

        # Rename columns for clarity
        combined_monthly_data.rename(columns={'Price': 'Renewals', 'Cumulative Revenue': 'Cumulative Renewals'},
                                     inplace=True)

        # Plotting the combined stacked bar chart
        fig, ax = plt.subplots()
        ax.bar(combined_monthly_data['Month-Year'], combined_monthly_data['Renewals'], label='Renewal Revenue')
        ax.bar(combined_monthly_data['Month-Year'], combined_monthly_data['Forecast Fee'],
               bottom=combined_monthly_data['Renewals'], label='Forecast Fee')
        ax.set_xlabel('Month-Year')
        ax.set_ylabel('Revenue')
        ax.set_title('Stacked Monthly Renewal Revenue and Forecast Fee')
        ax.legend()
        plt.xticks(rotation=90)
        st.pyplot(fig)

        # Plotting the cumulative total chart with a horizontal line at 1 million
        fig, ax = plt.subplots()
        ax.plot(combined_monthly_data['Month-Year'], combined_monthly_data['Cumulative Total'], marker='o',
                linestyle='-', color='red')

        # Draw a horizontal line at 1 million
        ax.axhline(y=1000000, color='blue', linestyle='--', label='1 Million Target')

        # Draw vertical lines for each month
        for i in range(len(combined_monthly_data['Month-Year'])):
            ax.axvline(x=i, color='grey', linestyle='--', alpha=0.5)

        # Customize x-axis to show all month-year labels
        ax.set_xticks(range(len(combined_monthly_data['Month-Year'])))
        ax.set_xticklabels(combined_monthly_data['Month-Year'], rotation=90)

        ax.set_xlabel('Month-Year')
        ax.set_ylabel('Cumulative Total Revenue')
        ax.set_title('Cumulative Total Revenue Over Time')

        # Ensure that all x-tick labels are displayed
        fig.tight_layout()

        ax.legend()
        st.pyplot(fig)


        # Display the combined table
        st.write("Combined Monthly Data:")
        st.table(combined_monthly_data)

    #-----------------------------------------------------------------
    #
    #           Counts
    #
    #-----------------------------------------------------------------

      # Process Renewals for counts
        renewals_count_by_month = combined_df.groupby('Month-Year').size().reset_index(name='Renewals Count')

        # Process Forecast for counts
        forecast_count_by_month = forecast_df.groupby('Month-Year').size().reset_index(name='Forecast Count')

        # Combine Renewals and Forecast counts data
        combined_counts_data = pd.merge(renewals_count_by_month, forecast_count_by_month, on='Month-Year', how='outer').fillna(0)

        # Plotting the stacked bar chart for counts
        fig, ax = plt.subplots()
        ax.bar(combined_counts_data['Month-Year'], combined_counts_data['Renewals Count'], label='Licences Renewed')
        ax.bar(combined_counts_data['Month-Year'], combined_counts_data['Forecast Count'], bottom=combined_counts_data['Renewals Count'], label='Forecast Count')
        ax.set_xlabel('Month-Year')
        ax.set_ylabel('Counts')
        ax.set_title('Stacked Bar Chart of Licences Renewed and Forecast Count')
        ax.legend()
        plt.xticks(rotation=45)
        st.pyplot(fig)

        # Display the combined counts table
        st.write("Combined Monthly Counts Data:")
        st.table(combined_counts_data)

        #---------------------------------------------------------------

        st.write("Download the cumulative revenue file as Excel")
        # Generate Excel file for download

        excel_file = to_excel(combined_monthly_data)
        st.download_button(label='Download Excel file',
                           data=excel_file,
                           file_name='combined_monthly_data.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')



if __name__ == '__main__':
    main()
