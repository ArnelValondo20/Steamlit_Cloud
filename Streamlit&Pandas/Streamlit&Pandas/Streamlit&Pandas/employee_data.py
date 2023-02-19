import pandas as pd
import streamlit as st

# Load the data from the CSV file
data = pd.read_csv(r'C:\Users\User\Desktop\NEL_CODES\Streamlit&Pandas\archive\HRDataset_v14.csv')

# Create the Streamlit app
st.title('AUTOCAD USAGE DASHBOARD')

# Create date sliders to select the start and end dates for filtering the data
start_date = st.slider('Select a start date:', 
                       min_value=pd.to_datetime(data['Date'], format='%m/%d/%Y').min().date(),
                       max_value=pd.to_datetime(data['Date'], format='%m/%d/%Y').max().date(),
                       step=pd.Timedelta(days=1))
end_date = st.slider('Select an end date:', 
                     min_value=pd.to_datetime(data['Date'], format='%m/%d/%Y').min().date(),
                     max_value=pd.to_datetime(data['Date'], format='%m/%d/%Y').max().date(),
                     step=pd.Timedelta(days=1))

# Filter the data by the selected date range
filtered_data = data[(pd.to_datetime(data['Date'], format='%m/%d/%Y').dt.date >= start_date) & 
                     (pd.to_datetime(data['Date'], format='%m/%d/%Y').dt.date <= end_date)]

# Create a dropdown to select the department for displaying the employee data
departments = sorted(filtered_data['Department'].unique())
department = st.selectbox('Select a department:', departments)

# Filter the data by the selected department
department_data = filtered_data[filtered_data['Department'] == department]

# Display the employee data for the selected department
if len(department_data) > 0:
    st.write(f"Employee data for {department} department between {start_date} and {end_date}:")
    st.write(department_data)
    
    # Add a button to export the report to Excel
    if st.button('Export report to Excel'):
        file_name = f"{department}_department_{start_date}_to_{end_date}.xlsx"
        department_data.to_excel(file_name, index=False)
        st.write(f"Report has been exported to {file_name}")
else:
    st.write(f"No employee data found for {department} department between {start_date} and {end_date}.")
