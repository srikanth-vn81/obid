import streamlit as st
import pandas as pd
from io import BytesIO

# Function to process the data
def process_data(ob_data, spl_data, style_data):
    # Check if 'VPO No' column exists before transforming it into 'PO'
    if 'VPO No' in ob_data.columns:
        def transform_vpo_no(vpo_no):
            if isinstance(vpo_no, str):
                if vpo_no.startswith('8'):
                    return vpo_no[:8]
                elif vpo_no.startswith('D'):
                    return 'P' + vpo_no[1:-3]
            return vpo_no
        
        ob_data['PO'] = ob_data['VPO No'].apply(transform_vpo_no)
        ob_data['PO'] = ob_data['PO'].astype(str)
    
    # Ensure that 'Production Plan ID' column exists, and if not, create it with default value 0
    if 'Production Plan ID' not in ob_data.columns:
        ob_data['Production Plan ID'] = 0

    def update_production_plan_id(row):
        if pd.isna(row['Production Plan ID']) or row['Production Plan ID'] == 0:
            if row['PO'].startswith('8'):
                return row['PO']
            elif row['Season'][-2:] == '23':
                return 'Season-23'
        return row['Production Plan ID']

    ob_data['Production Plan ID'] = ob_data.apply(update_production_plan_id, axis=1)

    # Filter the DataFrame for rows where 'Group Tech Class' equals 'BELUNIQLO'
    filtered_data = ob_data[ob_data['Group Tech Class'] == 'BELUNIQLO']

    # SPL Processing
    filtered_data['PO'] = filtered_data['PO'].astype(str).str.strip()
    spl_data['PO Order NO'] = spl_data['PO Order NO'].astype(str).str.strip()

    # Clean up any Unnamed columns in both datasets
    filtered_data = filtered_data.loc[:, ~filtered_data.columns.str.contains('^Unnamed')]
    spl_data = spl_data.loc[:, ~spl_data.columns.str.contains('^Unnamed')]

    # Create a lookup dictionary for 'Production Plan ID' using the SPL data
    lookup_dict = spl_data.set_index('PO Order NO')['Production Plan ID'].to_dict()

    # Update 'Production Plan ID' using the SPL data
    filtered_data['Production Plan ID'] = filtered_data['PO'].map(lookup_dict)

    # Fill missing 'Production Plan ID' with 'PO' if applicable
    filtered_data['Production Plan ID'] = filtered_data.apply(
        lambda row: row['PO'] if pd.isna(row['Production Plan ID']) and row['PO'].startswith('8') else row['Production Plan ID'],
        axis=1
    )

    # Filter rows where 'CO Qty' is non-negative
    filtered_data = filtered_data[filtered_data['CO Qty'] >= 0]

    # Style Product Mapping
    style_lookup_dict = style_data.set_index('Style')['Master Item'].to_dict()
    filtered_data['Style'] = filtered_data['Cust Style No'].apply(lambda x: x[2:10] if isinstance(x, str) else x)
    filtered_data['Product'] = filtered_data['Style'].apply(lambda x: style_lookup_dict.get(x, None))

    return filtered_data

# Streamlit App UI
st.title("Order Book Processing App")

# Sidebar for file uploads and actions
st.sidebar.title("Options")

# File uploaders in the sidebar
ob_file = st.sidebar.file_uploader("Upload the OB Excel file", type="xlsx")
spl_file = st.sidebar.file_uploader("Upload the SPL CSV file", type="csv")
style_file = st.sidebar.file_uploader("Upload the Style Product Mapping Summary Excel file", type="xlsx")

# Run button in sidebar
if st.sidebar.button("Run") and ob_file and spl_file and style_file:
    # Load the files
    ob_data = pd.read_excel(ob_file, sheet_name='Sheet1')
    spl_data = pd.read_csv(spl_file)
    style_data = pd.read_excel(style_file)

    # Process the data
    final_data = process_data(ob_data, spl_data, style_data)

    # Display the final processed data
    st.subheader("Processed Data")
    st.write(final_data)

    # Prepare the data for download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_data.to_excel(writer, index=False)
    output.seek(0)

    # Download button
    st.download_button(
        label="Download Processed Data",
        data=output,
        file_name='pid_final.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.write("Please upload all required files and click 'Run'.")

# Custom CSS for UI styling
st.markdown("""
    <style>
        .stApp {
            background-color: #f4f4f4;
            font-family: 'Arial', sans-serif;
        }
        .sidebar .sidebar-content {
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }
        .stButton>button {
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 5px;
            padding: 10px 20px;
            cursor: pointer;
        }
        .stButton>button:hover {
            background-color: #0056b3;
        }
    </style>
""", unsafe_allow_html=True)
