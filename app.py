import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import warnings
import xmlrpc.client
from datetime import date, timedelta
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

warnings.filterwarnings('ignore')

# Set page configuration
st.set_page_config(
    page_title="Inventory & ODOO Merger",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #2ca02c;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #d4edda;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #28a745;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #17a2b8;
        margin: 1rem 0;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 0.5rem 1rem;
        border-radius: 0.3rem;
        font-size: 1rem;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #dee2e6;
        text-align: center;
    }
    .tab-container {
        padding: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Odoo configuration from environment variables
def get_odoo_config():
    return {
        'url': os.getenv('ODOO_URL', ''),
        'db': os.getenv('ODOO_DB', ''),
        'username': os.getenv('ODOO_USERNAME', ''),
        'password': os.getenv('ODOO_PASSWORD', ''),
        'hq_company_name': os.getenv('ODOO_COMPANY_NAME', '')
    }

def test_odoo_connection(config):
    try:
        common = xmlrpc.client.ServerProxy(config['url'] + 'xmlrpc/2/common')
        uid = common.authenticate(config['db'], config['username'], config['password'], {})
        if uid:
            return True, "✅ Connected to Odoo successfully!"
        else:
            return False, "❌ Authentication failed"
    except Exception as e:
        return False, f"❌ Connection error: {str(e)}"

def process_files(inventory_file, odoo_file):
    try:
        # Read inventory report - Processed Returns sheet
        inventory_df = pd.read_excel(inventory_file, sheet_name='Processed Returns')
        
        # Read ODOO PO results
        odoo_df = pd.read_excel(odoo_file, sheet_name='PO_Results')
        
        # Process inventory data - KEEP ALL RECORDS (no duplicate removal)
        inventory_processed = inventory_df[['lot', 'product_name', 'vendor', 'cost_price']].copy()
        inventory_processed['vendor_name'] = inventory_processed['vendor']
        inventory_processed['product_name'] = inventory_processed['product_name']
        inventory_processed['unit_price'] = inventory_processed['cost_price']
        inventory_processed['label'] = inventory_processed['lot'].astype(str)
        inventory_processed['source'] = 'inventory'
        inventory_processed['quantity'] = 1  # Set quantity to 1 for all records
        
        # Process ODOO data - KEEP ALL RECORDS (no duplicate removal)
        odoo_processed = odoo_df[['barcode', 'product_name', 'product_ref', 'vendor_name', 'Unit_Price']].copy()
        odoo_processed['label'] = odoo_processed['barcode'].astype(str)
        odoo_processed['unit_price'] = odoo_processed['Unit_Price']
        
        # Combine product_name and product_ref for ODOO data
        odoo_processed['product_name'] = odoo_processed['product_name'] + ' ' + odoo_processed['product_ref'].astype(str)
        odoo_processed['source'] = 'odoo'
        odoo_processed['quantity'] = 1  # Set quantity to 1 for all records
        
        # Select and rename columns for final output
        inventory_final = inventory_processed[['vendor_name', 'product_name', 'unit_price', 'label', 'quantity', 'source']]
        odoo_final = odoo_processed[['vendor_name', 'product_name', 'unit_price', 'label', 'quantity', 'source']]
        
        # Combine both datasets - ALL RECORDS INCLUDED
        final_df = pd.concat([inventory_final, odoo_final], ignore_index=True)
        
        return final_df
        
    except Exception as e:
        raise Exception(f"Error processing files: {str(e)}")

def process_odoo_integration(uploaded_file, config):
    try:
        # Read the uploaded Excel file
        df = pd.read_excel(uploaded_file)
        
        # Check if the required columns exist
        required_columns = ["vendor_name", "product_name", "unit_price", "quantity", "label"]
        for col in required_columns:
            if col not in df.columns:
                raise Exception(f"Missing required column: {col}")
        
        # Rename columns to match expected format
        df = df.rename(columns={
            "vendor_name": "Vendor",
            "product_name": "Product",
            "unit_price": "CostPrice",
            "quantity": "Quantity",
            "label": "LotNumber"
        })
        
        # === Odoo XML-RPC Connection ===
        common = xmlrpc.client.ServerProxy(config['url'] + 'xmlrpc/2/common')
        uid = common.authenticate(config['db'], config['username'], config['password'], {})
        models = xmlrpc.client.ServerProxy(config['url'] + 'xmlrpc/2/object')
        
        # === Fetch Company ID ===
        company_ids = models.execute_kw(config['db'], uid, config['password'],
            'res.company', 'search',
            [[['name', '=', config['hq_company_name']]]],
            {'limit': 1})
        if not company_ids:
            raise Exception(f"Company '{config['hq_company_name']}' not found.")
        company_id = company_ids[0]
        
        # === Static Values ===
        credit_note_date = date.today().strftime("%Y-%m-%d")
        due_date = (date.today() + timedelta(days=1)).strftime("%Y-%m-%d")
        reference = "Damage"
        
        # === Group rows by Vendor, Product, CostPrice ===
        df_grouped = df.groupby(["Vendor", "Product", "CostPrice"], as_index=False).agg({
            "Quantity": "sum",
            "LotNumber": lambda x: ", ".join(str(v) for v in x if pd.notna(v))
        })
        
        results = []
        
        # === Process Vendor Groups ===
        for vendor_name, group in df_grouped.groupby("Vendor"):
            # === Fetch Vendor ID ===
            vendor_ids = models.execute_kw(config['db'], uid, config['password'],
                'res.partner', 'search',
                [[['name', '=', vendor_name], '|', ['company_id', '=', company_id], ['company_id', '=', False]]],
                {'limit': 1})
            if not vendor_ids:
                results.append(f"❌ Vendor '{vendor_name}' not found. Skipping vendor.")
                continue
            vendor_id = vendor_ids[0]
            
            # === Fetch Journal ID ===
            journal_ids = models.execute_kw(config['db'], uid, config['password'],
                'account.journal', 'search',
                [[['type', '=', 'purchase'], ['name', 'ilike', 'Vendor Bills'], ['company_id', '=', company_id]]],
                {'limit': 1})
            if not journal_ids:
                results.append("❌ 'Vendor Bills' journal not found for specified company.")
                continue
            journal_id = journal_ids[0]
            
            # === Build Line Items ===
            line_vals = []
            for _, row in group.iterrows():
                product_name = str(row["Product"])
                qty = float(row["Quantity"])
                price = float(row["CostPrice"])
                lot_number = str(row["LotNumber"])
                
                # === Fetch Product ID ===
                product_ids = models.execute_kw(config['db'], uid, config['password'],
                    'product.product', 'search',
                    [[['name', 'ilike', product_name], '|', ['company_id', '=', company_id], ['company_id', '=', False]]],
                    {'limit': 1})
                
                if not product_ids:
                    results.append(f"❌ Product '{product_name}' not found. Skipping this line.")
                    continue
                product_id = product_ids[0]
                
                # === Create Line Value ===
                line_vals.append((0, 0, {
                    'product_id': product_id,
                    'quantity': qty,
                    'price_unit': price,
                    'name': f"{product_name} (Lots: {lot_number})",
                }))
            
            if not line_vals:
                results.append(f"❌ No valid lines for vendor '{vendor_name}'. Skipping.")
                continue
            
            # === Create Vendor Credit Note ===
            credit_note_id = models.execute_kw(config['db'], uid, config['password'],
                'account.move', 'create',
                [{
                    'move_type': 'in_refund',
                    'partner_id': vendor_id,
                    'invoice_date': credit_note_date,
                    'invoice_date_due': due_date,
                    'journal_id': journal_id,
                    'ref': reference,
                    'invoice_line_ids': line_vals,
                    'company_id': company_id,
                }]
            )
            
            results.append(f"✅ Vendor Credit Note created for '{vendor_name}' with ID: {credit_note_id}")
        
        return results
        
    except Exception as e:
        raise Exception(f"Error processing Odoo integration: {str(e)}")

def main():
    # Header
    st.markdown('<h1 class="main-header">📊 Inventory & ODOO Merger</h1>', unsafe_allow_html=True)
    
    # Create tabs
    tab1, tab2 = st.tabs(["📁 File Merger", "🔄 Odoo Integration"])
    
    with tab1:
        st.markdown("### Combine Inventory Reports with ODOO Purchase Order Data")
        
        # Sidebar
        with st.sidebar:
            st.header("📋 Instructions")
            st.info("""
            1. Upload Inventory Report Excel
            2. Upload ODOO PO Results Excel
            3. Click 'Process Files' to generate combined report
            4. Download the results as Excel file
            """)
        
        # File upload section
        st.markdown('<div class="sub-header">📤 Upload Files</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Inventory Report**")
            inventory_file = st.file_uploader(
                "Upload Excel file with 'Processed Returns' sheet", 
                type=['xlsx'],
                key='inventory'
            )
            if inventory_file:
                try:
                    inventory_preview = pd.read_excel(inventory_file, sheet_name='Processed Returns')
                    st.success(f"✅ Inventory file uploaded! Records: {len(inventory_preview):,}")
                except Exception as e:
                    st.error(f"❌ Error reading inventory file: {str(e)}")
        
        with col2:
            st.markdown("**ODOO PO Results**")
            odoo_file = st.file_uploader(
                "Upload Excel file with 'PO_Results' sheet", 
                type=['xlsx'],
                key='odoo'
            )
            if odoo_file:
                try:
                    odoo_preview = pd.read_excel(odoo_file, sheet_name='PO_Results')
                    st.success(f"✅ ODOO file uploaded! Records: {len(odoo_preview):,}")
                except Exception as e:
                    st.error(f"❌ Error reading ODOO file: {str(e)}")
        
        # Process button
        if inventory_file and odoo_file:
            if st.button("🚀 Process Files", use_container_width=True, key="process_files"):
                with st.spinner("🔄 Processing files... Please wait."):
                    try:
                        # Process files
                        result_df = process_files(inventory_file, odoo_file)
                        
                        # Success message
                        st.markdown('<div class="success-box">✅ Files processed successfully!</div>', unsafe_allow_html=True)
                        
                        # Display results
                        st.markdown('<div class="sub-header">📊 Processed Data</div>', unsafe_allow_html=True)
                        
                        # Show data with pagination
                        st.dataframe(
                            result_df,
                            use_container_width=True,
                            hide_index=True,
                            height=400
                        )
                        
                        # Summary statistics
                        st.markdown('<div class="sub-header">📈 Summary Statistics</div>', unsafe_allow_html=True)
                        
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                            st.metric("Total Records", f"{len(result_df):,}")
                            st.markdown('</div>', unsafe_allow_html=True)
                        
                        with col2:
                            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                            st.metric("From Inventory", f"{len(result_df[result_df['source'] == 'inventory']):,}")
                            st.markdown('</div>', unsafe_allow_html=True)
                        
                        with col3:
                            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                            st.metric("From ODOO", f"{len(result_df[result_df['source'] == 'odoo']):,}")
                            st.markdown('</div>', unsafe_allow_html=True)
                        
                        with col4:
                            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                            total_value = (result_df['unit_price'] * result_df['quantity']).sum()
                            st.metric("Total Value", f"₹{total_value:,.2f}")
                            st.markdown('</div>', unsafe_allow_html=True)
                        
                        # Download section
                        st.markdown('<div class="sub-header">💾 Download Results</div>', unsafe_allow_html=True)
                        
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='Combined_Results')
                        
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Download Combined Excel File",
                            data=output,
                            file_name="credit_note_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="download_combined"
                        )
                            
                    except Exception as e:
                        st.error(f"❌ Error processing files: {str(e)}")
        
        else:
            st.info("📝 Please upload both files to begin processing.")
    
    with tab2:
        st.markdown("### Odoo Integration - Create Vendor Credit Notes")
        
        with st.sidebar:
            st.header("📋 Instructions")
            st.info("""
            1. Connect to Odoo using the button below
            2. Upload the combined Excel file
            3. Click 'Process Odoo Integration'
            4. Review the results
            """)
            
            # Test Odoo connection
            if st.button("🔌 Connect to Odoo", use_container_width=True, key="connect_odoo"):
                config = get_odoo_config()
                if not all(config.values()):
                    st.error("❌ Odoo configuration not found in environment variables")
                else:
                    with st.spinner("Connecting to Odoo..."):
                        success, message = test_odoo_connection(config)
                        if success:
                            st.success(message)
                            st.session_state.odoo_connected = True
                            st.session_state.odoo_config = config
                        else:
                            st.error(message)
                            st.session_state.odoo_connected = False
            
            if st.session_state.get('odoo_connected', False):
                st.success("✅ Odoo is connected")
        
        st.markdown('<div class="sub-header">📤 Upload Combined File</div>', unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "Upload the combined Excel file", 
            type=['xlsx'],
            key='combined_file'
        )
        
        if uploaded_file:
            try:
                preview_df = pd.read_excel(uploaded_file)
                st.success(f"✅ Combined file uploaded! Records: {len(preview_df):,}")
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)}")
        
        if st.button("🚀 Process Odoo Integration", use_container_width=True, key="process_odoo"):
            if not st.session_state.get('odoo_connected', False):
                st.error("❌ Please connect to Odoo first")
            elif not uploaded_file:
                st.error("❌ Please upload the combined Excel file")
            else:
                with st.spinner("🔄 Processing Odoo integration... This may take a while."):
                    try:
                        config = st.session_state.odoo_config
                        results = process_odoo_integration(uploaded_file, config)
                        
                        # Display results
                        st.markdown('<div class="sub-header">📋 Processing Results</div>', unsafe_allow_html=True)
                        
                        for result in results:
                            if result.startswith("✅"):
                                st.success(result)
                            elif result.startswith("❌"):
                                st.error(result)
                            else:
                                st.info(result)
                        
                        st.markdown('<div class="success-box">✅ Odoo integration completed!</div>', unsafe_allow_html=True)
                        
                    except Exception as e:
                        st.error(f"❌ Error processing Odoo integration: {str(e)}")
    
    # Footer
    st.markdown("---")
    st.markdown("**Note:** All quantities are set to 1 as per requirements. All records from both files are included without duplicate removal.")

if __name__ == "__main__":
    main()
