import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font

st.set_page_config(page_title="HOT Report Extractor", layout="wide")
st.title("HOT Report Extractor")

st.markdown("""
Upload a DB Roberts HOT (Hot Order Tracker) Excel report and get a formatted SALES output with today's date.
""")

uploaded_file = st.file_uploader("Upload HOT Report (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        
        today = datetime.today().date()
        
        cols_needed = ['PART', 'TYPE', 'LOCATION', 'MPE_PO', 'ORDER_NO', 'OP_NO', 
                       'PO_LINE_NO', 'SO_REL_NO', 'SO_SEQ_NO', 'LINE_ITEM_NO', 
                       'LOC_ORDER_NO', 'REQ_QTY']
        
        missing = [c for c in cols_needed if c not in df.columns]
        if missing:
            st.error(f"Missing columns: {', '.join(missing)}")
        else:
            df_output = pd.DataFrame()
            df_output['PART'] = df['PART']
            df_output['TYPE'] = df['TYPE']
            df_output['WC'] = df['LOCATION'].str.replace('DBR-', '', regex=False)
            df_output['CONCAT_KEY'] = (df['MPE_PO'].astype(str) + '|' + 
                                       df['ORDER_NO'].astype(str) + '|' + 
                                       df['OP_NO'].astype(str) + '|' + 
                                       df['PO_LINE_NO'].astype(str) + '|' + 
                                       df['SO_REL_NO'].astype(str) + '|' + 
                                       df['SO_SEQ_NO'].astype(str) + '|' + 
                                       df['PO_LINE_REL'].astype(str) + '|' + 
                                       df['LINE_ITEM_NO'].astype(str))
            df_output['REQ DATE'] = today.strftime('%m/%d/%Y')
            df_output['Total'] = df['REQ_QTY']
            df_output['New PO #'] = df['LOC_ORDER_NO']
            
            st.subheader("Preview")
            st.dataframe(df_output, use_container_width=True)
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_output.to_excel(writer, sheet_name='SALES', index=False)
                workbook = writer.book
                worksheet = writer.sheets['SALES']
                for col in worksheet.columns:
                    worksheet.column_dimensions[col[0].column_letter].width = 20
            
            output.seek(0)
            
            st.download_button(
                label="Download SALES Report",
                data=output.getvalue(),
                file_name=f"SALES_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
