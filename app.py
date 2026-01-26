import streamlit as st
import pandas as pd
from io import BytesIO
import time

st.set_page_config(page_title="RND Stok Automation", layout="wide")

# UI Styling
st.markdown("""
    <style>
        .main { background-color: #f8f9fa; }
        div.stButton > button {
            background-color: #2e7d32;
            color: white;
            width: 100%;
            border-radius: 10px;
            height: 50px;
            font-weight: bold;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📦 RND Stok Report")

# File Uploaders
c1, c2 = st.columns(2)
with c1:
    up_barcode = st.file_uploader("📂 1. Barkodlu Ürün Raporu", type=["xlsx"])
    up_sales = st.file_uploader("📂 2. Net Satış Raporu", type=["xlsx"])
with c2:
    up_orders = st.file_uploader("📂 3. Orders In Excel", type=["xlsx"])
    up_template = st.file_uploader("📂 4. Template File", type=["xlsx"])

def get_col_val(df, letter):
    """Safely get column by Excel letter (A=0, B=1, etc.)"""
    idx = ord(letter.upper()) - 65
    if idx < len(df.columns):
        return df.iloc[:, idx]
    return pd.Series([0] * len(df))

if all([up_barcode, up_sales, up_orders, up_template]):
    if st.button("RUN AUTOMATION"):
        try:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("Reading Excel files...")
            df_bc = pd.read_excel(up_barcode)
            df_sl = pd.read_excel(up_sales)
            df_ord = pd.read_excel(up_orders)
            df_temp_cols = pd.read_excel(up_template).columns
            progress_bar.progress(20)

            def clean_key(series):
                return series.astype(str).str.strip().str.upper()

            # --- Data Normalization & Lookups ---
            df_bc_clean = df_bc.copy()
            bc_key = clean_key(get_col_val(df_bc_clean, 'D'))
            df_bc_clean['lookup_key'] = bc_key
            bc_lk = df_bc_clean.drop_duplicates(subset=['lookup_key']).set_index('lookup_key')

            df_sl_clean = df_sl.copy()
            sl_key = clean_key(get_col_val(df_sl_clean, 'A'))
            df_sl_clean['lookup_key'] = sl_key
            sl_lk = df_sl_clean.drop_duplicates(subset=['lookup_key']).set_index('lookup_key')
            progress_bar.progress(40)

            # --- Sheet 1: Satılmış Ürün Raporu ---
            s1 = pd.DataFrame(columns=df_temp_cols)
            s1['Customer Invoice'] = get_col_val(df_ord, 'A')
            s1['Guess Code'] = get_col_val(df_ord, 'M')
            s1['Title'] = get_col_val(df_ord, 'I')
            s1['Season'] = get_col_val(df_ord, 'D')
            s1['Model'] = get_col_val(df_ord, 'G')
            s1['Part'] = get_col_val(df_ord, 'H')
            s1['N Order'] = get_col_val(df_ord, 'J')
            s1['Color By Style'] = get_col_val(df_ord, 'L')
            s1['V Net Order'] = get_col_val(df_ord, 'N')
            s1['Q Order'] = get_col_val(df_ord, 'O')

            s1_keys = clean_key(s1['Guess Code'])
            s1['UrunTipi'] = s1_keys.map(bc_lk.iloc[:, 12]) 
            s1['Line'] = s1_keys.map(bc_lk.iloc[:, 13])     
            s1['Stock'] = s1_keys.map(bc_lk.iloc[:, 22])    
            s1['Cinsiyet'] = s1_keys.map(bc_lk.iloc[:, 17]) 
            s1['Ürün Grubu'] = s1_keys.map(bc_lk.iloc[:, 23]) 
            s1['SaleCount'] = s1_keys.map(sl_lk.iloc[:, 9])      
            s1['EcomSaleCount'] = s1_keys.map(sl_lk.iloc[:, 10]) 
            s1['MPSaleCount'] = s1_keys.map(sl_lk.iloc[:, 11])   

            # --- Sheet 2: Tüm Ürün Raporu ---
            s2 = pd.DataFrame(columns=df_temp_cols)
            s2['Guess Code'] = get_col_val(df_bc, 'D')
            s2['Title'] = get_col_val(df_bc, 'A')
            s2['UrunTipi'] = get_col_val(df_bc, 'M')
            s2['Line'] = get_col_val(df_bc, 'N')
            s2['Stock'] = get_col_val(df_bc, 'W')
            s2['Cinsiyet'] = get_col_val(df_bc, 'R')
            s2['Ürün Grubu'] = get_col_val(df_bc, 'X')
            s2['Season'] = get_col_val(df_bc, 'O')
            s2['Color By Style'] = get_col_val(df_bc, 'J')

            s2_keys = clean_key(s2['Guess Code'])
            s2['SaleCount'] = s2_keys.map(sl_lk.iloc[:, 9])
            s2['EcomSaleCount'] = s2_keys.map(sl_lk.iloc[:, 10])
            s2['MPSaleCount'] = s2_keys.map(sl_lk.iloc[:, 11])

            df_ord_clean = df_ord.copy()
            df_ord_clean['clean_m'] = clean_key(get_col_val(df_ord, 'M'))
            ord_agg = df_ord_clean.groupby('clean_m').agg({
                df_ord.columns[9]: 'sum',  
                df_ord.columns[13]: 'sum', 
                df_ord.columns[14]: 'sum'  
            })
            s2['N Order'] = s2_keys.map(ord_agg.iloc[:, 0])
            s2['V Net Order'] = s2_keys.map(ord_agg.iloc[:, 1])
            s2['Q Order'] = s2_keys.map(ord_agg.iloc[:, 2])

            # --- Prepare Full Breakdown Pivot ---
            # Now including Cinsiyet in the grouping
            combined_data = pd.concat([s1, s2], ignore_index=True).drop_duplicates(subset=['Guess Code', 'UrunTipi', 'Line', 'Cinsiyet'])
            numeric_cols = ['Q Order', 'SaleCount', 'MPSaleCount', 'EcomSaleCount', 'Stock']
            for col in numeric_cols:
                combined_data[col] = pd.to_numeric(combined_data[col], errors='coerce').fillna(0)
            
            # Pivot grouping by UrunTipi, Line, AND Cinsiyet
            final_pivot = combined_data.groupby(['UrunTipi', 'Line', 'Cinsiyet', 'Guess Code'])[numeric_cols].sum().reset_index()

            progress_bar.progress(80)

            # --- Excel Workbook Creation ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                s1.to_excel(writer, sheet_name='Satılmış Ürün Raporu', index=False)
                s2.to_excel(writer, sheet_name='Tüm Ürün Raporu', index=False)
                
                ws_ozet = writer.book.add_worksheet('Özet')
                writer.sheets['Özet'] = ws_ozet
                
                # Format Library
                fmt_h1 = writer.book.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
                fmt_h2 = writer.book.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                fmt_lbl = writer.book.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
                fmt_num = writer.book.add_format({'num_format': '#,##0', 'border': 1})
                fmt_pct = writer.book.add_format({'num_format': '0.00%', 'border': 1})
                fmt_inp = writer.book.add_format({'bg_color': '#FFF2CC', 'border': 2, 'bold': True})

                # --- TABLE 1: TOTAL SALE / Q ORDER (STATIC) ---
                ws_ozet.write('A1', 'Full Ratio', fmt_h1)
                ws_ozet.write('A2', 'Metric', fmt_h2)
                ws_ozet.write('B2', 'Total Value', fmt_h2)
                
                total_sc = pd.to_numeric(s2['SaleCount'], errors='coerce').sum()
                total_qo = pd.to_numeric(s2['Q Order'], errors='coerce').sum()
                
                ws_ozet.write('A3', 'Total Sale Count (Col M)')
                ws_ozet.write('B3', total_sc, fmt_num)
                ws_ozet.write('A4', 'Total Q Order (Col L)')
                ws_ozet.write('B4', total_qo, fmt_num)
                ws_ozet.write('A5', 'Efficiency Ratio %')
                ws_ozet.write_formula('B5', '=IF(B4=0, 0, B3/B4)', fmt_pct)

                # --- TABLE 2: DYNAMIC BY LINE (FILTER B12) ---
                ws_ozet.write('D1', 'Line Based Ratio', fmt_h1)
                ws_ozet.write('D2', 'Metric', fmt_h2)
                ws_ozet.write('E2', 'Dynamic Value', fmt_h2)

                # Formula logic: Col M (SaleCount), Col L (Q Order), Col Q (Line) in sheet 2
                f_dyn_sale = "=SUMIFS('Tüm Ürün Raporu'!M:M, 'Tüm Ürün Raporu'!Q:Q, B12)"
                f_dyn_qord = "=SUMIFS('Tüm Ürün Raporu'!L:L, 'Tüm Ürün Raporu'!Q:Q, B12)"
                f_dyn_ratio = "=IF(E4=0, 0, E3/E4)"

                ws_ozet.write('D3', 'Line Sale Count')
                ws_ozet.write_formula('E3', f_dyn_sale, fmt_num)
                ws_ozet.write('D4', 'Line Q Order')
                ws_ozet.write_formula('E4', f_dyn_qord, fmt_num)
                ws_ozet.write('D5', 'Efficiency Ratio %')
                ws_ozet.write_formula('E5', f_dyn_ratio, fmt_pct)

                # --- FILTER SETUP ---
                ws_ozet.write('A11', 'Table Control', fmt_lbl)
                ws_ozet.write('A12', 'Selected Line:', fmt_lbl)
                ws_ozet.write('B12', '(Type Line)', fmt_inp)

                # --- FULL BREAKDOWN TABLE (WITH CINSIYET) ---
                ws_ozet.write('A15', 'Full Summary', fmt_h1)
                final_pivot.to_excel(writer, sheet_name='Özet', index=False, startrow=16)
                
                # Column sizing
                ws_ozet.set_column('A:A', 25)
                ws_ozet.set_column('B:B', 20)
                ws_ozet.set_column('C:C', 20)
                ws_ozet.set_column('D:D', 20)
                ws_ozet.set_column('E:E', 20)
                ws_ozet.set_column('F:Z', 15)

            progress_bar.progress(100)
            status_text.text("Task Finished!")
            time.sleep(1)
            status_text.empty()
            progress_bar.empty()

            st.success("✅ Report generated successfully!")
           
            st.download_button(
                label="📥 Download Excel Report",
                data=output.getvalue(),
                file_name="Satış Raporu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error during automation: {str(e)}")
else:
    st.info("Waiting for all 4 Excel files to be uploaded.")
