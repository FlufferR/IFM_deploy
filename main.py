# Use Streamlit to upload excel file
import tempfile
from datetime import datetime

import streamlit as st
import pandas as pd


def bo_process(bo_file, map_file):
    # Read the first sheet of bo_file
    bo_df = pd.read_excel(bo_file, sheet_name=0)
    map_vendor = pd.read_excel(map_file, sheet_name='Sending Entity_Vendor Mapping')
    map_receiving = pd.read_excel(map_file, sheet_name='Receiving Entity')
    map_area = pd.read_excel(map_file, sheet_name='Country Area Mapping')

    # Keep 'Vendor Id - Ven' 'Vendor Name1 - Ven' 'Vendor Tyep' 'Country' in map_vendor
    map_vendor = map_vendor[['Vendor Id - Ven', 'Vendor Name1 - Ven', 'Vendor Tyep', 'Country']]
    map_receiving = map_receiving[['AP Business Unit', 'Receiving Country', 'LE Name']]
    map_area = map_area[['Country', 'Area']]

    # Drop duplicates in map_vendor based on 'Vendor Id - Ven'
    # map_vendor.drop_duplicates(subset=['Vendor Id - Ven'], keep='first', inplace=True)

    # Merge bo_df with map_vendor on 'Vendor Id - Ven' in map_vendor and 'Vendor Id - AP' in bo_df
    bo_df = pd.merge(bo_df, map_vendor, left_on='Vendor Id - AP', right_on='Vendor Id - Ven', how='left')
    bo_df = pd.merge(bo_df, map_receiving, left_on='Business Unit - AP', right_on='AP Business Unit', how='left')
    bo_df = pd.merge(bo_df, map_area, left_on='Country', right_on='Country', how='left')

    # Rename 'Vendor Tyep' to 'External/INTERFIRM' 'Country' to 'Bill To/FmCountry'
    bo_df.rename(columns={'Vendor Tyep': 'External/INTERFIRM', 'Country': 'Bill To/FmCountry',
                          'Vendor Name1 - Ven': 'Bill To/Fm Legal Entity', 'Receiving Country': 'GC Country',
                          'LE Name': 'GC Legal Entity', 'Area': 'Bill To/Fm Area', 'Invoice Id - AP': 'Invoice No.',
                          'Invoice Date - AP': 'Invoice Date', 'Currency Cd - AP': 'Base currency of Country',
                          'Monetary Amount Detail - AP': 'Base amount of Country',
                          'Foreign Currency - AP': 'Original Currency',
                          'Foreign Amount Detail - AP': 'Original billing amount',
                          'Business Unit - AP': 'Business Unit - AP/AR', 'Vendor Id - AP': 'Vendor ID (AP)',
                          }, inplace=True)

    # Add 'Account' = if 'External/INTERFIRM' == 'IFM' then '38200015' else '38000000'
    bo_df['Account'] = bo_df.apply(lambda x: '38200015' if x['External/INTERFIRM'] == 'IFM' else '38000000', axis=1)

    bo_df = bo_df[['External/INTERFIRM', 'GC Country', 'GC Legal Entity', 'Bill To/Fm Area',
                   'Bill To/FmCountry', 'Bill To/Fm Legal Entity', 'Invoice No.', 'Invoice Date',
                   'Base currency of Country', 'Base amount of Country', 'Original Currency',
                   'Original billing amount', 'Business Unit - AP/AR', 'Account', 'Vendor ID (AP)']]
    return bo_df


def ifm_process(bo_df, last_month_file, exchange_rate):
    # Read last_month_file
    last_month_df = pd.read_excel(last_month_file, sheet_name=0)

    # Concat bo_df and last_month_df
    ifm_df = pd.concat([last_month_df, bo_df], axis=0, join='outer', ignore_index=True)
    # ifm_df['Ex. Rate to USD'] = if ['Original Currency'] == 'USD' then 1 else exchange_rate
    ifm_df['Ex. Rate to USD'] = ifm_df.apply(lambda x: 1 if x['Original Currency'] == 'USD' else exchange_rate, axis=1)
    ifm_df['Function '] = 'AP'
    ifm_df['Amount in USD'] = ifm_df['Original billing amount'] / ifm_df['Ex. Rate to USD']

    # Sort in ascending order by 'Invoice No.'
    ifm_df.sort_values(by=['Invoice No.'], ascending=True, inplace=True)

    # ifm_df['Bill To/Fm SubArea'] = if ifm_df['Bill To/FmCountry'] ==
    # 'CHINA' or 'HONG KONG' or 'TAIWAN' then 'Great China'
    ifm_df['Bill To/Fm SubArea'] = ifm_df.apply(lambda x: 'Great China' if x['Bill To/FmCountry'] in
                                                                           ['China', 'China-HK', 'China-TW'] else '', axis=1)

    # Group by y 'Invoice No.' and 'Vendor ID (AP)'
    ifm_df = ifm_df.groupby(['Invoice No.', 'Vendor ID (AP)']).agg({
        'External/INTERFIRM': 'first',
        'GC Country': 'first',
        'GC Legal Entity': 'first',
        'Bill To/Fm Area': 'first',
        'Bill To/FmCountry': 'first',
        'Bill To/Fm Legal Entity': 'first',
        'Bill To/Fm SubArea': 'first',
        'Invoice No.': 'first',
        'Invoice Date': 'first',
        # 'Base currency of Country': 'first',
        'Base amount of Country': 'sum',
        'Original Currency': 'first',
        'Original billing amount': 'sum',
        'Business Unit - AP/AR': 'first',
        'Account': 'first',
        'Vendor ID (AP)': 'first',
        'Ex. Rate to USD': 'first',
        'Function ': 'first',
        'Amount in USD': 'sum'
    })

    # Delete rows which 'Base amount of Country' is 0
    ifm_df = ifm_df[ifm_df['Base amount of Country'] != 0]

    return ifm_df


def df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8-sig')


def main():
    st.set_page_config(
        page_title='IFM Tool',
        page_icon='kt.ico'  # http is fine too
    )

    # Hide the streamlit style with CSS
    with open('hide_streamlit_style.css') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

    st.title("IFM Report Generator")
    bo_file = st.file_uploader("Upload BO file", type=["xlsx", "xls"])
    last_month_file = st.file_uploader("Upload Last Month file", type=["xlsx", "xls"])
    map_file = st.file_uploader("Upload Map file", type=["xlsx", "xls"])

    # Add a label used to input the exchange rate
    exchange_rate = st.number_input("Exchange Rate", value=1.0, step=0.01)

    # Add a button to trigger the calculation
    if st.button("Generate"):
        # Perform the calculation and display the result
        bo_df = bo_process(bo_file, map_file)
        result = ifm_process(bo_df, last_month_file, exchange_rate)
        st.write(result)
        output = df_to_csv(result)
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Add a button to download the result
        st.download_button(
            label='Download data',
            data=output,
            file_name=f'IFM report {current_time}.csv',
            # mime='application/vnd.ms-excel'
            mime='text/csv'
        )


if __name__ == "__main__":
    main()