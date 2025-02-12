import streamlit as st
import pandas as pd
# from rapidfuzz import process, fuzz
from datetime import datetime


# # Function to fuzzy match
# def fuzzy_match(query, choices, limit=50):
#     """
#     对输入的关键字进行模糊匹配。
#     :param query: 用户输入的关键字
#     :param choices: 数据源列表
#     :param limit: 返回的最大匹配结果数量
#     :return: 匹配结果列表
#     """
#     if not query:
#         return []
    
#     results = process.extract(query, choices, limit=limit, scorer=fuzz.WRatio)
    
#     return [result[0] for result in results if result[1] > 70] # 相似度大于 70


# Function to match iscontain
def substring_search(query, choices):
    """
    对输入的关键字进行子字符串匹配。
    :param query: 用户输入的关键字
    :param choices: 数据源列表
    :return: 匹配结果列表
    """
    if not query:
        return []
    # 不区分大小写的匹配
    query = query.lower()
    matches = [item for item in choices if query in item.lower()]
    return matches    


def parse_dates(df):
    """解析Dates列为开始月份和结束月份"""
    month_map = {
        'Jan':1, 'Feb':2, 'Mar':3, 'Apr':4, 'May':5, 'Jun':6,
        'Jul':7, 'Aug':8, 'Sep':9, 'Oct':10, 'Nov':11, 'Dec':12,
        'May ':5, 'Jun ':6, 'Jul ':7, 'Aug ':8, 'Sep ':9, 'Oct ':10, 'Nov ':11, 'Dec ':12
    }
    
    def extract_month(date_str):
        parts = date_str.split(' to ')
        start_month = parts[0].split('-')[0].strip().capitalize()[:3]  # 取前3个字符
        end_month = parts[1].split('-')[0].strip().capitalize()[:3]
        
        try:
            return month_map[start_month], month_map[end_month]
        except KeyError as e:
            raise ValueError(f"无效的月份缩写: {e.args[0]},请检查日期格式是否为'Jan-1 to Dec-31'") from e
    
    df[['Start Month', 'End Month']] = df['Dates'].apply(
        lambda x: pd.Series(extract_month(x))
    )
    return df


def get_city_cap(city_name, country_name, month, rank, mate):
    if city_name != '':
        city_cap = df_city[(df_city['City'] == city_name) & (df_city['Start Month'] <= month) & (df_city['End Month'] >= month)]['City Cap'].iloc[0]
        currency = df_city[(df_city['City'] == city_name) & (df_city['Start Month'] <= month) & (df_city['End Month'] >= month)]['Currency'].iloc[0]
        vat_included = df_city[(df_city['City'] == city_name) & (df_city['Start Month'] <= month) & (df_city['End Month'] >= month)]['VAT/GST Included'].iloc[0]
        
        if mate is True and (rank == 'PPEDD' or rank == 'SM & M'):
            city_cap = city_cap * 2
            msg = f'{city_name} in {month} month, {currency} {city_cap}, VAT/GST Included: {vat_included}'
        else:
            msg = f'{city_name} in {month} month, {currency} {city_cap}, VAT/GST Included: {vat_included}'
            
    elif country_name != '':
        city_cap = df_country[(df_country['Location'] == country_name) & (df_country['Start Month'] <= month) & (df_country['End Month'] >= month)]['City Cap'].iloc[0]
        currency = df_country[(df_country['Location'] == country_name) & (df_country['Start Month'] <= month) & (df_country['End Month'] >= month)]['Currency'].iloc[0]
        vat_included = df_country[(df_country['Location'] == country_name) & (df_country['Start Month'] <= month) & (df_country['End Month'] >= month)]['VAT/GST Included'].iloc[0]

        if mate is True and (rank == 'PPEDD' or rank == 'SM & M'):
            city_cap = city_cap * 2
            msg = f'{country_name} in {month} month, {currency} {city_cap}, VAT/GST Included: {vat_included}'
        else:
            msg = f'{country_name} in {month} month, {currency} {city_cap}, VAT/GST Included:{vat_included}'

    elif rank != '':
        if rank == 'PPEDD':
            city_cap = 150
            currency = 'USD'
            vat_included = 'No'

            if mate is True:
                city_cap = city_cap * 2
                msg = f'{currency} {city_cap}, VAT/GST Included: {vat_included}'
            else:
                msg = f'{currency} {city_cap}, VAT/GST Included: {vat_included}'

        elif rank == 'SM & M':
            city_cap = 120
            currency = 'USD'
            vat_included = 'No'

            if mate is True:
                city_cap = city_cap * 2
                msg = f'{currency} {city_cap}, VAT/GST Included: {vat_included}'
            else:
                msg = f'{currency} {city_cap}, VAT/GST Included: {vat_included}'
        
        elif rank == 'Senior & Staff':
            city_cap = 120
            currency = 'USD'
            vat_included = 'No'
            
            msg = f'{currency} {city_cap}, VAT/GST Included: {vat_included}'

    return city_cap, currency, vat_included, msg


# Function to calculate amount
def calculate_amount(city_cap, vat_included, 
                     total_cost, num_nights, service_fee, tax, other):
    
    amount_accommodation = total_cost - service_fee - other

    if vat_included == 'Yes':
        if tax > 1:
            max_amount = amount_accommodation * num_nights + tax
            if amount_accommodation > max_amount:
                amount = max_amount
            else:
                amount = amount_accommodation

        elif tax > 0:
            max_amount = amount_accommodation * num_nights * (1 + tax)
            if amount_accommodation > max_amount:
                amount = max_amount
            else:
                amount = amount_accommodation

    else:
        if tax > 1:
            max_amount = city_cap * 1.15 * num_nights + tax
            if amount_accommodation > max_amount:
                amount = max_amount
            else:
                amount = amount_accommodation

        elif tax >= 0:
            max_amount = city_cap * 1.15 * num_nights * (1 + tax)
            if amount_accommodation > max_amount:
                amount = max_amount
            else:
                amount = amount_accommodation

    return amount





# Change style of streamlit
st.set_page_config(
page_title='BP',
page_icon='kt.ico'  # http is fine too
)
# Change style with CSS
with open('style.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)


# Title
st.title('🎰 City Cap')

rank_list = ['PPEDD', 'SM & M', 'Senior & Staff']

# Initialize session state
if 'city_cap_msg' not in st.session_state:
    st.session_state.city_cap_msg = ''

# With expander:
with st.expander('Upload City Cap'):
    # Upload files
    tb_path = st.file_uploader("Select data", type=["xlsx", "xls"], accept_multiple_files=False)

    # Initialize lists
    city_list = []
    country_list = []

    # Read data
    if tb_path:
        df_city = pd.read_excel(tb_path, sheet_name=2, header=7, index_col=None, dtype=object)
        df_country = pd.read_excel(tb_path, sheet_name=3, header=7, index_col=None, dtype=object)

        # Replace \n with space in columns' names
        df_city.columns = df_city.columns.str.replace('\n', ' ')
        df_country.columns = df_country.columns.str.replace('\n', ' ')

        city_list = df_city['City'].tolist()  # Override if file is uploaded
        country_list = df_country['Location'].tolist()

        df_city = parse_dates(df_city)  # Parse dates
        df_country = parse_dates(df_country)

        # st.write(df_city, df_country)

# 使用表单包装城市cap相关输入
with st.form('city_cap_form'):
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # Input City 
        city_input = st.text_input('City', key='city_name_form')
        
        if city_input and tb_path:
            city_matches = substring_search(city_input, city_list)
            if city_matches:
                st.write(city_matches)
        

        country_input = ''  # 初始化变量
        if not (city_input and tb_path):
            country_input = st.text_input('Country', key='country name')
        if country_input:
                country_matches = substring_search(country_input, country_list)
                if country_matches:
                    st.write(country_matches)
                else:
                    pass
            

    with col2:

        # Select rank
        rank_input = st.selectbox('Rank', rank_list, key='rank')

        # Month selection
        month_input = st.selectbox('Month', range(1, 13), key='month')

        # Y/N selection
        mate_check = st.checkbox('Room mate?', key='room mate')


    # City cap button
    button_city_cap = st.form_submit_button('Get City Cap')
    if button_city_cap:
        return_city = get_city_cap(city_input, country_input, month_input, rank_input, mate_check)
        
        st.session_state.cal_city_cap = return_city[0]  # 存储到session_state
        st.session_state.cal_currency = return_city[1]
        st.session_state.cal_vat_included = return_city[2]
        st.session_state.city_cap_msg = return_city[3]

    
st.success(st.session_state.city_cap_msg)
if st.session_state.city_cap_msg != '':
    # Input form
    with st.form('form'):
        col3, col4 = st.columns([1, 1])
        with col3:
            # Total cost
            total_cost = st.number_input('Total cost', key='total cost')
            # Check-in date
            check_in_date = st.date_input('Check-in date', key='check in date')
            # Check-out date
            check_out_date = st.date_input('Check-out date', key='check out date')

            # num_nights = check_out_date - check_in_date
            num_nights = (check_out_date - check_in_date).days

        
        with col4:
            # Service Fee
            service_fee = st.number_input('Service fee', key='service fee')
            # Tax
            tax = st.number_input('Tax', key='tax')
            # Other
            other = st.number_input('Other', key='other')

        # Submit button
        submitted = st.form_submit_button('Calculate')
        if submitted:
            calcu_cost = calculate_amount(
                        city_cap=st.session_state['cal_city_cap'],
                        vat_included=st.session_state['cal_vat_included'],
                        total_cost=total_cost,
                        num_nights=num_nights, 
                        service_fee=service_fee, 
                        tax=tax, 
                        other=other)
            
            st.success(f"Reimbursable Amount: {calcu_cost}")


# Debugging
st.write(st.session_state)
