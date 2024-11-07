


# import os
#
# try:
#    os.system('python -m pip install --upgrade pip')
# except ImportError:
#     os.system('python -m pip install --upgrade pip')
#
# # breakpoint()
#
#
# #
# try:
#     import streamlit
# except ImportError:
#     os.system('python -m pip install streamlit')
#
#
#
#
# try:
#     import streamlit_option_menu
# except ImportError:
#     os.system('python -m pip install streamlit-option-menu')
#
# try:
#     import openpyxl
# except ImportError:
#     os.system('python -m pip install openpyxl')
#
#
# try:
#     import xlwings
# except ImportError:
#     os.system('python -m pip install xlwings==0.30.12')
# try:
#     import pandas
# except ImportError:pip install pywin32
#     os.system('python -m pip install pandas')

# pip install pywin32
#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


from pandas.io.json._normalize import json_normalize
import streamlit as st  #  pip install streamlit  pyarrow==14.0.1
import pandas as pd
# import plotly.express as px       #   pip install json_normaliz                          pip install plotly
import copy
from streamlit_option_menu import option_menu     # pip install streamlit-option-menu
import  vlookup_automation_27_10_2024  as va


st.set_page_config(page_title="EXCEL AUTOMATION",page_icon="📊",layout="wide",initial_sidebar_state="expanded",)  # "auto" or "expanded" or "collapsed"
# Set dark theme using custom CSS
st.markdown(
    """
    <style>
        body {
            color: #FFFFFF;  /* Text color */
            background-color: 'black';  /* Background color */
        }
        /* Add more custom styles as needed */
    </style>
    """,
    unsafe_allow_html=True)


st.title('LIT FINANCE DASHBOARD')



import datetime, time

def time_():
    time_ = time.strftime("%H:%M:%S", time.localtime())
    return (time_)

# import streamlit as st


def active_Workbook_name():
    if xw.apps:
        wb = xw.apps.active.books.active
        return wb.name  # Return the workbook name instead of printing it
    else:
        return "No active Excel application found."

total1,total2,total3,total4= st.columns(4, gap="small")
import streamlit as st


def styled_metric2(label, value, label_style="", value_style="", label_size="16px", value_size="16px", background_color="#FFFFFF", border_left_color="#f20045", border_left_size="5px", padding_size="10px", box_width="300px", box_height="150px"):
    styled_html = f"""
        <div style="
            background-color: {background_color};
            border-left: {border_left_size} solid {border_left_color};
            padding: {padding_size};
            width: {box_width};
            height: {box_height};
            ">
            <p style="{label_style} margin: 0; font-size: {label_size};">{label}</p>
            <p style="{value_style} margin: 10; font-size: {value_size};">{value}</p>
        </div>
    """
    st.markdown(styled_html, unsafe_allow_html=True)

def color_background_color(value):
    if value > 0:
        color = "#00FF00"         # Lime
    else:
        color = "#8B0000"         # red
    return color






def run():

        # all_qttar_dict_list = uploaded_file2(all_qttar_dict=all_qttar_dict, key="key2")
        st.sidebar.image("devlopar.png",
                        caption="Developed and   application problem solving   by: KUKAN MANOJ    : helpline   number -> 8000594016")
        # print("all_qttar_dict_list lin 173 ", time_())
        with st.sidebar:
            app = option_menu(
                menu_title='Pondering ',
                options=['EXCEL AUTOMATION','VLOOKUP clicked'],#, 'your uploaded_file','Add Your Capital', 'Holding', 'Intraday', 'Options','Charges','Charges Debits and Credits','Dividends', 'Settings', 'about'],
                #           https: // icons.getbootstrap.com /  # icons     'person-circle'

                icons=['house-fill','journal-arrow-up'],#,'journal-arrow-up', 'trophy-fill', 'apple', 'play-btn-fill','caret-right-square-fill','caret-right-square-fill','caret-right-square-fill','gear','info-circle-fill'],
                menu_icon='chat-text-fill',
                default_index=0,
                styles={"container": {"padding": "5!important", "background-color": 'black'},
                        "icon": {"color": "white", "font-size": "30px"},
                        "nav-link": {"color": "white", "font-size": "20px", "text-align": "left", "margin": "0px",
                                     "--hover-color": "blue"},
                       "nav-link-selected": {"background-color": "#02ab21"}, })

        if app == 'EXCEL AUTOMATION':
            st.image("devlopar.png",
                             caption="Developed and   application problem solving   by: KUKAN MANOJ    : helpline   number -> 8000594016")

            # print("option_menu lin 189 ", time_())
        if app == "VLOOKUP clicked":

            # Create a row with multiple text input fields in a single line table

            st.title('{1}selected_Tabal')

            col1, col2, col3= st.columns(3)

            with col1:

                st.text_input("active_Workbook_name",value= active_Workbook_name())

            with col2:
                selected_table_array_sheet_name = st.selectbox("selected_table_array_sheet_name",va.sheet_names_list(),key='selected_VLOOKUP_sheet_name')

            with col3:
                selected_table_array_sheet_range = st.text_input('selected_table_array_range','A1:Y30',key='selected_VLOOKUP_sheet_range')

            st.title('{2}selected_VLOOKUP_columns_and_row')

            col4, col5, col6 = st.columns(3)  #lookup_value

            with col4:
                lookup_value_sheets_name = st.selectbox('lookup_value_sheets_name', va.sheet_names_list(),index=1 , key='lookup_value_sheets_name')
            with col5:
                lookup_value_columns = st.text_input('lookup_value_sheets_columns','A1:Y1',key='lookup_value_sheets_columns')
            with col6:
                lookup_value_row = st.text_input('lookup_value_sheet_row', 'A1:A13', key='lookup_value_sheet_row')
            # print(va.sheet_names_list()[2])
            st.title('{3}final_output_sheet')

            col7, col8, col9 = st.columns(3)

            with col7:
                final_output_sheet_name = st.selectbox('final_output_sheet_name', va.sheet_names_list(), index=2,
                                                   key='final_output_sheet_name')
            with col8:
                final_output_cell = st.text_input('final_output_cell', 'A1', key='final_output_cell')

            if st.button("DOWNLOD clicked"):
                # st.write("Download initiated OK!")

                selected_table_array = va.sheets_tabal_selet(sheet=selected_table_array_sheet_name, range=selected_table_array_sheet_range)

                data_vlookup = va.df_vlookup(df=selected_table_array, vlookup_sheets= lookup_value_sheets_name,
                                             filtered_columns=lookup_value_columns, filtered_row=lookup_value_row)

                # print(data_vlookup)
                # #
                # print('DOWNLOD clicked', time_())

                final_output_cell = va.final_output_cell(Sheet_names=final_output_sheet_name, df=data_vlookup,
                                                         range=final_output_cell)


                lookup_value = f'${lookup_value_row[0]}{int(lookup_value_row[1]) + 1}'

                table_array_start_cell, table_array_end_cell = selected_table_array_sheet_range.split(':')

                table_array_start_cell_column = va.split_cell_column(cell_column=table_array_start_cell)[0]
                table_array_start_cell_row = int(va.split_cell_column(cell_column=table_array_start_cell)[1])
                table_array_start_cell =f'${table_array_start_cell_column}${table_array_start_cell_row}'

                table_array_end_cell_column = va.split_cell_column(cell_column=table_array_end_cell)[0]
                table_array_end_cell_row = int(va.split_cell_column(cell_column=table_array_end_cell)[1])
                table_array_end_cell = f'${table_array_end_cell_column}${table_array_end_cell_row}'



                table_array = f'{selected_table_array_sheet_name}!{table_array_start_cell}:{table_array_end_cell}'

                # MATCH = f'{lookup_value_sheets_name}!{lookup_value_row[0]}{int(lookup_value_row[1]) }'


                # Split the string into start and end cell references
                start_cell, end_cell = lookup_value_columns.split(':')

                column = va.split_cell_column(cell_column=start_cell)[0]
                row = int(va.split_cell_column(cell_column=start_cell)[1])

                column_letters_list = va.column_letters_list()
                column_letters_list_index =column_letters_list.index(column)

                MATCH = f'{lookup_value_sheets_name}!{column_letters_list[column_letters_list_index + 1 ]}${row}'



                lookup_value_start_cell, lookup_value_end_cell = lookup_value_columns.split(':') #lookup_value_columns

                lookup_value_start_column = va.split_cell_column(cell_column=lookup_value_start_cell)[0]
                lookup_value_start_row = int(va.split_cell_column(cell_column=lookup_value_start_cell)[1])
                lookup_value_start_cell =f'${lookup_value_start_column }${lookup_value_start_row}'

                lookup_value_end_cell_column = va.split_cell_column(cell_column=lookup_value_end_cell)[0]
                lookup_value_end_cell_row = int(va.split_cell_column(cell_column=lookup_value_end_cell)[1])
                table_array_end_cell = f'${lookup_value_end_cell_column}${lookup_value_end_cell_row}'
                #
                # st.code(f'{lookup_value_start_cell}:{table_array_end_cell}  //{lookup_value_columns}', language='excel')
                #

                MATCH_rang = f'{lookup_value_sheets_name}!{lookup_value_start_cell}:{table_array_end_cell}'


                code3 = f"=VLOOKUP({lookup_value},{table_array},MATCH({MATCH},{MATCH_rang},0),0)"

                st.title('copy_paste_code_to_excel')

                # Display the formula as a code block in Streamlit
                st.code(f'{code3}', language='excel')
                #
                # st.code(f'=VLOOKUP($A2,Sheet1!$A$1:$Y$30,MATCH(Sheet2!B$1,Sheet2!$A$1:$Y$1,0),0)', language='excel')






run()



# try:
#     st.text_input("active_Workbook_name", value=va.active_Workbook_name())
#     run()
# except:
#     print(" opne active_Workbook_name went")

# streamlit run vlookup_automation_home_29_10_2024.py

#  Install packages from requirements.txt: Run this command in your terminal or command prompt where the file is located.

# pip install -r requirements_2.txt


# Generate the requirements.txt File
# pip freeze > requirements.txt
