
# ver 3
import streamlit as st
import re

st.title("Excel Formula Replacer Tool")

colLeft, colCenter, colRight = st.columns([1,1,1])

with colLeft:
    val_12wk = st.text_input("Value of 12 WK MA Cell", value="", placeholder="e.g. AN26")
    val_8wk = st.text_input("Value of 8 WK MA Cell", value="", placeholder="e.g. AO26")
    val_4wk = st.text_input("Value of 4 WK MA Cell", value="", placeholder="e.g. AP26")

with colCenter:
    best_model = st.text_input("Best Model Indicator Cell", value="", placeholder="e.g. AR26")
    item_code = st.text_input("Item Code Cell", value="", placeholder="e.g. E26")

with colRight:
    last_friday = st.text_input("Updated Last Friday Cell (with $)", value="", placeholder="e.g. AX$24")
    full_truck = st.text_input("# Weeks Full Truck Cell (with $)", value="", placeholder="e.g. AT$22")

def excel_col_to_num(col_str: str) -> int:
    num = 0
    for c in col_str:
        if c.isalpha():
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

def get_col_and_row(cell_ref):
    m = re.match(r"(\$?)([A-Za-z]+)(\$?)(\d+)", cell_ref)
    if not m:
        return None, None
    return m.group(2).replace('$',''), m.group(4)

last_friday_col, _ = get_col_and_row(last_friday)
full_truck_col, full_truck_row = get_col_and_row(full_truck)
max_col_limit = excel_col_to_num("AN")

try:
    full_truck_val = int(full_truck_row)
except:
    full_truck_val = None

generate_col1, generate_col2, generate_col3 = st.columns([1,1,1])
if generate_col2.button("Generate Formula"):
    check_best_model = best_model.strip().lower()
    if ("wk ma" not in check_best_model) and ("promo" not in check_best_model):
        # Output INDIRECT formula to fetch dynamic value at best_model cell
        formula = f'=INDIRECT("{best_model}")'
    else:
        if last_friday_col is None or full_truck_val is None:
            formula = '"Error: Invalid cell references for last friday or weeks full truck"'
        else:
            last_friday_num = excel_col_to_num(last_friday_col)
            range_end = last_friday_num + full_truck_val
            if last_friday_num < max_col_limit < range_end:
                formula = '"Not enough predicted values - check inputs"'
            elif last_friday_num >= max_col_limit:
                formula = '"No prediction yet"'
            else:
                formula = (
                    f'=IF({best_model}="12 WK MA", {val_12wk},'
                    f' IF({best_model}="8 WK MA", {val_8wk},'
                    f'  IF({best_model}="4 WK MA", {val_4wk},'
                    f'   AVERAGE('
                    f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
                    f'     MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
                    f'     MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + 1'
                    f'    ):' 
                    f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
                    f'     MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
                    f'     MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {full_truck_val}'
                    f'    )'
                    f'   )'
                    f'  )'
                    f' )'
                    f')'
                )
    st.text_area("Excel formula (copy-paste):", formula, height=180)



# # -- verson 2
# import streamlit as st

# st.title("Excel Formula Replacer Tool")

# colLeft, colCenter, colRight = st.columns([1,1,1])

# with colLeft:
#     val_12wk = st.text_input("Value of 12 WK MA Cell", value="", placeholder="e.g. AN26")
#     val_8wk = st.text_input("Value of 8 WK MA Cell", value="", placeholder="e.g. AO26")
#     val_4wk = st.text_input("Value of 4 WK MA Cell", value="", placeholder="e.g. AP26")

# with colCenter:
#     best_model = st.text_input("Best Model Indicator Cell", value="", placeholder="e.g. AR26")
#     item_code = st.text_input("Item Code Cell", value="", placeholder="e.g. E26")

# with colRight:
#     last_friday = st.text_input("Updated Last Friday Cell (with $)", value="", placeholder="e.g. AX$24")
#     full_truck = st.text_input("# Weeks Full Truck Cell (with $)", value="", placeholder="e.g. AT$22")

# if st.button("Generate Formula"):
#     formula = (
#         f'=IF({best_model}="12 WK MA", {val_12wk},'
#         f' IF({best_model}="8 WK MA", {val_8wk},'
#         f'  IF({best_model}="4 WK MA", {val_4wk},'
#         f'   AVERAGE('
#         f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
#         f'     MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
#         f'     MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + 1'
#         f'    ):'
#         f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
#         f'     MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
#         f'     MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {full_truck}'
#         f'    )'
#         f'   )'
#         f'  )'
#         f' )'
#         f')'
#     )
#     st.text_area("Excel formula (copy-paste):", formula, height=180)



# import streamlit as st

# st.title("Excel Formula Replacer Tool")

# AR26 = st.text_input("Replace {AR}{26} (Best model indicator):", "AR26")
# AN26 = st.text_input("Replace {AN}{26} (Value of 12 WK MA):", "AN26")
# AO26 = st.text_input("Replace {AO}{26} (Value of 8 WK MA):", "AO26")
# AP26 = st.text_input("Replace {AP}{26} (Value of 4 WK MA):", "AP26")
# E26 = st.text_input("Replace {E}{26} (Item code):", "E26")
# AX24 = st.text_input("Replace {AX}${24} (Updated last Friday):", "AX$24")
# AT22 = st.text_input("Replace {AT}${22} (# Weeks full truck):", "AT$22")

# if st.button("Generate Formula"):
#     formula = (
#         f'=IF({AR26}="12 WK MA", {AN26},'
#         f' IF({AR26}="8 WK MA", {AO26},'
#         f'  IF({AR26}="4 WK MA", {AP26},'
#         f'   AVERAGE('
#         f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
#         f'     MATCH({E26}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
#         f'     MATCH({AX24}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + 1'
#         f'    ):'
#         f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
#         f'     MATCH({E26}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
#         f'     MATCH({AX24}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {AT22}'
#         f'    )'
#         f'   )'
#         f'  )'
#         f' )'
#         f')'
#     )
#     st.text_area("Excel formula (copy-paste):", formula, height=150)
