
-- verson 5
import streamlit as st

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

if st.button("Generate Formula"):
    formula = (
        f'=IF({best_model}="12 WK MA", {val_12wk},'
        f' IF({best_model}="8 WK MA", {val_8wk},'
        f'  IF({best_model}="4 WK MA", {val_4wk},'
        f'   IF(OR({best_model}="With Promo No Holidays",'
        f'         {best_model}="With Promo With Holidays",'
        f'         {best_model}="No promo no holidays",'
        f'         {best_model}="No Promo With Holidays"),'
        f'     AVERAGE('
        f'      INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
        f'       MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
        f'       MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + 1'
        f'      ):'
        f'      INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
        f'       MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
        f'       MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {full_truck}'
        f'      )'
        f'     ),'
        f'     {best_model}'   # << ELSE: just return the best_model value
        f'   )'
        f'  )'
        f' )'
        f')'
    )
    st.text_area("Excel formula (copy-paste):", formula, height=200)


# #ver 4
# import streamlit as st

# st.title("Excel Formula Replacer Tool")

# colLeft, colCenter, colRight = st.columns([1,1,1])

# with colLeft:
#     val_12wk = st.text_input("Value of 12 WK MA Cell", placeholder="e.g. AN26")
#     val_8wk = st.text_input("Value of 8 WK MA Cell", placeholder="e.g. AO26")
#     val_4wk = st.text_input("Value of 4 WK MA Cell", placeholder="e.g. AP26")

# with colCenter:
#     best_model = st.text_input("Best Model Indicator Cell", placeholder="e.g. AR29")
#     item_code = st.text_input("Item Code Cell", placeholder="e.g. AO29")

# with colRight:
#     last_friday = st.text_input("Updated Last Friday Cell (with $)", placeholder="e.g. AT$15")
#     full_truck = st.text_input("# Weeks Full Truck Cell (with $)", placeholder="e.g. AQ$12")

# if st.button("Generate Formula"):
#     formula = (
#         f'=IF(OR({best_model}="12 WK MA", {best_model}="8 WK MA", {best_model}="4 WK MA"),'
#         f' IF({best_model}="12 WK MA", {val_12wk},'
#         f'  IF({best_model}="8 WK MA", {val_8wk},'
#         f'   {val_4wk}'
#         f'  )'
#         f' ),'
#         f' IF(ISNUMBER(SEARCH("promo", {best_model})),'
#         f'  IF(MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {full_truck} - 1 > MATCH("AN", \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0),'
#         f'   "Not enough predicted values - check inputs",'
#         f'   IF(MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) > MATCH("AN", \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0),'
#         f'    "No prediction available at this stage",'
#         f'    AVERAGE('
#         f'     INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
#         f'      MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
#         f'      MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0)'
#         f'     ):' 
#         f'     INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
#         f'      MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
#         f'      MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {full_truck} - 1'
#         f'     )'
#         f'    )'
#         f'   )'
#         f'  ),'
#         f'  {best_model}'
#         f' )'
#         f')'
#     )
#     st.text_area("Excel formula (copy-paste):", formula, height=200)





# ver 3
# import streamlit as st
# import re

# st.title("Excel Formula Replacer Tool")

# colLeft, colCenter, colRight = st.columns([1,1,1])

# with colLeft:
#     val_12wk = st.text_input("Value of 12 WK MA Cell", placeholder="e.g. AN26")
#     val_8wk = st.text_input("Value of 8 WK MA Cell", placeholder="e.g. AO26")
#     val_4wk = st.text_input("Value of 4 WK MA Cell", placeholder="e.g. AP26")

# with colCenter:
#     best_model = st.text_input("Best Model Indicator Cell", placeholder="e.g. AR26")
#     item_code = st.text_input("Item Code Cell", placeholder="e.g. E26")

# with colRight:
#     last_friday = st.text_input("Updated Last Friday Cell (with $)", placeholder="e.g. AX$24")
#     full_truck = st.text_input("# Weeks Full Truck Cell (with $)", placeholder="e.g. AT$22")

# def col_letter_to_index(col):
#     num = 0
#     for c in col.upper():
#         num = num*26 + (ord(c) - ord('A') + 1)
#     return num

# def get_col(cell):
#     m = re.match(r"\$?([A-Za-z]+)", cell)
#     return m.group(1) if m else None

# def generate_formula():
#     bm_value_lower = best_model.strip().lower()
#     ma_models = ["12 wk ma", "8 wk ma", "4 wk ma"]

#     if not any(model == bm_value_lower for model in ma_models) and ("promo" not in bm_value_lower):
#         return f'={best_model}'
#     else:
#         last_friday_col = get_col(last_friday)
#         full_truck_num = int(re.search(r"(\d+)", full_truck).group())
#         max_col = col_letter_to_index("AN")
#         last_friday_col_idx = col_letter_to_index(last_friday_col)
#         range_end = last_friday_col_idx + full_truck_num

#         if last_friday_col_idx < max_col < range_end:
#             return '"Not enough predicted values - check inputs"'
#         elif last_friday_col_idx >= max_col:
#             return '"No prediction yet"'
#         else:
#             return (
#                 f'=IF({best_model}="12 WK MA", {val_12wk},'
#                 f' IF({best_model}="8 WK MA", {val_8wk},'
#                 f'  IF({best_model}="4 WK MA", {val_4wk},'
#                 f'   AVERAGE('
#                 f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
#                 f'     MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
#                 f'     MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + 1'
#                 f'    ):' 
#                 f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
#                 f'     MATCH({item_code}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
#                 f'     MATCH({last_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {full_truck_num}'
#                 f'    )'
#                 f'   )'
#                 f'  )'
#                 f' )'
#                 f')'
#             )

# if "formula" not in st.session_state:
#     st.session_state.formula = ""

# if st.button("Generate Formula"):
#     st.session_state.formula = generate_formula()

# st.text_area("Excel formula (copy-paste):", st.session_state.formula, height=180)





# -- verson 2
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
