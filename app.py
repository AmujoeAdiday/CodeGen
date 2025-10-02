import streamlit as st

st.title("Excel Formula Replacer Tool")

col12wk, colBestModel, colRight = st.columns([1, 1, 1])

with col12wk:
    cell_12wk = st.text_input("Cell of Value of 12 WK MA", value="", placeholder="e.g. AN26")
    cell_8wk = st.text_input("Cell of Value of 8 WK MA", value="", placeholder="e.g. AO26")
    cell_4wk = st.text_input("Cell of Value of 4 WK MA", value="", placeholder="e.g. AP26")

with colBestModel:
    cell_best_model = st.text_input("Cell of Best Model Indicator", value="", placeholder="e.g. AR26")
    cell_item = st.text_input("Cell of Item Code", value="", placeholder="e.g. E26")

with colRight:
    cell_updated_friday = st.text_input(
        "Cell of Updated Last Friday (includes $)", value="", placeholder="e.g. AX$24"
    )
    cell_full_truck = st.text_input(
        "Cell of # Weeks Full Truck (includes $)", value="", placeholder="e.g. AT$22"
    )

if st.button("Generate Formula"):
    formula = (
        f'=IF({cell_best_model}="12 WK MA", {cell_12wk},'
        f' IF({cell_best_model}="8 WK MA", {cell_8wk},'
        f'  IF({cell_best_model}="4 WK MA", {cell_4wk},'
        f'   AVERAGE('
        f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
        f'     MATCH({cell_item}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
        f'     MATCH({cell_updated_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + 1'
        f'    ):'
        f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
        f'     MATCH({cell_item}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
        f'     MATCH({cell_updated_friday}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {cell_full_truck}'
        f'    )'
        f'   )'
        f'  )'
        f' )'
        f')'
    )
    st.text_area("Excel formula (copy-paste):", formula, height=180)



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
