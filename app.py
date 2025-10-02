import streamlit as st

st.title("Excel Formula Replacer Tool")

AR26 = st.text_input("Replace {AR}{26} (Best model indicator):", "AR26")
AN26 = st.text_input("Replace {AN}{26} (Value of 12 WK MA):", "AN26")
AO26 = st.text_input("Replace {AO}{26} (Value of 8 WK MA):", "AO26")
AP26 = st.text_input("Replace {AP}{26} (Value of 4 WK MA):", "AP26")
E26 = st.text_input("Replace {E}{26} (Item code):", "E26")
AX24 = st.text_input("Replace {AX}${24} (Updated last Friday):", "AX$24")
AT22 = st.text_input("Replace {AT}${22} (# Weeks full truck):", "AT$22")

if st.button("Generate Formula"):
    formula = (
        f'=IF({AR26}="12 WK MA", {AN26},'
        f' IF({AR26}="8 WK MA", {AO26},'
        f'  IF({AR26}="4 WK MA", {AP26},'
        f'   AVERAGE('
        f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
        f'     MATCH({E26}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
        f'     MATCH({AX24}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + 1'
        f'    ):'
        f'    INDEX(\'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$AN,'
        f'     MATCH({E26}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$A:$A, 0),'
        f'     MATCH({AX24}, \'[Model_Suggestion.xlsx]Overall Suggestion\'!$1:$1, 0) + {AT22}'
        f'    )'
        f'   )'
        f'  )'
        f' )'
        f')'
    )
    st.text_area("Excel formula (copy-paste):", formula, height=150)
