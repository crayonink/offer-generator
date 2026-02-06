import streamlit as st
from docxtpl import DocxTemplate
import pandas as pd
import io

# -------------------------------------------------
# Session State Init
# -------------------------------------------------
if "excel_buffer" not in st.session_state:
    st.session_state.excel_buffer = None

if "word_buffer" not in st.session_state:
    st.session_state.word_buffer = None

# -------------------------------------------------
# Page Config
# -------------------------------------------------
st.set_page_config(page_title="Offer Generator", layout="centered")
st.title("Offer Generator")

# -------------------------------------------------
# STEP 1: COMMERCIAL DETAILS
# -------------------------------------------------
st.subheader("Step 1: Customer & Organization Details")

company_name = st.text_input("Company Name")
company_address = st.text_area("Company Address")

project_name = st.selectbox(
    "Project",
    ["Ladle Preheater", "Preheater Hoods", "Dryers"]
)

fuel_type = st.selectbox("Fuel Type", ["Oil", "Gas"])

poc_designation = st.text_input("Point of Contact (Designation)")
poc_name = st.text_input("POC Name")
mobile_no = st.text_input("Mobile Number")

st.divider()

# -------------------------------------------------
# STEP 2: BURNER SIZE INPUTS
# -------------------------------------------------
st.subheader("Step 2: Burner Size Calculation – Inputs")

input_df = pd.DataFrame(
    {
        "Parameter": [
            "Ti",
            "Tf",
            "Actual Refractory Weight",
            "MG Fuel CV",
            "Time Taken"
        ],
        "Value": [
            650.0,
            1200.0,
            21500.0,
            8500.0,
            1.0
        ],
        "Unit": ["°C", "°C", "Kg", "Kcal/Nm³", "Hours"]
    }
)

edited_df = st.data_editor(
    input_df,
    hide_index=True,
    num_rows="fixed",
    use_container_width=True
)

st.divider()

# -------------------------------------------------
# STEP 3: EXTRACT INPUTS
# -------------------------------------------------
values = dict(zip(edited_df["Parameter"], edited_df["Value"]))

Ti = values["Ti"]
Tf = values["Tf"]
weight = values["Actual Refractory Weight"]
fuel_cv = values["MG Fuel CV"]
time_taken = values["Time Taken"]

# -------------------------------------------------
# STEP 4: CALCULATIONS (EXACT EXCEL MATCH)
# -------------------------------------------------
avg_temp = ((Tf + 200) / 2) - ((Ti + 100) / 2)

firing_rate = weight * 0.25 * avg_temp
heat_load = firing_rate / 0.52
fuel_consumption = heat_load / fuel_cv
calculated_firing_rate = fuel_consumption / time_taken
extra_firing_rate = calculated_firing_rate * 1.1
final_firing_rate = (extra_firing_rate * fuel_cv) / (860 * 1000)
air_qty = fuel_cv * extra_firing_rate * 118 / 100000
cfm = air_qty / 1.7
blower_size_calc = cfm / 114

# -------------------------------------------------
# STEP 5: WORD CONTEXT
# -------------------------------------------------
offer_context = {
    "company_name": company_name,
    "company_address": company_address,
    "project_name": project_name,
    "fuel_type": fuel_type,
    "poc_name": poc_name,
    "poc_designation": poc_designation,
    "mobile_no": mobile_no
}

# -------------------------------------------------
# STEP 6: GENERATE OFFER + EXCEL (3 SHEETS)
# -------------------------------------------------
if st.button("Generate Offer & Calculation Excel"):

    if not company_name or not company_address or not poc_name or not mobile_no:
        st.error("Please fill all mandatory commercial fields")
        st.stop()

    # -------- Sheet 1: Calculation --------
    calculation_df = pd.DataFrame([
        ["Ti", Ti, "°C"],
        ["Tf", Tf, "°C"],
        ["Actual Refractory Weight", weight, "Kg"],
        ["MG Fuel CV", fuel_cv, "Kcal/Nm³"],
        ["Average Temperature to be raised", avg_temp, "°C"],
        ["Time Taken", time_taken, "Hr"],
        ["Firing rate", firing_rate, "Kcal"],
        ["Heat Load", heat_load, "Kcal"],
        ["Fuel Consumption", fuel_consumption, "Nm³"],
        ["Calculated Firing Rate", calculated_firing_rate, "Nm³/hr"],
        ["10% Extra Firing Rate", extra_firing_rate, "Nm³/hr"],
        ["Firing Rate", final_firing_rate, "MW"],
        ["Air Qty", air_qty, "Nm³/hr"],
        ["CFM", cfm, "CFM"],
        ["Blower Size as per Calculation", blower_size_calc, "HP"],
    ], columns=["Parameter", "Value", "Unit"])

    # -------- Sheet 2: VLPH-120T --------
    vlph_df = calculation_df.copy()

    # -------- Sheet 3: Cost Summary --------
    cost_summary_df = pd.DataFrame([
        ["Calculated Firing Rate (Nm³/hr)", calculated_firing_rate],
        ["10% Extra Firing Rate (Nm³/hr)", extra_firing_rate],
        ["Firing Rate (MW)", final_firing_rate],
        ["Air Quantity (Nm³/hr)", air_qty],
        ["CFM", cfm],
        ["Blower Size (HP)", blower_size_calc],
    ], columns=["Description", "Value"])

    st.session_state.excel_buffer = io.BytesIO()

    with pd.ExcelWriter(st.session_state.excel_buffer, engine="xlsxwriter") as writer:
        calculation_df.to_excel(writer, sheet_name="Calculation", index=False)
        vlph_df.to_excel(writer, sheet_name="VLPH-120T", index=False)
        cost_summary_df.to_excel(writer, sheet_name="Cost Summary", index=False)

    st.session_state.excel_buffer.seek(0)

    # -------- Word Offer --------
    doc = DocxTemplate("Offer_Template.docx")
    doc.render(offer_context)

    st.session_state.word_buffer = io.BytesIO()
    doc.save(st.session_state.word_buffer)
    st.session_state.word_buffer.seek(0)

# -------------------------------------------------
# DOWNLOAD BUTTONS
# -------------------------------------------------
if st.session_state.excel_buffer:
    st.download_button(
        "⬇ Download Costing Excel",
        st.session_state.excel_buffer,
        file_name="Costing.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if st.session_state.word_buffer:
    st.download_button(
        "⬇ Download Word Offer",
        st.session_state.word_buffer,
        file_name="Final_Offer.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
