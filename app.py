import streamlit as st
from docxtpl import DocxTemplate
import pandas as pd
import io
import math

# -------------------------------------------------
# CONFIG
# -------------------------------------------------
MARKUP = 1.8

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

input_df = pd.DataFrame({
    "Parameter": ["Ti", "Tf", "Actual Refractory Weight", "MG Fuel CV", "Time Taken"],
    "Value": [650.0, 1200.0, 21500.0, 8500.0, 1.0],
    "Unit": ["°C", "°C", "Kg", "Kcal/Nm³", "Hours"]
})

edited_df = st.data_editor(
    input_df,
    hide_index=True,
    num_rows="fixed",
    use_container_width=True
)

values = dict(zip(edited_df["Parameter"], edited_df["Value"]))

Ti = values["Ti"]
Tf = values["Tf"]
weight = values["Actual Refractory Weight"]
fuel_cv = values["MG Fuel CV"]
time_taken = values["Time Taken"]

# -------------------------------------------------
# STEP 3: CALCULATIONS
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
# WORD CONTEXT
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
# GENERATE FILES
# -------------------------------------------------
if st.button("Generate Offer & Calculation Excel"):

    if not company_name or not company_address or not poc_name or not mobile_no:
        st.error("Please fill all mandatory commercial fields")
        st.stop()

    # -----------------------------
    # CALCULATION SHEET
    # -----------------------------
    calculation_df = pd.DataFrame([
        ["Ti", Ti, "°C"],
        ["Tf", Tf, "°C"],
        ["Actual Refractory Weight", weight, "Kg"],
        ["MG Fuel CV", fuel_cv, "Kcal/Nm³"],
        ["Average Temperature Rise", avg_temp, "°C"],
        ["Time Taken", time_taken, "Hr"],
        ["Firing Rate", firing_rate, "Kcal"],
        ["Heat Load", heat_load, "Kcal"],
        ["Fuel Consumption", fuel_consumption, "Nm³"],
        ["Final Firing Rate", final_firing_rate, "MW"],
        ["Air Qty", air_qty, "Nm³/hr"],
        ["Blower Size", blower_size_calc, "HP"],
    ], columns=["Parameter", "Value", "Unit"])

    # -----------------------------
    # BOM DATA
    # -----------------------------
    bom_columns = [
        "S. No.", "MEDIA", "ITEM NAME", "Data Sheet No / Reference",
        "QTY", "UNIT", "MAKE", "BASIC", "TOTAL"
    ]

    bom_data = [
        [1,"COMB AIR","COMPENSATOR","300 NB F 150 #",1,"No.","ENCON",8000,8000],
        [2,"COMB AIR","PRESSURE GAUGE WITH TNV","RANGE- 0-2000 mm WC, Dial-4\"",1,"No.","WIKA",4000,4000],
        [3,"COMB AIR","PRESSURE SWITCH LOW (Set Pt -L)","RANGE- 5-150 mBAR",1,"No.","SWITZER",6500,6500],
        [6,"COMB AIR","MOTERIZED CONTROL VALVE","250 NB, FLOW- 4000 Nm3/hr",1,"No.","CAIR",80000,80000],
        [7,"COMB AIR","BUTTERFLY VALVE","300 NB",1,"No","AUDCO/ L&T/ LEADER",16950,16950],
        [8,"COMB AIR","ROTARY JOINT","300 NB",1,"No","KRATOS/ENCON",50000,50000],
        [9,"COMB AIR","BALL VALVE (Pilot Burner)","20 NB",1,"No.","AUDCO/ L&T/ LEADER",1600,1600],
        [10,"COMB AIR","BALL VALVE (UV - LINE)","15 NB",1,"No.","AUDCO/ L&T/ LEADER",1500,1500],
        [11,"COMB AIR","FLEXIBLE HOSE (Pilot Burner)","20 NB, 1500 mm LONG",1,"No.","BIL/ FLEXIBLE",1200,1200],
        [12,"COMB AIR","FLEXIBLE HOSE (UV - LINE)","15 NB, 1500 mm LONG",1,"No.","BIL/ FLEXIBLE",1000,1000],
        [13,"NG PILOT LINE","BALL VALVE","20 NB",2,"No.","AUDCO/ L&T/ LEADER",1600,3200],
        [14,"NG PILOT LINE","PRESSURE GAUGE WITH NV","0-1600 mm WC, DIAL: 4\"",1,"No.","WIKA",4000,4000],
        [15,"NG PILOT LINE","PRESSURE SWITCH HIGH + LOW","",2,"No.","SWITZER",12000,24000],
        [16,"NG PILOT LINE","SOLENOID VALVE","15 NB",1,"No.","MADAS",6000,6000],
        [17,"NG PILOT LINE","PRESSURE REGULATING VALVE","15 NB",1,"No.","MADAS",8000,8000],
        [18,"NG PILOT LINE","FLEXIBALE HOSE PIPE","15 NB - 1500 mm LONG",1,"No.","BIL/FLEXIBLE",1200,1200],
        [19,"MG LINE","NG GAS TRAIN","FLOW: 400 Nm3/hr",1,"No.","MADAS",295200,295200],
        [24,"MG LINE","AGR","80 NB",1,"No.","MADAS",48250,48250],
        [29,"MISC ITEMS","THERMOCOUPLE","R TYPE, 505 mm, 1200C",1,"No.","TEMPSENS",32000,32000],
        [30,"MISC ITEMS","COMPENSATING LEAD","FOR R TYPE TC",1,"Roll.","TEMPSENS",5000,5000],
        [32,"MISC ITEMS","LIMIT SWITCHES","",2,"Nos.","BCH",2300,4600],
        [33,"MISC ITEMS","CONTROL PANEL","MCC",1,"No.","ENCON",150000,150000],
        [34,"MISC ITEMS","HYDRAULIC POWER PACK & CYLINDER","10 HP, 1500 mm",1,"No.","VARITECH",310000,310000],
        [35,"MISC ITEMS","CABLE FOR IGNITION TRANSFORMER","",20,"m","ENCON",200,4000],
        [36,"MISC ITEMS","TEMPERATURE TRANSMITTER","",1,"No.","HONEYWELL",13000,13000],
        [37,"MISC ITEMS","P.PID","",1,"No","HONEYWELL",8000,8000],
        [38,"MISC ITEMS","RATIO CONTROLLER","",1,"No.","HONEYWELL",55000,55000],
    ]

    bom_df = pd.DataFrame(bom_data, columns=bom_columns)

    bought_out_cost = bom_df[bom_df["ITEM NAME"] != "RATIO CONTROLLER"]["TOTAL"].sum()
    bought_out_sell = bought_out_cost * MARKUP

    # -----------------------------
    # ENCON ITEMS
    # -----------------------------
    encon_df = pd.DataFrame([
        [1,"MISC ITEMS","ENCON MG BURNER WITH B. BLOCK","NATURAL GAS FLOW: 440 Nm3/hr G7A",1,"No.","ENCON",118000,118000],
        [8,"MISC ITEMS","COMBUSTION AIR BLOWER",'25 HP, 28" WC, 5100 Nm3/hr',1,"No.","ENCON",195000,195000],
        [9,"MISC ITEMS","PILOT BURNER","10 KW",1,"No.","ENCON",12000,12000],
        [10,"MISC ITEMS","IGNITION TRANSFORMER","",1,"No.","COFI/DANFOSS",5500,5500],
        [11,"MISC ITEMS","SEQUENCE CONTROLLER","",1,"No.","LINEAR",10000,10000],
        [12,"MISC ITEMS","UV SENSOR WITH AIR JACKET","",1,"No.","LINEAR",13000,13000],
    ], columns=bom_columns)

    inhouse_sell = encon_df["TOTAL"].sum()
    inhouse_cost = inhouse_sell / MARKUP

    # -----------------------------
    # BUILD VLPH-120T WITH SEPARATION
    # -----------------------------
    blank = pd.DataFrame([[""] * len(bom_columns)], columns=bom_columns)

    vlph_df = pd.concat([
        bom_df,
        blank, blank,
        pd.DataFrame([["", "", "BOUGHT OUT ITEMS", "", "", "", "", "", bought_out_cost]], columns=bom_columns),
        pd.DataFrame([["", "", f"TOTAL x {MARKUP}", "", "", "", "", "", bought_out_sell]], columns=bom_columns),
        blank, blank,
        encon_df,
        blank,
        pd.DataFrame([["", "", "ENCON ITEMS", "", "", "", "", "", inhouse_sell]], columns=bom_columns),
    ], ignore_index=True)

    # -----------------------------
    # COST SUMMARY (VERTICAL ONLY)
    # -----------------------------
    unit_cost = bought_out_cost + inhouse_cost
    unit_sell = bought_out_sell + inhouse_sell
    designing_10 = unit_sell * 0.10
    negotiation_10 = unit_sell * 0.10
    total_price = unit_sell + designing_10 + negotiation_10
    usd_price = total_price / 85

    cost_summary_df = pd.DataFrame([[
        1, "Vertical Ladle Preheater",
        bought_out_cost, bought_out_sell,
        inhouse_cost, inhouse_sell,
        unit_cost, unit_sell,
        designing_10, negotiation_10,
        1, total_price, MARKUP, usd_price
    ]], columns=[
        "S.No.", "Item Description",
        "Bought Out Cost Price", "Bought Out Sell Price",
        "Inhouse Cost Price", "Inhouse Sell Price",
        "Unit Cost Price", "Unit Sell Price",
        "10% Designing", "10% Negotiation",
        "Qty/Set", "Total Price", "Markup", "USD"
    ])

    # -----------------------------
    # WRITE EXCEL
    # -----------------------------
    st.session_state.excel_buffer = io.BytesIO()
    with pd.ExcelWriter(st.session_state.excel_buffer, engine="xlsxwriter") as writer:
        calculation_df.to_excel(writer, sheet_name="Calculation", index=False)
        vlph_df.to_excel(writer, sheet_name="VLPH-120T", index=False)
        cost_summary_df.to_excel(writer, sheet_name="Cost Summary", index=False)

    st.session_state.excel_buffer.seek(0)

    # -----------------------------
    # WORD OFFER
    # -----------------------------
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
