
import streamlit as st
import pandas as pd
from lxml import etree
import io

st.set_page_config(page_title="Excel a XML - WA Prevailing Wage", layout="centered")

st.title("Excel a XML - WA Prevailing Wage (Multi-semana)")

uploaded_file = st.file_uploader("Carga tu archivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if st.button("Generar XML"):
        root = etree.Element("WaPWCPR")

        for end_of_week, group in df.groupby("endOfWeekDate"):
            project_intent = etree.SubElement(root, "projectIntent")
            etree.SubElement(project_intent, "intentId").text = str(group["intentId"].iloc[0])

            payroll = etree.SubElement(root, "payroll")
            payroll_week = etree.SubElement(payroll, "payrollWeek")
            etree.SubElement(payroll_week, "endOfWeekDate").text = str(end_of_week)
            etree.SubElement(payroll_week, "noWorkPerformFlag").text = "false"
            employees = etree.SubElement(payroll_week, "employees")

            for _, row in group.iterrows():
                emp = etree.SubElement(employees, "employee")
                etree.SubElement(emp, "firstName").text = str(row["firstName"])
                etree.SubElement(emp, "midName").text = str(row.get("midName", ""))
                etree.SubElement(emp, "lastName").text = str(row["lastName"])
                etree.SubElement(emp, "ssn").text = str(row["ssn"])
                etree.SubElement(emp, "ethnicity").text = str(row["ethnicity"])
                etree.SubElement(emp, "gender").text = str(row["gender"])
                etree.SubElement(emp, "veteranStatus").text = str(row["veteranStatus"])
                etree.SubElement(emp, "address1").text = str(row["address1"])
                etree.SubElement(emp, "address2").text = str(row["address2"])
                etree.SubElement(emp, "city").text = str(row["city"])
                etree.SubElement(emp, "state").text = str(row["state"])
                etree.SubElement(emp, "zip").text = str(row["zip"])
                etree.SubElement(emp, "grossPay").text = str(row["grossPay"])
                etree.SubElement(emp, "fica").text = str(row["fica"])
                etree.SubElement(emp, "taxWitholding").text = str(row["taxWitholding"])

                deductions = etree.SubElement(emp, "otherDeductions")
                for i in range(1, 3):
                    if pd.notna(row.get(f"deductionName{i}")):
                        d = etree.SubElement(deductions, "otherDeduction")
                        etree.SubElement(d, "deductionName").text = str(row[f"deductionName{i}"])
                        etree.SubElement(d, "deductionHourlyAmt").text = str(row[f"deductionHourlyAmt{i}"])

                thw_container = etree.SubElement(emp, "tradeHoursWages")
                thw = etree.SubElement(thw_container, "tradeHoursWage")
                etree.SubElement(thw, "trade").text = str(row["trade"])
                etree.SubElement(thw, "jobClass").text = str(row["jobClass"])
                etree.SubElement(thw, "tradeNotes").text = str(row.get("tradeNotes", ""))
                etree.SubElement(thw, "county").text = str(row["county"])
                etree.SubElement(thw, "regularHourRateAmt").text = str(row["regularHourRateAmt"])
                etree.SubElement(thw, "overtimeHourRateAmt").text = str(row["overtimeHourRateAmt"])
                etree.SubElement(thw, "doubletimeHourRateAmt").text = str(row["doubletimeHourRateAmt"])
                etree.SubElement(thw, "hourlyPensionRateAmt").text = str(row["hourlyPensionRateAmt"])
                etree.SubElement(thw, "hourlyMedicalAmt").text = str(row["hourlyMedicalAmt"])
                etree.SubElement(thw, "hourlyVacationAmt").text = str(row["hourlyVacationAmt"])
                etree.SubElement(thw, "hourlyHolidayAmt").text = str(row["hourlyHolidayAmt"])
                etree.SubElement(thw, "apprenticeBenefitAmt").text = str(row["apprenticeBenefitAmt"])
                etree.SubElement(thw, "apprenticeFlg").text = str(row["apprenticeFlg"])
                etree.SubElement(thw, "apprenticeId").text = str(row["apprenticeId"])
                etree.SubElement(thw, "apprenticeState").text = str(row["apprenticeState"])
                etree.SubElement(thw, "apprenticeOccpnName").text = str(row["apprenticeOccpnName"])
                etree.SubElement(thw, "apprenticeStepName").text = str(row["apprenticeStepName"])
                etree.SubElement(thw, "apprenticeStepBeginHours").text = str(row["apprenticeStepBeginHours"])
                etree.SubElement(thw, "apprenticeStepEndHours").text = str(row["apprenticeStepEndHours"])

                for prefix in ["regular", "overtime", "doubletime"]:
                    for day in range(1, 8):
                        col = f"{prefix}Day{day}Hours"
                        etree.SubElement(thw, col).text = str(row[col])

                trade_benefits = etree.SubElement(thw, "tradeBenefits")
                for i in range(1, 3):
                    if pd.notna(row.get(f"benefitHourlyName{i}")):
                        b = etree.SubElement(trade_benefits, "tradeBenefit")
                        etree.SubElement(b, "benefitHourlyName").text = str(row[f"benefitHourlyName{i}"])
                        etree.SubElement(b, "benefitHourlyAmt").text = str(row[f"benefitHourlyAmt{i}"])

        xml_output = io.BytesIO()
        tree = etree.ElementTree(root)
        tree.write(xml_output, pretty_print=True, xml_declaration=True, encoding="UTF-8")

        st.download_button(
            label="Descargar XML",
            data=xml_output.getvalue(),
            file_name="wa_prevailing_wage.xml",
            mime="application/xml"
        )
else:
    st.info("Por favor, sube un archivo Excel para continuar.")
