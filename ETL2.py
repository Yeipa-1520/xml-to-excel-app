import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO

st.set_page_config(page_title="XML to Excel Converter", layout="centered")
st.title("📤 Convert XML Payroll to Excel")

uploaded_file = st.file_uploader("Upload your XML file", type="xml")

if uploaded_file:
    try:
        tree = ET.parse(uploaded_file)
        root = tree.getroot()

        summary_general, employees, deductions = [], [], []
        wages_hours, hours_by_day, benefits = [], [], []

        intent_id = root.findtext(".//intentId") or ""

        for payroll_week in root.findall(".//payrollWeek"):
            end_date = payroll_week.findtext("endOfWeekDate") or ""
            no_work = payroll_week.findtext("noWorkPerformFlag") or ""

            for emp in payroll_week.find(".//employees").findall("employee"):
                ssn = emp.findtext("ssn") or ""

                summary_general.append({
                    "intentId": intent_id,
                    "endOfWeekDate": end_date,
                    "noWorkPerformFlag": no_work
                })

                employees.append({
                    "firstName": emp.findtext("firstName") or "",
                    "midName": emp.findtext("midName") or "",
                    "lastName": emp.findtext("lastName") or "",
                    "ssn": ssn,
                    "ethnicity": emp.findtext("ethnicity") or "",
                    "gender": emp.findtext("gender") or "",
                    "veteranStatus": emp.findtext("veteranStatus") or "",
                    "address1": emp.findtext("address1") or "",
                    "address2": emp.findtext("address2") or "",
                    "city": emp.findtext("city") or "",
                    "state": emp.findtext("state") or "",
                    "zip": emp.findtext("zip") or "",
                    "grossPay": emp.findtext("grossPay") or "",
                    "fica": emp.findtext("fica") or "",
                    "taxWitholding": emp.findtext("taxWitholding") or ""
                })

                for ded in emp.findall(".//otherDeduction"):
                    deductions.append({
                        "ssn": ssn,
                        "deductionName": ded.findtext("deductionName") or "",
                        "deductionHourlyAmt": ded.findtext("deductionHourlyAmt") or ""
                    })

                for wage in emp.findall(".//tradeHoursWage"):
                    wages_hours.append({
                        "ssn": ssn,
                        "trade": wage.findtext("trade") or "",
                        "jobClass": wage.findtext("jobClass") or "",
                        "tradeNotes": wage.findtext("tradeNotes") or "",
                        "county": wage.findtext("county") or "",
                        "regularHourRateAmt": wage.findtext("regularHourRateAmt") or "",
                        "overtimeHourRateAmt": wage.findtext("overtimeHourRateAmt") or "",
                        "doubletimeHourRateAmt": wage.findtext("doubletimeHourRateAmt") or "",
                        "hourlyPensionRateAmt": wage.findtext("hourlyPensionRateAmt") or "",
                        "hourlyMedicalAmt": wage.findtext("hourlyMedicalAmt") or "",
                        "hourlyVacationAmt": wage.findtext("hourlyVacationAmt") or "",
                        "hourlyHolidayAmt": wage.findtext("hourlyHolidayAmt") or "",
                        "apprenticeBenefitAmt": wage.findtext("apprenticeBenefitAmt") or "",
                        "apprenticeFlg": wage.findtext("apprenticeFlg") or "",
                        "apprenticeId": wage.findtext("apprenticeId") or "",
                        "apprenticeState": wage.findtext("apprenticeState") or "",
                        "apprenticeOccpnName": wage.findtext("apprenticeOccpnName") or "",
                        "apprenticeStepName": wage.findtext("apprenticeStepName") or "",
                        "apprenticeStepBeginHours": wage.findtext("apprenticeStepBeginHours") or "",
                        "apprenticeStepEndHours": wage.findtext("apprenticeStepEndHours") or ""
                    })

                    hours = {"ssn": ssn}
                    for i in range(1, 8):
                        hours[f"regularDay{i}Hours"] = wage.findtext(f"regularDay{i}Hours") or ""
                        hours[f"overtimeDay{i}Hours"] = wage.findtext(f"overtimeDay{i}Hours") or ""
                        hours[f"doubletimeDay{i}Hours"] = wage.findtext(f"doubletimeDay{i}Hours") or ""
                    hours_by_day.append(hours)

                    for ben in wage.findall(".//tradeBenefit"):
                        benefits.append({
                            "ssn": ssn,
                            "benefitHourlyName": ben.findtext("benefitHourlyName") or "",
                            "benefitHourlyAmt": ben.findtext("benefitHourlyAmt") or ""
                        })

        dfs = {
            "SummaryGeneral": pd.DataFrame(summary_general),
            "Employees": pd.DataFrame(employees),
            "Deductions": pd.DataFrame(deductions),
            "WagesHours": pd.DataFrame(wages_hours),
            "HoursByDay": pd.DataFrame(hours_by_day),
            "Benefits": pd.DataFrame(benefits)
        }

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for name, df in dfs.items():
                df.to_excel(writer, sheet_name=name, index=False)
        output.seek(0)

        st.success("✅ File processed successfully! You can now download the Excel file.")
        st.download_button(
            label="📥 Download Excel File",
            data=output,
            file_name="payroll_from_xml.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
    