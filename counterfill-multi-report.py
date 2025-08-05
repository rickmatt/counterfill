# script to create TrueView reports
# Copyright 2025, Secure340B
# Author: Rick Matthews

import numbers
import mysql.connector
import xlsxwriter
import time
import decimal
import datetime
import sys
from icecream import ic
import calendar
import numpy

starttime = datetime.datetime.now()
# report_period looks like "2025-03"
report_period = "2025-06"
# report_label looks like "2025-3"
report_label = "2025-6"
report_year = int(report_period.split("-")[0])
report_month = int(report_period.split("-")[1])
this_month = report_month
ic(report_period)
ic(report_year)
ic(report_month)
status = "340B"

# get report range as YYYY-MM-DD
last_day = calendar.monthrange(report_year, report_month)[1]
report_end_date = f"{report_year}-{report_month:02d}-{last_day:02d}"
tpa_audit_end_date = report_period.split("-")[0] + "-" + report_period.split("-")[1] + "-" + "20"
if report_month == 1:
    report_start_date = str(report_year - 1) + "-11-01"
elif report_month == 2:
    report_start_date = str(report_year - 1) + "-12-01"
else:
    report_start_date = str(report_year) + "-" + str(report_month - 2).zfill(2) + "-01"
ic(report_start_date)
ic(report_end_date)
ic(tpa_audit_end_date)





# report_id = 1137
report_id = input("enter the report id number: ")
print(report_id)
print(type(report_id))

def get_indicator(ndc11):
    # get indicator
    # print("getting indicator for", ndc11)
    gi_query = ""
    gi_indicator_result = ""
    gi_indicator = ""
    gi_query = """SELECT * FROM drug_catalog WHERE ndc11 = %s LIMIT 1;"""
    cursor.execute(gi_query, (ndc11,))
    gi_indicator_result = cursor.fetchone()
    if gi_indicator_result:
        gi_indicator = gi_indicator_result["indicator"]
    else:
        gi_indicator = ""
    return(gi_indicator)

# establishing the connection
conn = mysql.connector.connect(
    user="root", password="root1234", host="127.0.0.1", database="dev_secure340b"
)

# Creating a cursor object using the cursor() method
cursor = conn.cursor(dictionary=True, buffered=True)

# get the list of reports
sql = """SELECT * 
    FROM report_queue 
    WHERE id = %s  and report_type = 'Counterfill' ORDER BY id ASC"""
record = (report_id,)
cursor.execute(sql, record)
report = cursor.fetchone()
ic(report)
report_name = report["salesforce_report_name"]
pharmacy_name = report["pharmacy"]
pharmacy_state = report["cp_state"]
ic(pharmacy_name)
ic(pharmacy_state)

# get the list of report_identifiers
sql = """SELECT report_identifier FROM counterfill_meta WHERE counterfill_name = %s;"""
record = (pharmacy_name,)
cursor.execute(sql, record)
report_identifiers = cursor.fetchall()
ic(report_identifiers)
qms_list = []


print(__file__)

# create workbook
file_prefix = report_name + "-" + report_period.replace("-", "")
timestr = "-" + time.strftime("%Y%m%d-%H%M")
path = "CounterfillReports/"
filename = file_prefix + timestr + ".xlsx"
path_filename = path + filename
workbook = xlsxwriter.Workbook(path_filename)

# excel format statements
date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
money = workbook.add_format({"num_format": "$#,##0.00"})
pct_format = workbook.add_format({"num_format": "0.00%", "border": 1})
pct_format2 = workbook.add_format({"num_format": "0.00%"})
title_format = workbook.add_format(
    {"border": 1, "bold": True, "bg_color": "#E0FFFF"}
)
format1 = workbook.add_format({"border": 1, "num_format": "$#,##0.00"})
format2 = workbook.add_format({"border": 1, "num_format": "0"})
yippee = workbook.add_format(
    {"border": 1, "num_format": "$#,##0.00", "bg_color": "#7FFF00"}
)

f_title = workbook.add_format({"border": 0, "bold": True, "font_size": 14})
f_title2 = workbook.add_format({"border": 1, "font_size": 14})


# create qc tab
qctab = workbook.add_worksheet("QC")
qctab.set_column(0, 0, 30)
qctab.set_column(1, 1, 80)
qctab.set_column(2, 5, 30)
qcrow=1
for field in report:
    if report[field]:
        qctab.write(qcrow, 0, field)
        qctab.write(qcrow, 1, report[field])
        qcrow += 1


# hide qc tab
# qctab.hide()

# create summary tab
summarytab = workbook.add_worksheet("Summary")
summarytab.set_column(0,1, 30)
summarytab.hide_gridlines(2)
summarytab.insert_image(
        "A1",
        "secure340b-logo-281.png",
        {
            "x_scale": 0.5,
            "y_scale": 0.5,
            "x_offset": 15,
        },
    )
summarytab.write("B1", report["salesforce_report_name"], f_title)
summarytab.write("B2", "Secure340B Counterfill Report", f_title)
summarytab.write("B3", report_label, f_title)

# create pharmacy dispensing data tab
pddtab = workbook.add_worksheet("Pharmacy Dispensing Data")
pddtab.set_tab_color("81A3A7")
pddtab.set_column(0, 27, 15)
pddtab.freeze_panes(1, 0)
pdd_row = 0
pdd_headers = ['Rx Number',
'Fill Number',
'Rx + Fill',
'Covered Entity',
'TPA',
'Date of Service',
'Prescriber Name',
'Prescriber NPI',
'NDC',
'NDC Description',
'Indicator',
'Quantity Dispense',
'Total Paid Amount',
'Acquistion Cost',
'Retail Margin',
'Household',
'Patient DOB',
'Payor',
'BIN',
'PCN',
'Group',
'Days Supply',
'Dispense Fee',
'True Margin',
'Status',
'Qualification Date',
'NDC Replenished by 340B?',
'Manufacturer',
'Medicaid']
for idx, header in enumerate(pdd_headers):
    pddtab.write(pdd_row, idx, header, title_format)
pdd_row += 1

is340b_count = 0

pdd_query_starttime = datetime.datetime.now()
pdd_query = """SELECT * FROM counterfill_claims 
    WHERE pharmacy_name = %s
    AND fill_date BETWEEN %s AND %s;"""
pdd_inputs = (pharmacy_name, report_start_date, report_end_date)
cursor.execute(pdd_query, pdd_inputs)
pdd_claims = cursor.fetchall()
pdd_query_endtime = datetime.datetime.now()
pdd_query_duration = pdd_query_endtime - pdd_query_starttime
ic(pdd_query_duration)
print("looking at 340b_claims")
no_ndc = 0
no_total_payment = 0
for claim in pdd_claims:
    # if there is no ndc11 or if total_payment is 0, skip this claim
    if claim["ndc11"] == "":
        no_ndc += 1
        continue
    if claim["ndc11"] == None:
        no_ndc += 1
        continue
    if claim["ndc11"] == "00000000000":
        no_ndc += 1
        continue
    if claim["total_payment"] == 0:
        no_total_payment += 1
        continue

    is340b_query = """SELECT * FROM 340b_claims
        WHERE rx_number = %s
        AND fill_number = %s
        AND ndc = %s
        AND fill_date = %s
        ORDER BY id DESC LIMIT 1;"""
    is340b_inputs = (claim["rx_number"], claim["fill_number"], claim["ndc11"], claim["fill_date"])
    cursor.execute(is340b_query, is340b_inputs)
    is340b_results = cursor.fetchone()
    if is340b_results:
        is340b_count += 1
        # print("340b claim")
        # ic(is340b_results)
        covered_entity = is340b_results["covered_entity"]
        tpa = is340b_results["tpa"]
        disp_fee = is340b_results["disp_fee"]
        status = "340B"
        qual_date = is340b_results["bill_date"]
        if claim["rx_fill_concat"] == is340b_results["rx_fill_concat"]:
            replenished = "YES"
        else:
            replenished = "NO"
    else:
        covered_entity = "Not 340B"
        tpa = "Not 340B"
        disp_fee = 0
        status = "RETAIL"
        qual_date = "Not 340B"
        replenished = "N/A"
    pdd_col = 0
    pddtab.write(pdd_row, pdd_col, claim["rx_number"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["fill_number"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["rx_fill_concat"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, covered_entity)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, tpa)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["fill_date"], date_format)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["prescriber_name"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["prescriber_npi"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["ndc11"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["drug_name"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["indicator"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["qty_disp"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["total_payment"], money)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["drug_cost"], money)
    pdd_col += 1
    retail_margin = float(claim["total_payment"]) - float(claim["drug_cost"])
    pddtab.write(pdd_row, pdd_col, retail_margin, money)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["pat_address"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["pat_dob"], date_format)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["plan_name"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["bin"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["pcn"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["rx_group"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["days_supply"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, disp_fee, money)
    pdd_col += 1
    true_margin = max(retail_margin, disp_fee)
    pddtab.write(pdd_row, pdd_col, true_margin, money)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, status)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, qual_date, date_format)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, replenished)
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["manufacturer"])
    pdd_col += 1
    pddtab.write(pdd_row, pdd_col, claim["medicaid"])
    pdd_col += 1

    pdd_row += 1
pddtab.autofilter(0, 0, pdd_row, len(pdd_headers)-1)
pdd_endtime = datetime.datetime.now()
pdd_duration = pdd_endtime - pdd_query_endtime
ic(pdd_duration)


# get ndcs for this report period
# ndc_query = """SELECT DISTINCT ndc11 FROM counterfill_claims
#     WHERE pharmacy_name = %s
#     AND year(fill_date) = %s
#     AND month(fill_date) = %s;"""
# ndc_inputs = (pharmacy_name, report_year, report_month)
# cursor.execute(ndc_query, ndc_inputs)
# ndcs = cursor.fetchall()
# ndc_row = 3
# for ndc in ndcs:
#     ndccol = 0
#     ndctab.write(ndc_row, ndccol, ndc["ndc11"])
#     ndccol += 1
#     # get drug name
#     drug_query = """SELECT * FROM drug_catalog WHERE ndc11 = %s LIMIT 1;"""
#     cursor.execute(drug_query, (ndc["ndc11"],))
#     drug = cursor.fetchone()
#     try:
#         ndctab.write(ndc_row, ndccol, drug["description"])
#     except:
#         ndctab.write(ndc_row, ndccol, "")
#     ndccol += 1

#     ndc_row += 1

# create RXs for review tab
rxreviewtab = workbook.add_worksheet("RXs for Review")
rxreviewtab.write_url('A1',  "internal:'Summary'!A1", string="Return to Summary")

# create Qualified Prescribers tab
qptab = workbook.add_worksheet("Qualified Prescribers")
qptab.set_tab_color("81A3A7")
qptab.write_url('A1',  "internal:'Summary'!A1", string="Return to Summary")
qptab.set_column(0, 6, 25)
qprow = 2
qpheaders = ["Prescriber Number", "Prescriber Name", "340B Match?", "Percent Qualified", "Prescriber CE", "Doctor Found at Multiple CEs?", "Possible CEs"]
for idx, header in enumerate(qpheaders):
    qptab.write(qprow, idx, header, title_format)
qprow += 1

doctors = []
for report in report_identifiers:
    doctor_query = f"""SELECT DISTINCT prescriber_npi FROM 340b_claims
        WHERE report_identifier = %s and fill_date BETWEEN %s AND %s;"""
    doctor_inputs = (report['report_identifier'], report_start_date, report_end_date)
    cursor.execute(doctor_query, doctor_inputs)
    physicians = cursor.fetchall()
    for physician in physicians:
        doctors.append(physician['prescriber_npi'])
    # ic(physicians, len(physicians))
# ic(doctors, len(doctors))
doctors = list(set(doctors))
ic(doctors, len(doctors))
qual_npi_list = []
qual_doctor_list = []
for doctor in doctors:
    pharm_count_query = """SELECT COUNT(*) 
        FROM counterfill_claims 
        WHERE prescriber_npi = %s
        AND indicator = 'B'
        AND pharmacy_name = %s
        AND fill_date BETWEEN %s AND %s;"""
    pharm_count_inputs = (doctor, pharmacy_name, report_start_date, report_end_date)
    cursor.execute(pharm_count_query, pharm_count_inputs)
    pharm_count = cursor.fetchone()
    # ic(pharm_count)
    b_count_query = """SELECT COUNT(DISTINCT rx_fill_concat)
        FROM 340b_claims
        WHERE prescriber_npi = %s
        AND indicator = 'B'
        AND fill_date BETWEEN %s AND %s;"""
    b_count_inputs = (doctor, report_start_date, report_end_date)
    cursor.execute(b_count_query, b_count_inputs)
    b_count = cursor.fetchone()
    ic(b_count)
    try:
        qp_pct = float(b_count["COUNT(DISTINCT rx_fill_concat)"]) / float(pharm_count["COUNT(*)"])
    except:
        qp_pct = 0
    if qp_pct < 0.05:
        continue
    if qp_pct > 1:
        qp_pct = 1
    qpcol = 0
    qptab.write(qprow, qpcol, doctor)
    qpcol += 1
    qpdoc_query = """SELECT covered_entity, count(*) 
        FROM 340b_claims 
        WHERE prescriber_npi = %s 
        AND fill_date between %s AND %s
        AND report_identifier IN (SELECT report_identifier FROM counterfill_meta WHERE counterfill_name = %s)
        GROUP BY covered_entity 
        ORDER BY count(*) DESC;"""
    cursor.execute(qpdoc_query, (doctor, report_start_date, report_end_date, pharmacy_name))
    qpdoc = cursor.fetchall()

    ce_count_query = """SELECT DISTINCT(covered_entity) FROM 340b_claims 
        WHERE prescriber_npi = %s 
        AND fill_date between %s AND %s;"""
    ce_count_inputs = (doctor, report_start_date, report_end_date)
    cursor.execute(ce_count_query, ce_count_inputs)
    ce_count = cursor.fetchall()

    ic(ce_count, len(ce_count))
    qp2doc_query = """SELECT * FROM counterfill_claims WHERE prescriber_npi = %s ORDER BY id DESC LIMIT 1;"""
    cursor.execute(qp2doc_query, (doctor,))
    qp2doc = cursor.fetchone()
    qptab.write(qprow, qpcol, qp2doc["prescriber_name"])
    qpcol += 1
    qptab.write(qprow, qpcol, "YES")
    qpcol += 1
    qptab.write(qprow, qpcol, qp_pct, pct_format2)
    qpcol += 1
    qptab.write(qprow, qpcol, qpdoc[0]["covered_entity"])
    qpcol += 1
    if len(ce_count) == 1:
        qptab.write(qprow, qpcol, "NO")
    else:
        qptab.write(qprow, qpcol, "YES")
    qual_npi_list.append(doctor)
    qual_doctor_list.append(doctor + '--' + qp2doc["prescriber_name"])
    ces = ""
    ces_count = 0
    for ce1 in ce_count:
        if ces_count >= 1:
            ces += "; "
        ces += ce1["covered_entity"]
        ces_count += 1
    print(ces)
    qpcol += 1
    qptab.write(qprow, qpcol, ces)



    qprow += 1
ic(qual_npi_list)
ic(qual_doctor_list)

# create Qualified Manufacturers tab
qmtab = workbook.add_worksheet("Qualified Manufacturers")
qmtab.set_tab_color("81A3A7")
qmtab.write_url('A1',  "internal:'Summary'!A1", string="Return to Summary")
qmtab.set_column(0, 5, 25)
qmrow = 2
qmheaders = ["Manufacturer", "Times Qualified", "Replenished 340B", "CP-CE"]
for idx, header in enumerate(qmheaders):
    qmtab.write(qmrow, idx, header, title_format)
qmrow += 1

for report in report_identifiers:
    qm_query = """SELECT a.manufacturer, count(*) FROM manuf_exclusions as a
        RIGHT JOIN 340b_claims AS b
        ON a.ndc11 = b.ndc
        WHERE b.report_identifier = %s
        AND b.fill_date BETWEEN %s AND %s
        GROUP BY a.manufacturer ORDER BY a.manufacturer ASC;"""
    qm_inputs = (report["report_identifier"], report_start_date, report_end_date)
    cursor.execute(qm_query, qm_inputs)
    qms = cursor.fetchall()
    ce_query = """SELECT covered_entity FROM counterfill_meta WHERE report_identifier = %s;"""
    ce_inputs = (report["report_identifier"],)
    cursor.execute(ce_query, ce_inputs)
    report["covered_entity"] = cursor.fetchone()["covered_entity"]
    ic(report)
    ic(qms)
    for qm in qms:
        qmcol = 0
        manuf_concat = ""
        manuf_concat = qm["manufacturer"]
        if qm["manufacturer"] == None:
            qm["manufacturer"] = "Not restricted manufacturer"
        else:
            qms_list.append(manuf_concat)
        qmtab.write(qmrow, qmcol, qm["manufacturer"])
        qmcol += 1
        qmtab.write(qmrow, qmcol, qm["count(*)"])
        qmcol += 1
        qmtab.write(qmrow, qmcol, "YES")
        qmcol += 1
        qmtab.write(qmrow, qmcol, report["covered_entity"])
        qmrow += 1

ic(manuf_concat)

# create TPA Audit RXs tab
# clean up counterfill_audit_rxs table if this report is being rerun
cursor.execute("DELETE FROM counterfill_audit_rxs WHERE pharmacy = %s AND report_period = %s;", (pharmacy_name, report_period))
print("creating TPA Audit RXs tab")
tpa_audit_tab = workbook.add_worksheet("TPA Rx Review")
tpa_audit_tab.set_tab_color("81A3A7")
tpa_audit_tab.write_url('A1',  "internal:'Summary'!A1", string="Return to Summary")

# get historical list of audit tab entries
tpa_query = """SELECT * FROM counterfill_audit_rxs WHERE pharmacy = %s AND report_period < %s;"""
tpa_inputs = (pharmacy_name, report_period)
cursor.execute(tpa_query, tpa_inputs)
tpa_audit_history = cursor.fetchall()
tpa_audit_history_list = []
for tpa_audit_item in tpa_audit_history:
    tpa_audit_history_list.append(tpa_audit_item["rx_fill_num"])
ic(tpa_audit_history_list)
ic(len(tpa_audit_history))

tpa_row = 2
tpa_headers = [
    "Rx Number",
    "Fill Number",
    "Fill Date",
    "NDC",
    "Description",
    "Est. Dispense Fee",
    "Est. Paid to CE",
    "Est. True Margin",
    "Pharmacy Impact",
    "Rx Number ever 340B?",
    "Qualifying Prescriber Name",
    "Qualifying Manufacturer",
    "Potential Covered Entity",
    "Prescriber NPI"
]
tpa_audit_tab.set_column(0, len(tpa_headers)-1, 20)
for idx, header in enumerate(tpa_headers):
    tpa_audit_tab.write(tpa_row, idx, header, title_format)
tpa_row += 1
# ninetydaysago = datetime.datetime.now() - datetime.timedelta(days=90)
# ninetydaysago = ninetydaysago.strftime("%Y-%m-%d")

ic(qms_list)
audit_skips = 0
for doctor in qual_npi_list:
    ic(doctor)
    prescription_query = """SELECT * FROM counterfill_claims where prescriber_npi = %s
        AND pharmacy_name = %s
        AND fill_date BETWEEN %s AND %s
        AND indicator = 'B'
        AND medicaid = 'NO';"""
    prescription_inputs = (doctor, pharmacy_name, report_start_date, tpa_audit_end_date)
    cursor.execute(prescription_query, prescription_inputs)
    prescriptions = cursor.fetchall()
    ic(len(prescriptions))

    for prescription in prescriptions:
        if prescription["rx_fill_concat"] in tpa_audit_history_list:
            # audit_query = """SELECT * FROM counterfill_audit_rxs WHERE rx_fill_num = %s AND pharmacy = %s;"""
            # audit_inputs = (prescription["rx_fill_concat"], pharmacy_name)
            # cursor.execute(audit_query, audit_inputs)
            # audit_results = cursor.fetchone()
            # first_audit_period = audit_results["report_period"]
            audit_skips += 1
            continue
        else:
            first_audit_period = ""
            record_sql = """INSERT INTO counterfill_audit_rxs
            (pharmacy, rx_fill_num, ndc11, report_period)
            VALUES (%s, %s, %s, %s);"""
            record_inputs = (pharmacy_name, prescription["rx_fill_concat"], prescription["ndc11"], report_period)
            ic(record_inputs)
            cursor.execute(record_sql, record_inputs)
            conn.commit()
        if prescription["manufacturer"] not in qms_list:
            print(f"{prescription["manufacturer"]} not in manuf list")
            continue
        q_for_340b = """SELECT * FROM 340b_claims WHERE rx_fill_concat = %s AND prescriber_npi = %s;"""
        # ic(doctor)
        q_for_340b_inputs = (prescription["rx_fill_concat"], doctor)
        cursor.execute(q_for_340b, q_for_340b_inputs)
        q_for_340b_results = cursor.fetchone()
        # ic(q_for_340b_results)
        # don't include this claim in list if it is already 340B Qualified.
        if q_for_340b_results:
            print("already 340B")
            continue
        #if prescription["est_disp_fee"] < 15:
        #    continue
        # if float(prescription["total_payment"])-float(prescription["est_disp_fee"]) < 25:
        #     continue
        # if float(prescription["est_disp_fee"])-float(prescription["retail_margin"]) < 0:
        #     continue
        pharm_paid_to_ce = 0
        pharm_paid_to_ce = float(prescription["total_payment"])-float(prescription["est_disp_fee"])
        if pharm_paid_to_ce < 50:
            continue
        pharmacy_impact = 0
        pharmacy_impact = float(prescription["est_disp_fee"])-float(prescription["retail_margin"])
        if pharmacy_impact < 20:
            continue
        col = 0
        tpa_audit_tab.write(tpa_row, col, prescription["rx_number"])
        col += 1
        tpa_audit_tab.write(tpa_row, col, prescription["fill_number"])
        col += 1
        tpa_audit_tab.write(tpa_row, col, prescription["fill_date"], date_format)
        col += 1
        tpa_audit_tab.write(tpa_row, col, prescription["ndc11"])
        col += 1
        tpa_audit_tab.write(tpa_row, col, prescription["drug_name"])
        col += 1
        
        tpa_audit_tab.write(tpa_row, col, prescription["est_disp_fee"], money)
        col += 1
        tpa_audit_tab.write(tpa_row, col, pharm_paid_to_ce, money)
        col += 1
        tpa_audit_tab.write(tpa_row, col, prescription["retail_margin"], money)
        col += 1
        tpa_audit_tab.write(tpa_row, col, pharmacy_impact, money)

        col += 1
        # ic(q_for_340b_results)
        ever_query = """SELECT * FROM 340b_claims WHERE rx_number = %s AND prescriber_npi = %s ORDER BY fill_date DESC LIMIT 1;"""
        ever_inputs = (prescription["rx_number"], doctor)
        cursor.execute(ever_query, ever_inputs)
        ever_results = cursor.fetchone()
        ic(ever_results)
        if ever_results:
            tpa_audit_tab.write(tpa_row, col, ever_results["fill_date"], date_format)
            ce = ever_results["covered_entity"]
        else:
            tpa_audit_tab.write(tpa_row, col, "NO")
            ce = ""
        col += 1
        tpa_audit_tab.write(tpa_row, col, prescription["prescriber_name"])
        col += 1
        
        tpa_audit_tab.write(tpa_row, col, prescription["manufacturer"])
        col += 1

        if ever_results:
            ic(ever_results)
            tpa_audit_tab.write(tpa_row, col, ce)
        else:
            # change this to lookup from NPI instead of Qualified Manufacturers
            tpa_audit_tab.write_formula(tpa_row, col, f"=VLOOKUP(N{tpa_row+1},'Qualified Prescribers'!A:E,5,FALSE)")
        col += 1
        tpa_audit_tab.write(tpa_row, col, prescription["prescriber_npi"])
        col += 1
        tpa_audit_tab.write(tpa_row, col, first_audit_period)
        col += 1

        tpa_row += 1
tpa_audit_tab.autofilter(2, 0, tpa_row, len(tpa_headers)-1)


# create TPA Rx Review - ROI tab
roitab = workbook.add_worksheet("TPA Rx Review - ROI Tab")
roitab.set_tab_color("81A3A7")
roitab.write_url('A1',  "internal:'Summary'!A1", string="Return to Summary")
roirow = 2
roi_headers = ["Rx Number",
               "Fill Number",
               "Fill Date",
               "NDC",
               "Description",
               "Total Paid Amount",
               "Est. True Margin",
               "Dispense Fee",
               "Pharmacy Impact",
               "Status",
               "Prescriber Number",
               "Qualifying Prescriber Name",
               "Qualifying Manufacturer",
               "Potential Covered Entity"]
roitab.set_column(0, len(roi_headers)-1, 20)
for idx, header in enumerate(roi_headers):
    roitab.write(roirow, idx, header, title_format)
roirow += 1
# get the list of roi candidates
roi_query = """SELECT * FROM counterfill_audit_rxs WHERE pharmacy = %s ORDER BY rx_fill_num ASC;"""
roi_inputs = (pharmacy_name,)
cursor.execute(roi_query, roi_inputs)
roi_candidates = cursor.fetchall()
for roi_candidate in roi_candidates:
    # get pharm_data prescription info
    pharm_query = """SELECT * FROM counterfill_claims WHERE pharmacy_name = %s AND rx_fill_concat = %s;"""
    pharm_inputs = (pharmacy_name, roi_candidate["rx_fill_num"])
    cursor.execute(pharm_query, pharm_inputs)
    pharm_data = cursor.fetchone()
    if pharm_data is None:
        print(f"Skipping {roi_candidate['rx_fill_num']} - no pharm_data found")
        continue
    # get 340B claim info
    roi340b_query = """SELECT * FROM 340b_claims WHERE rx_fill_concat = %s AND prescriber_npi = %s ORDER BY fill_date DESC LIMIT 1;"""
    roi340b_inputs = (roi_candidate["rx_fill_num"], pharm_data["prescriber_npi"])
    cursor.execute(roi340b_query, roi340b_inputs)
    roi340b_data = cursor.fetchone()
    if roi340b_data is None:
        status = "Pending Investigation"
        ever340b = "NO"
        pharmacy_impact = 0
        disp_fee = 0
    else:
        status = roi340b_data["status"]
        ever340b = roi340b_data["fill_date"]
        disp_fee = roi340b_data["disp_fee"]
        pharmacy_impact = float(roi340b_data["disp_fee"]) - float(roi340b_data["retail_margin"])
        if pharmacy_impact < 0:
            continue
    if pharmacy_impact == 0 and pharm_data["fill_date"] < datetime.datetime.strptime(report_start_date, "%Y-%m-%d").date():
        continue
    est_paid_to_ce = float(pharm_data["total_payment"]) - float(pharm_data["est_disp_fee"])
    roitab.write(roirow, 0, str(roi_candidate["rx_fill_num"].split("-")[0]))
    roitab.write(roirow, 1, str(roi_candidate["rx_fill_num"].split("-")[1]))
    roitab.write(roirow, 2, pharm_data["fill_date"], date_format)
    roitab.write(roirow, 3, str(roi_candidate["ndc11"]))
    roitab.write(roirow, 4, pharm_data["drug_name"])
    roitab.write(roirow, 5, pharm_data["total_payment"], money)
    roitab.write(roirow, 6, pharm_data["retail_margin"], money)
    roitab.write(roirow, 7, disp_fee, money)
    roitab.write(roirow, 8, pharmacy_impact, money)
    roitab.write(roirow, 9, status)
    roitab.write(roirow, 10, pharm_data["prescriber_npi"])
    roitab.write(roirow, 11, pharm_data["prescriber_name"])
    roitab.write(roirow, 12, pharm_data["manufacturer"])
    roitab.write_formula(roirow, 13, f"=VLOOKUP(K{roirow+1},'Qualified Prescribers'!A:E,5,FALSE)")
    roirow += 1
roitab.autofilter(2, 0, roirow, len(roi_headers)-1)

# create Medicaid plan tab
medicaidtab = workbook.add_worksheet("Medicaid Plan Info")
medicaidtab.set_tab_color("81A3A7")
medicaidtab.set_column(0, 12, 20)
medicaidtab.write_url('A1',  "internal:'Summary'!A1", string="Return to Summary")
medicaidheaders = ["Plan Name", "BIN", "PCN", "Group","Concat","State"]
msql = """SELECT * FROM counterfill_medicaid WHERE state = %s;"""
mrecord = (pharmacy_state,)
cursor.execute(msql, mrecord)
medicaids = cursor.fetchall()
medicaidrow = 1
for idx, header in enumerate(medicaidheaders):
    medicaidtab.write(0, idx, header, title_format)
for medicaid in medicaids:
    medicaidcol = 0
    medicaidtab.write(medicaidrow, medicaidcol, medicaid["plan_name"])
    medicaidcol += 1
    medicaidtab.write(medicaidrow, medicaidcol, medicaid["bin"])
    medicaidcol += 1
    medicaidtab.write(medicaidrow, medicaidcol, medicaid["pcn"])
    medicaidcol += 1
    medicaidtab.write(medicaidrow, medicaidcol, medicaid["rx_group"])
    medicaidcol += 1
    medicaidtab.write(medicaidrow, medicaidcol, medicaid["concat"])
    medicaidcol += 1
    medicaidtab.write(medicaidrow, medicaidcol, medicaid["state"])
    medicaidrow += 1

# create tpa qualified claims tab
tpa_qc_tab = workbook.add_worksheet("Invoices")
tpa_qc_tab.set_tab_color("81A3A7")
tpa_qc_tab.set_column(0, 27, 15)
tpa_qc_tab.freeze_panes(1, 0)
tpa_row = 0
tpa_headers = ['RxNumber',
'FillNumber',
'Rx+Fill',
'Status',
'Dispensed',
'Doctor NPI',
'NDC',
'NDC Description',
'Days Supply',
'Amount',
'Copay',
'DispenseFee',
'Revenue',
'Quantity',
'BUPP',
'MHI Pkgs',
'BIN',
'PCN',
'GRP',
'Qualified Date',
'Prescriber Last Name',
'Brand',
'Covered Entity',
'TPA',
'Pharmacy',
'Input File']
for idx, header in enumerate(tpa_headers):
    tpa_qc_tab.write(tpa_row, idx, header, title_format)
tpa_row += 1

for report in report_identifiers:
    print(f"adding claims for {report['report_identifier']}")
    tpa_query = """SELECT * FROM 340b_claims
        WHERE report_identifier = %s
        AND fill_date BETWEEN %s AND %s;"""
    tpa_inputs = (report["report_identifier"], report_start_date, report_end_date)
    cursor.execute(tpa_query, tpa_inputs)
    tpa_claims = cursor.fetchall()
    for claim in tpa_claims:
        # print(f"adding claim {claim['rx_number']}")
        col = 0
        tpa_qc_tab.write(tpa_row, col, claim["rx_number"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["fill_number"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["rx_fill_concat"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["status"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["fill_date"], date_format)
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["prescriber_npi"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["ndc"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["drug_name"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["days_supply"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["total_payment"], money)
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["copay"], money)
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["disp_fee"], money)
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["revenue"], money)
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["qty_disp"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["bupp"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["pkgs_disp"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["bin"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["pcn"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["rx_group"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["bill_date"], date_format)
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["prescriber_name"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["indicator"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["covered_entity"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["tpa"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["pharmacy_name"])
        col += 1
        tpa_qc_tab.write(tpa_row, col, claim["input_file"])

        tpa_row += 1
    
tpa_qc_tab.autofilter(0, 0, tpa_row, len(tpa_headers)-1)
tpa_qc_tab.set_column(25, 25, None, None, {'hidden': 1})

# create InvenSTORY tab
print("creating InvenSTORY tab")
inventab = workbook.add_worksheet("InvenSTORY")
inventab.set_tab_color("81A3A7")
inventab.set_column(0, 12, 20)
inventab.freeze_panes(1, 0)
inv_row = 0
inv_headers = [
    'Covered Entity',
    'NDC',
    'Description',
    'Indicator',
    'Manufacturer',
    'Package Price',
    'Dispensed Packages',
    'Dispensed Value',
    'Replenished Packages',
    'Replenished Value',
    'Variance',
    'Variance Value',
    'Accumulator Packages'
]
for idx, header in enumerate(inv_headers):
    inventab.write(inv_row, idx, header, title_format)
inv_row += 1
for report in report_identifiers:
    report_sql = """SELECT * FROM report_queue WHERE report_identifier = %s;"""
    report_inputs = (report["report_identifier"],)
    cursor.execute(report_sql, report_inputs)
    report_info = cursor.fetchone()
    payment_model = report_info["payment_model"]
    data_source = report_info["data_source"]
    if payment_model == "POR" or data_source == "Invoices":
        invs_query = """SELECT DISTINCT ndc FROM 340b_claims WHERE bill_date BETWEEN %s AND %s AND report_identifier = %s
            UNION
            SELECT DISTINCT ndc11 as ndc FROM replenishments WHERE replenishment_date BETWEEN %s AND %s AND report_identifier = %s;"""
    else:
        invs_query = """SELECT DISTINCT ndc FROM 340b_claims WHERE fill_date BETWEEN %s AND %s AND report_identifier = %s
            UNION
            SELECT DISTINCT ndc11 as ndc FROM replenishments WHERE replenishment_date BETWEEN %s AND %s AND report_identifier = %s;"""
    invs_input = (
        report_start_date,
        report_end_date,
        report_info["report_identifier"],
        report_start_date,
        report_end_date,
        report_info["report_identifier"],
    )
    cursor.execute(invs_query, invs_input)
    invs_ndcs = cursor.fetchall()
    for inv_ndc in invs_ndcs:
        # get drug info from drug_catalog
        drug_query = """SELECT * FROM drug_catalog WHERE ndc11 = %s LIMIT 1;"""
        cursor.execute(drug_query, (inv_ndc["ndc"],))
        drug_info = cursor.fetchone()
        if drug_info is None:
            description = ""
            indicator = ""
            package_price = None
        else:
            description = drug_info["description"]
            indicator = drug_info["indicator"]
            package_price = drug_info["price"]

        # get manufacturer info from manuf_exclusions
        manuf_query = """SELECT * FROM manuf_exclusions WHERE ndc11 = %s LIMIT 1;"""
        cursor.execute(manuf_query, (inv_ndc["ndc"],))
        manuf_info = cursor.fetchone()
        if manuf_info is None:
            manufacturer = "Not restricted manufacturer"
        else:
            manufacturer = manuf_info["manufacturer"]

        # get dispensed packages
        dispensed_query = """SELECT IFNULL(SUM(pkgs_disp), 0) as pkgs_dispensed FROM 340b_claims
            WHERE ndc = %s
            AND fill_date BETWEEN %s AND %s
            AND report_identifier = %s;"""
        dispensed_inputs = (inv_ndc["ndc"], report_start_date, report_end_date, report_info["report_identifier"])
        cursor.execute(dispensed_query, dispensed_inputs)
        dispensed_result = cursor.fetchone()

        # get replenished packages
        replenished_query = """SELECT IFNULL(SUM(num_pkgs), 0) as pkgs_dispensed FROM replenishments
            WHERE ndc11 = %s
            AND replenishment_date BETWEEN %s AND %s
            AND report_identifier = %s;"""
        replenished_inputs = (inv_ndc["ndc"], report_start_date, report_end_date, report_info["report_identifier"])
        cursor.execute(replenished_query, replenished_inputs)
        replenished_result = cursor.fetchone()

        # get accumulator packages
        accumulator_query = """SELECT IFNULL(SUM(num_pkgs), 0) as pkgs_dispensed FROM accumulator
            WHERE ndc11 = %s
            AND accumulator_date BETWEEN %s AND %s
            AND report_identifier = %s;"""
        accumulator_inputs = (inv_ndc["ndc"], report_start_date, report_end_date, report_info["report_identifier"])
        cursor.execute(accumulator_query, accumulator_inputs)
        accumulator_result = cursor.fetchone()

        invs_col = 0
        inventab.write(inv_row, invs_col, report_info["covered_entity"])
        invs_col += 1
        inventab.write(inv_row, invs_col, inv_ndc["ndc"])
        invs_col += 1
        inventab.write(inv_row, invs_col, description)
        invs_col += 1
        inventab.write(inv_row, invs_col, indicator)
        invs_col += 1
        inventab.write(inv_row, invs_col, manufacturer)
        invs_col += 1
        inventab.write(inv_row, invs_col, package_price, money)
        invs_col += 1
        inventab.write(inv_row, invs_col, dispensed_result["pkgs_dispensed"])
        invs_col += 1
        disp_value = float(dispensed_result["pkgs_dispensed"]) * float(package_price) if package_price else 0
        inventab.write(inv_row, invs_col, disp_value, money)
        invs_col += 1
        inventab.write(inv_row, invs_col, replenished_result["pkgs_dispensed"])
        invs_col += 1
        replenished_value = float(replenished_result["pkgs_dispensed"]) * float(package_price) if package_price else 0
        inventab.write(inv_row, invs_col, replenished_value, money)
        invs_col += 1
        variance_pkgs = float(replenished_result["pkgs_dispensed"]) - float(dispensed_result["pkgs_dispensed"])
        inventab.write(inv_row, invs_col, variance_pkgs)
        invs_col += 1
        variance_value = replenished_value - disp_value
        inventab.write(inv_row, invs_col, variance_value, money)
        invs_col += 1
        inventab.write(inv_row, invs_col, accumulator_result["pkgs_dispensed"])
        invs_col += 1

        inv_row += 1
inventab.autofilter(0, 0, inv_row, len(inv_headers)-1)



# create Accumulator tab
print("creating Accumulator tab")
accumtab = workbook.add_worksheet("Accumulator")
accumtab.set_tab_color("81A3A7")
accumtab.set_column(0, 11, 20)
accumtab.freeze_panes(1, 0)
accum_row = 0
accum_headers = [
    'Covered Entity',
    'NDC11',
    'NDC Description',
    'Indicator',
    '340B Pkgs',
    'Prev Report Pkgs',
    'Est Acq Cost Per Pkg',
    'Ext Cost',
    'Manufacturer',
    'Accumulator Date',
    'Last Replenishment Date',
    'Input File']
for idx, header in enumerate(accum_headers):
    accumtab.write(accum_row, idx, header, title_format)
accum_row += 1

for report in report_identifiers:
    accum_date = None
    # get the max accumulator date for this report
    accum_date_query = """SELECT MAX(accumulator_date) as max_date FROM accumulator
        WHERE report_identifier = %s;"""
    accum_date_inputs = (report["report_identifier"],)
    cursor.execute(accum_date_query, accum_date_inputs)
    accum_date_result = cursor.fetchone()
    
    accum_query = """SELECT * FROM accumulator
        WHERE report_identifier = %s
        AND accumulator_date = %s;"""
    accum_inputs = (report["report_identifier"], accum_date_result["max_date"])
    cursor.execute(accum_query, accum_inputs)
    accumulators = cursor.fetchall()
    for accumulator in accumulators:
        col = 0
        accumtab.write(accum_row, col, accumulator["covered_entity"])
        col += 1
        accumtab.write(accum_row, col, accumulator["ndc11"])
        col += 1
        accumtab.write(accum_row, col, accumulator["drug_name"])
        col += 1
        accumtab.write(accum_row, col, '@todo')
        col += 1
        accumtab.write(accum_row, col, accumulator["num_pkgs"])
        col += 1
        accumtab.write(accum_row, col, '')
        col += 1
        accumtab.write(accum_row, col, accumulator["wac_price"], money)
        col += 1
        accumtab.write(accum_row, col, accumulator["extended_cost"], money)
        col += 1
        accumtab.write(accum_row, col, accumulator["manufacturer"])
        col += 1
        accumtab.write(accum_row, col, accumulator["accumulator_date"], date_format)
        col += 1
        accumtab.write(accum_row, col, '')
        col += 1
        accumtab.write(accum_row, col, accumulator["input_file"])

        accum_row += 1
accumtab.autofilter(0, 0, accum_row, len(accum_headers)-1)
accumtab.set_column(11, 11, None, None, {'hidden': 1})

# create Replenishments tab (using the old purchases tab, hence the odd variable names)
print("creating Replenishments tab")
purchtab = workbook.add_worksheet("Replenishments")
purchtab.set_tab_color("81A3A7")
purchtab.set_column(0, 8, 15)
purchtab.freeze_panes(1, 0)
purch_row = 0
purch_headers = [
    'Covered Entity',
    'NDC11',
    'NDC Description',
    'Indicator',
    '340B Pkgs',
    'Est Acq Cost Per Pkg',
    'Ext Cost',
    'Manufacturer',
    'Replenishment Date',
    'Input File']
for idx, header in enumerate(purch_headers):
    purchtab.write(purch_row, idx, header, title_format)
purch_row += 1

for report in report_identifiers:
    purchase_query = """SELECT * FROM replenishments
        WHERE report_identifier = %s
        AND replenishment_date BETWEEN %s AND %s;"""
    purchase_inputs = (report["report_identifier"], report_start_date, report_end_date)
    cursor.execute(purchase_query, purchase_inputs)
    purchases = cursor.fetchall()
    for purchase in purchases:
        col = 0
        purchtab.write(purch_row, col, purchase["covered_entity"])
        col += 1
        purchtab.write(purch_row, col, purchase["ndc11"])
        col += 1
        purchtab.write(purch_row, col, purchase["drug_name"])
        col += 1
        purchtab.write(purch_row, col, purchase["indicator"])
        col += 1
        purchtab.write(purch_row, col, purchase["num_pkgs"])
        col += 1
        purchtab.write(purch_row, col, purchase["wac_price"], money)
        col += 1
        purchtab.write(purch_row, col, purchase["extended_cost"], money)
        col += 1
        purchtab.write(purch_row, col, purchase["manufacturer"])
        col += 1
        purchtab.write(purch_row, col, purchase["replenishment_date"], date_format)
        col += 1
        purchtab.write(purch_row, col, purchase["input_file"])


        purch_row += 1
    
purchtab.autofilter(0, 0, purch_row, len(purch_headers)-1)
purchtab.set_column(9, 9, None, None, {'hidden': 1})

# create NDC tab
ndctab = workbook.add_worksheet("NDC")
ndctab.write_url('A1',  "internal:'Summary'!A1", string="Return to Summary")


# add to qc tab
qcrow += 1
qctab.write(qcrow, 2, "Count")
qctab.write(qcrow, 3, "Brand Count")
qctab.write(qcrow, 4, "Generic Count")
qcrow += 1
sql = """SELECT input_file, count(*) FROM counterfill_claims 
    WHERE pharmacy_name = %s
    AND fill_date BETWEEN %s AND %s
    GROUP BY input_file;"""
inputs = (pharmacy_name, report_start_date, report_end_date)
cursor.execute(sql, inputs)
pharm_files = cursor.fetchall()
for pharm_file in pharm_files:
    brand_count = 0
    generic_count = 0
    brand_sql = """SELECT COUNT(*) FROM counterfill_claims where input_file = %s and indicator = 'B';"""
    generic_sql = """SELECT COUNT(*) FROM counterfill_claims where input_file = %s and indicator = 'G';"""
    cursor.execute(brand_sql, (pharm_file["input_file"],))
    brand_count = cursor.fetchone()
    cursor.execute(generic_sql, (pharm_file["input_file"],))
    generic_count = cursor.fetchone()
    qctab.write(qcrow, 0, "pharm_input_file")
    qctab.write(qcrow, 1, pharm_file["input_file"])
    qctab.write(qcrow, 2, pharm_file["count(*)"])
    qctab.write(qcrow, 3, brand_count["COUNT(*)"])
    qctab.write(qcrow, 4, generic_count["COUNT(*)"])
    qcrow += 1
# ic(report_identifiers)
qcrow += 1
for report_identifier in report_identifiers:
    qcrow += 1
    qctab.write(qcrow, 0, report_identifier["report_identifier"])
    qcrow += 1
    # ic(report_identifier)
    tpa_query = """SELECT tpa, input_file, report_identifier, date(timestamp), count(*) FROM 340b_claims
            WHERE report_identifier = %s
            AND fill_date BETWEEN %s AND %s
            GROUP BY tpa, input_file, report_identifier, date(timestamp)
            ORDER BY date(timestamp) ASC;"""
    tpa_inputs = (report_identifier['report_identifier'], report_start_date, report_end_date)
    cursor.execute(tpa_query, tpa_inputs)
    tpa_claims = cursor.fetchall()
    # ic(tpa_claims)
    for tpa_claim in tpa_claims:
        qctab.write(qcrow, 0, tpa_claim["tpa"]+" TPA Invoice File")
        qctab.write(qcrow, 1, tpa_claim["input_file"])
        qctab.write(qcrow, 2, tpa_claim["count(*)"])
        qctab.write(qcrow, 3, tpa_claim["report_identifier"])
        qctab.write(qcrow, 4, "tpa upload date "+ tpa_claim["date(timestamp)"].strftime("%Y-%m-%d"))
        qcrow += 1
    replenishment_query = """SELECT input_file, report_identifier, date(timestamp), count(*) FROM replenishments
        WHERE report_identifier = %s
        AND replenishment_date BETWEEN %s AND %s
        GROUP BY input_file, report_identifier, date(timestamp)
        ORDER BY date(timestamp) ASC;"""
    replenishment_inputs = (report_identifier['report_identifier'], report_start_date, report_end_date)
    cursor.execute(replenishment_query, replenishment_inputs)
    replenishments = cursor.fetchall()
    for replenishment in replenishments:
        qctab.write(qcrow, 0, "replenishment_input_file")
        qctab.write(qcrow, 1, replenishment["input_file"])
        qctab.write(qcrow, 2, replenishment["count(*)"])
        qctab.write(qcrow, 3, replenishment["report_identifier"])
        qctab.write(qcrow, 4, "rep upload "+replenishment["date(timestamp)"].strftime("%Y-%m-%d"))
        qcrow += 1
qcrow += 1
for report_identifier in report_identifiers:
    qctab.write(qcrow, 0, "report_identifier")
    qctab.write(qcrow, 1, report_identifier["report_identifier"])
    qcrow += 1
qcrow += 1
qcrow += 1
qctab.write(qcrow, 0, f"Claims with no NDC (skipped): {no_ndc}")
qcrow += 1
qctab.write(qcrow, 0, f"Claims with no total payment (skipped): {no_total_payment}")
qcrow += 1
qctab.write(qcrow, 0, f"Size of audit history no-fly list: {len(tpa_audit_history_list)}")
qcrow += 1
qctab.write(qcrow, 0, f"Audit History Entries we have seen before (skipped): {audit_skips}")


workbook.close()

# record the report
rec_sql = """INSERT INTO reports_created (quarter, pharmacy, covered_entity, 
    filename, num_claims, tot_payment, tpa, report_identifier, report_queue_id) 
    VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
rec_sql_data = (
    report_label,
    pharmacy_name,
    None,
    filename,
    None,
    None,
    None,
    None,
    report_id,
)

try:
    cursor.execute(rec_sql, rec_sql_data)
    conn.commit()
    print("report created table updated successfully")
except mysql.connector.Error as err:
    # Rolling back in case of error
    conn.rollback()
    print("couldn't update report table {}".format(err))
    sys.exit("something wrong with reports_created table")

# update report queue
rpt_queue_sql = """UPDATE report_queue SET last_report_period = %s, last_report_status = %s, last_report_path = %s WHERE id = %s"""
rpt_data = (report_period, "cf-created", path_filename, report_id)
cursor.execute(rpt_queue_sql, rpt_data)
conn.commit()


ic(is340b_count)
# Closing the connection
cursor.close()
conn.close()
end_time = datetime.datetime.now()
duration = end_time - starttime
ic(duration)
print(f"report id = {report_id}")