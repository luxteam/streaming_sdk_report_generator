import os
import win32com.client
from common import Reports, Jobs
from jenkins_export import get_latest_report, get_report_link
from jira_export import get_issues
from datetime import timedelta, datetime
import lxml.html as lh
from copy import deepcopy
from enum import Enum

reports_titles = {
    Reports.PUBG: "PUBG Report",
    Reports.Dota2_DX11: "Dota 2 DX11 Report",
    Reports.Dota2_Vulkan: "Dota 2 Vulkan Report",
    Reports.LoL: "League of Legends Report",
    Reports.Heaven_Benchmark_DX9: "Heaven Benchmark DX9 Report",
    Reports.Valley_Benchmark_DX9: "Valley Benchmark DX9 Report",
    Reports.Heaven_Benchmark_DX11: "Heaven Benchmark DX11 Report",
    Reports.Valley_Benchmark_DX11: "Heaven Benchmark DX11 Report",
    Reports.Heaven_Benchmark_OpenGL: "Heaven Benchmark OpenGL Report",
    Reports.Valley_Benchmark_OpenGL: "Heaven Benchmark OpenGL Report",
}

class Machines(Enum):
    AMD_6750 = 0
    AMD_7900 = 1


machines_names = {
    Machines.AMD_6750: "AMD Radeon RX 6750 XT Windows 10(64bit)",
    Machines.AMD_7900: "AMD Radeon RX 7900 XT Windows 10(64bit)",
}


LETTER1_HTML_TABLE = {
    Jobs.Full_Samples:
    {
        Machines.AMD_6750: "FULL_SAMPLES_6750_TABLE",
        Machines.AMD_7900: "FULL_SAMPLES_7900_TABLE"
    },
    Jobs.Win_Full: {
        Machines.AMD_6750: "REMOTE_SAMPLES_6750_TABLE",
        Machines.AMD_7900: "REMOTE_SAMPLES_7900_TABLE"
    },
}

LETTER2_HTML_TABLE = "ISSUES_TABLE"

RECIPIENTS_TO = os.getenv("STREAMING_SDK_EMAIL_RECIPIENTS_TO", "")
RECIPIENTS_CC = os.getenv("STREAMING_SDK_EMAIL_RECIPIENTS_CC", "")

def load_xml(file_path: str):
    tree = None
    with open(file_path, "r") as file:
        tree = lh.parse(file)

    return tree


def write_xml(tree, file_path: str):
    tree.write(file_path, xml_declaration=True, encoding="ascii")


def append_row_to_summary_table(tbody: lh.Element, report_type: Reports, report: dict, row_template: lh.Element, report_url: str):
    # append new row
    row = deepcopy(row_template)
    tbody.append(row)

    columns = row.findall("./td")

    columns[0].find("./p/span/a").set("href", report_url)
    columns[0].find("./p/span/a/span").text = reports_titles[report_type]

    columns[1].find("./p/span").text = str(report["total"])
    columns[2].find("./p/span").text = str(report["passed"])
    columns[3].find("./p/span").text = str(report["failed"])
    columns[4].find("./p/span").text = str(report["error"])
    h, m, s = str(timedelta(seconds=int(
        report["execution_time"]))).split(":")
    columns[5].find("./p/span").text = f"{h}h {m}m {s}s"



def generate_first_letter(recipients_to: str = "", recipients_cc: str = ""):
    html = load_xml("letters_templates/Letter1.html")

    for job in [Jobs.Full_Samples, Jobs.Win_Full]:
        tables = {
            Machines.AMD_6750: html.find("//table[@id='{id}']".format(id=LETTER1_HTML_TABLE[job][Machines.AMD_6750])),
            Machines.AMD_7900: html.find("//table[@id='{id}']".format(id=LETTER1_HTML_TABLE[job][Machines.AMD_7900])),
        }
        tbodies = {machine: tables[machine].find("./tbody") for machine in tables}
        rows = {machine: tbodies[machine].findall("./tr")[1] for machine in tbodies}

        # copy row and remove it from table
        row_templates = {machine: deepcopy(rows[machine]) for machine in rows}
        for machine in tbodies:
            tbodies[machine].remove(rows[machine])

        for report in Reports:
            if report is Reports.summary:
                continue

            since_date = (datetime.today() - timedelta(weeks=1) + timedelta(days=1)
                          ).replace(hour=0, minute=0, second=0, microsecond=0)

            latest_report = get_latest_report(
                job, report, newer_than=since_date
            )

            if latest_report is None:
                continue

            report_url = get_report_link(job, latest_report["version"], report, json=False)
            json_report = latest_report["report"]

            for machine in Machines:
                machine_name = machines_names[machine]
                append_row_to_summary_table(tbody=tbodies[machine], report_type=report, report=json_report[machine_name]["summary"], row_template=row_templates[machine], report_url=report_url)

    dir = os.getcwd()
    html_file = os.path.join(dir, "Letter_1.html")
    write_xml(html, html_file)

    oft_file = os.path.join(dir, "Letter_1.oft")
    html2oft(html_file, oft_file, message_subject="Streaming SDK Report", recipients_to=recipients_to, recipients_cc=recipients_cc)


def generate_second_letter(report_date: datetime, recipients_to: str = "", recipients_cc: str = ""):
    html = load_xml("letters_templates/Letter2.html")

    table = html.find("//table[@id='{id}']".format(id=LETTER2_HTML_TABLE))
    tbody = table.find("./tbody")
    row = tbody.findall("./tr")[0]

    # copy row and remove it from table
    row_template = deepcopy(row)
    tbody.remove(row)

    issues = get_issues()
    for issue in issues:
        # append new row
        row = deepcopy(row_template)
        tbody.append(row)

        columns = row.findall("./td")

        columns[0].find("./p/span/a").set("href", issue.url)
        columns[0].find("./p/span/a/span").text = issue.key
        columns[1].find("./p/span").text = issue.summary

        columns[2].find("./p/span").text = issue.created_at
        if datetime.strptime(issue.created_at, "%d/%b/%y").year == 2021:
            span = columns[2].find("./p/span")
            span.attrib['style'] = span.attrib['style'].replace("color:black", "color:#C00000")

        columns[3].find("./p/span").text = issue.severity
        if issue.severity.lower() in ["blocker", "critical"]:
            span = columns[3].find("./p/span")
            span.attrib['style'] = span.attrib['style'].replace("color:black", "color:#C00000")

    dir = os.getcwd()
    html_file = os.path.join(dir, "Letter_2.html")
    write_xml(html, html_file)

    oft_file = os.path.join(dir, "Letter_2.oft")
    html2oft(html_file, oft_file, message_subject="Weekly QA Report " + report_date.strftime("%d-%b-%Y"), recipients_to=recipients_to, recipients_cc=recipients_cc)


def html2oft(html_file_path: str, otf_file_path: str, recipients_to: str = "", recipients_cc: str = "", message_subject: str = ""):
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")

    msg = obj.CreateItem(olMailItem)
    msg.Subject = message_subject
    msg.To = recipients_to
    msg.Cc = recipients_cc
    # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    msg.BodyFormat = 2
    msg.HTMLBody = open(html_file_path).read()
    # newMail.display()
    save_format = 2  # olTemplate	2	Microsoft Outlook template (.oft)
    msg.SaveAs(otf_file_path, save_format)


if __name__ == "__main__":
    report_date = datetime.today()
    generate_first_letter(recipients_to=RECIPIENTS_TO, recipients_cc=RECIPIENTS_CC)
    generate_second_letter(recipients_to=RECIPIENTS_TO, recipients_cc=RECIPIENTS_CC, report_date=report_date)
