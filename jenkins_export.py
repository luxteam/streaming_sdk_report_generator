import os
import requests
import json
from requests.auth import HTTPBasicAuth
from common import Jobs, Reports
from typing import Dict, Optional
from datetime import datetime

JENKINS_HOST = os.getenv("JENKINS_HOST", "rpr.cis.luxoft.com")
CIS_HOST = os.getenv("CIS_HOST", "cis.nas.luxoft.com")
JENKINS_USERNAME = os.environ["JENKINS_USERNAME"]
JENKINS_TOKEN = os.environ["JENKINS_TOKEN"]


jobs_names = {
    Jobs.Full_Samples: "FullSamples-Weekly",
    Jobs.Win_Full: "StreamingSDK-Windows-WeeklyFull",
    Jobs.Win_APU: "StreamingSDK-APU-WeeklyFull",
    Jobs.Android_Full: "StreamingSDK-Android-WeeklyFull",
    Jobs.Android_Xiaomi_TV: "StreamingSDK-XiaomiTVStick-WeeklyFull",
    Jobs.Android_Chromecast_TV: "StreamingSDK-Chromecast-WeeklyFull",
    Jobs.Ubuntu_Full: "StreamingSDK-Ubuntu-WeeklyFull",
    Jobs.AMD_Full: "AMDLink-Weekly",
    Jobs.Win_Latency: "StreamingSDK-LatencyTests",
    Jobs.Win_Long_Term: "StreamingSDK-LongTermTests",
}


reports_names = {
    Reports.summary: "Test_Report",
    Reports.PUBG: "Test_Report_PUBG",
    Reports.Dota2_DX11: "Test_Report_Dota2DX11",
    Reports.Dota2_Vulkan: "Test_Report_Dota2Vulkan",
    Reports.LoL: "Test_Report_LoL",
    Reports.Heaven_Benchmark_DX9: "Test_Report_HeavenDX9",
    Reports.Valley_Benchmark_DX9: "Test_Report_ValleyDX9",
    Reports.Heaven_Benchmark_DX11: "Test_Report_HeavenDX11",
    Reports.Valley_Benchmark_DX11: "Test_Report_ValleyDX11",
    Reports.Heaven_Benchmark_OpenGL: "Test_Report_HeavenOpenGL",
    Reports.Valley_Benchmark_OpenGL: "Test_Report_ValleyOpenGL",
}

jobs_representative_reports = {
    Jobs.Full_Samples: Reports.PUBG,
    Jobs.Win_Full: Reports.LoL,
    Jobs.Android_Full: Reports.LoL,
    Jobs.Ubuntu_Full: Reports.Valley_Benchmark_OpenGL,
}


def get_build_link(job: Jobs, latest_build_number: int):
    name = jobs_names.get(job)

    return f"http://{JENKINS_HOST}/job/{name}/{latest_build_number}/"


def get_latest_build_number(job: Jobs) -> Optional[int]:
    name = jobs_names.get(job)
    url = f"http://{JENKINS_HOST}/job/{name}/api/json?tree=lastBuild[id]"

    resp = requests.get(url, auth=HTTPBasicAuth(JENKINS_USERNAME, JENKINS_TOKEN))

    last_build = resp.json()["lastBuild"]

    if not last_build:
        return None

    id = last_build["id"]

    return int(id)


def get_report_link(
    job: Jobs, build_number: int, report: Reports, json: bool = True
) -> str:
    report_url = (
        "https://{host}/{job_name}/{build_number}/{report_name}/{report_type}".format(
            host=CIS_HOST,
            job_name=jobs_names[job],
            build_number=build_number,
            report_name=reports_names[report],
            report_type="summary_report.json" if json else "summary_report.html",
        )
    )

    return report_url


def get_latest_report(
    job: Jobs, report: Reports, newer_than: Optional[datetime] = None
) -> Optional[dict]:
    build_number = get_latest_build_number(job)
    if build_number is None:
        return None

    report_url = get_report_link(job, build_number, report)
    resp = requests.get(report_url, auth=HTTPBasicAuth(JENKINS_USERNAME, JENKINS_TOKEN))

    while (resp.status_code != 200) and build_number >= 0:
        report_url = get_report_link(job, build_number, report)
        resp = requests.get(
            report_url, auth=HTTPBasicAuth(JENKINS_USERNAME, JENKINS_TOKEN)
        )
        build_number -= 1

    if resp is None or resp.status_code != 200:
        return None

    json_report = resp.json()

    if not json_report:
        print(f"WARNING: JSON report {report_url} is not available!")
        return None
    

    for machine_report in json_report.values():
        for report in list(machine_report["results"].values()):
            if not report[""].get("machine_info"):
                print(f"ERROR: Report {report_url} is broken!")
                return None


    if newer_than is not None:
        reporting_date = max(
            [
                datetime.strptime(reporting_date, "%m/%d/%Y %H:%M:%S")
                for reporting_date in [
                    report[""]["machine_info"]["reporting_date"]
                    for machine_report in json_report.values()
                    for report in list(machine_report["results"].values())
                ]
            ]
        )

        if reporting_date < newer_than:
            return None

    return {"version": build_number, "report": json_report}


def get_skipped_or_observed_per_group(
    job: Jobs, report: Reports = None
) -> Optional[Dict[str, int]]:
    if report is None:
        if job in jobs_representative_reports:
            report = jobs_representative_reports[job]
        else:
            return None

    latest_report = get_latest_report(job, report)

    if latest_report is None:
        return {}

    json_report = latest_report["report"]

    machine_name = [machine for machine in json_report.keys() if "7900" in machine]
    if machine_name:
        machine_name = machine_name[0]
    else:
        machine_name = list(json_report.keys())[0]

    groups_list = json_report[machine_name][
        "results"
    ]  # report for the AMD 7900 machine prioritized
    skipped_or_observed_per_group = {
        key: groups_list[key][""]["observed"] + groups_list[key][""]["skipped"]
        for key in groups_list
    }

    return {
        key: skipped_or_observed_per_group[key]
        for key in skipped_or_observed_per_group
        if skipped_or_observed_per_group[key] > 0
    }


if __name__ == "__main__":
    print("Runs:")
    for job in Jobs:
        id = get_latest_build_number(job)
        if id is None:
            print("latest build not found")
            continue

        print("latest build number: " + str(id))
        print(get_build_link(job, id))

    print("Representative reports:")
    for job in jobs_representative_reports:
        print(
            json.dumps(
                get_skipped_or_observed_per_group(
                    job, jobs_representative_reports[job]
                ),
                indent=4,
            )
        )

    # print("Reports")
    # for job in Jobs:
    #     for report in Reports:
    #         print(
    #             json.dumps(
    #                 get_latest_report(
    #                     job, report
    #                 ),
    #                 indent=4,
    #             )
    #         )