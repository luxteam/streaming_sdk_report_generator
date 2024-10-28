import os
import shutil
from datetime import datetime, timedelta
from typing import List, Dict
from common import Jobs, WORKING_DIR_PATH, REPORT_FILE_PATH, TEMPLATE_PATH, Issue
from jira_export import get_issues
from jenkins_export import (
    get_latest_build_number,
    get_build_link,
    get_skipped_or_observed_per_group,
)
from confluence_export import get_project_status
import word
import ids
from lxml import etree

jobs_link_title = {
    Jobs.Full_Samples: "FullSamples-Weekly #{num}",
    Jobs.Win_Full: "StreamingSDK-Windows-WeeklyFull #{num}",
    Jobs.Win_APU: "StreamingSDK-APU-WeeklyFull #{num}",
    Jobs.Android_Full: "StreamingSDK-Android-Weekly #{num}",
    Jobs.Android_Xiaomi_TV: "StreamingSDK-XiaomiTVStick-Weekly #{num}",
    Jobs.Android_Chromecast_TV: "StreamingSDK-Chromecast-Weekly #{num}",
    Jobs.Ubuntu_Full: "StreamingSDK-Ubuntu-WeeklyFull #{num}",
    Jobs.AMD_Full: "AMDLink-Windows-WeeklyFull #{num}",
    Jobs.Win_Latency: "StreamingSDK-LatencyTests #{num}",
    Jobs.Win_Long_Term: "StreamingSDK-LongTermTests #{num}",
}


def template_validation(tree) -> bool:
    # validate presence of all ids in template
    for id in ids.IDS:
        if word.find_by_id(tree, id) is None:
            return False

    return True


def prepare_working_directory():
    # remove tmp dir if exists
    if os.path.exists(WORKING_DIR_PATH):
        shutil.rmtree(WORKING_DIR_PATH)

    # remove report if exists
    if os.path.exists(REPORT_FILE_PATH):
        os.remove(REPORT_FILE_PATH)

    # copy template to the working directory
    shutil.copytree(TEMPLATE_PATH, WORKING_DIR_PATH)


def clean_working_dir():
    # remove tmp directories
    shutil.rmtree(WORKING_DIR_PATH)


def finalize_report():
    # archive directory
    shutil.make_archive(
        "report",
        "zip",
        WORKING_DIR_PATH,
    )
    # and change it extension to ".docx"
    os.rename("report.zip", REPORT_FILE_PATH)


def append_bullet_list_element_after(
    element: etree.Element, content: str, list_id
) -> etree.Element:
    bullet = word.create_bullet(list_id=list_id, lvl=0, content=content)

    # append this bullet element after specified
    word.append_element_after(new_el=bullet, after=element)

    # return new bullet element
    return bullet


def fill_task_list(tree: etree.Element, task_list_id: str, tasks: List):
    task_list_header = word.find_by_id(tree, task_list_id)

    # fill completed tasks list
    if tasks:  # fill list with tasks
        element = task_list_header
        for task in tasks:
            element = append_bullet_list_element_after(element, task, list_id=1)
    else:  # remove empty list header
        word.remove_element(task_list_header)


def fill_issues_table(tree: etree.Element, issues: List[Issue]):
    # find table by id
    table = word.find_by_id(tree, ids.ISSUES_BACKLOG_TABLE)

    # add rows to the table accordingly to data rows amount
    rows_number = len(issues)
    if rows_number > 1:
        word.table_add_rows(table, rows_number - 1)

    # copy data to the table
    table_rows = table.findall("./{*}tr")[1:]  # find all rows (skip header row)
    for row, issue in enumerate(issues):
        cells = table_rows[row].findall("./{*}tc")

        word.set_table_cell_value(cells[0], word.Link(url=issue.url, text=issue.key))
        word.set_table_cell_value(cells[1], issue.summary)
        word.set_table_cell_value(cells[2], issue.created_at)
        word.set_table_cell_value(cells[3], issue.severity)


def fill_skipped_or_observed_table(
    tree: etree.Element, table_id: str, skip_or_obs_cases_per_group: Dict[str, int]
):
    # find table by id
    table = word.find_by_id(tree, table_id)

    # add rows to the table accordingly to data rows amount
    rows_number = len(skip_or_obs_cases_per_group)
    if rows_number > 1:
        word.table_add_rows(table, rows_number - 1)

    # copy data to the table
    table_rows = table.findall("./{*}tr")[1:]  # find all rows (skip header row)
    for row, group in enumerate(skip_or_obs_cases_per_group):
        cells = table_rows[row].findall("./{*}tc")

        word.set_table_cell_value(
            cells[0],
            "{group} ({cases} cases)".format(
                group=group, cases=skip_or_obs_cases_per_group[group]
            ),
        )


def main():
    prepare_working_directory()

    # eval report dates
    report_date = datetime.today()

    # load document.xml (main xml file)
    tree = word.load_xml(word.DOCUMENT_PATH)

    # validate template
    if not template_validation(tree):
        print("Template is invalid! Some IDs are missing!")
        exit()

    ##################################################################
    # Update jobs latest run links
    print("Step 1/6 - Updating jobs' runs latest links...")

    for job in Jobs:
        link_el_id = ids.REPORT_LINKS[job]

        run_number = get_latest_build_number(job)
        if run_number is None:
            continue

        title = jobs_link_title[job].format(num=run_number)
        link = get_build_link(job, run_number)

        word.update_link(tree, link_id=link_el_id, url=link, text=title)

    ##################################################################
    # Update tasks
    print("Step 2/6 - Constructing task list...")

    summary, planned = get_project_status(report_date)

    fill_task_list(tree, ids.SUMMARY_TASK_LIST, summary)
    fill_task_list(tree, ids.PLANNED_TASK_LIST, planned)

    ##################################################################
    # Issues backlog table
    print("Step 3/6 - Constructing issue table...")

    issues = get_issues()
    fill_issues_table(tree, issues)

    ##################################################################
    # Skipped or observed tables
    print("Step 4/6 - Constructing skipped and observed tables")

    for job in ids.SKIP_OBS_CASES_TABLE:
        table_id = ids.SKIP_OBS_CASES_TABLE[job]
        skip_or_obs_cases_per_group = get_skipped_or_observed_per_group(job)
        fill_skipped_or_observed_table(tree, table_id, skip_or_obs_cases_per_group)

    ##################################################################
    # save report document.xml

    word.write_xml(tree, word.DOCUMENT_PATH)

    ##################################################################
    # update footer
    print("Step 5/6 - Updating footer...")

    # load footer.xml
    footer_tree = word.load_xml(word.FOOTER_PATH)

    report_start_date = report_date - timedelta(weeks=1) + timedelta(days=1)

    report_period_field = word.find_by_id(footer_tree, ids.REPORT_PERIOD_FIELD_ID)
    report_period_field.text = "{from_date} — {to_date}".format(
        from_date=report_start_date.strftime("%d-%B-%y"),
        to_date=report_date.strftime("%d-%B-%y"),
    )

    word.write_xml(footer_tree, word.FOOTER_PATH)

    ##################################################################
    # combine files into docx
    print("Step 6/6 - Saving report...")

    finalize_report()

    print(f"Report '{REPORT_FILE_PATH}' generated!")

    clean_working_dir()


if __name__ == "__main__":
    main()
