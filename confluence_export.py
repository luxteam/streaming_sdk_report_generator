import os
import requests
from datetime import datetime, timedelta
from lxml import html
import json
from bs4 import BeautifulSoup
import re

CONFLUENCE_TOKEN = os.environ["CONFLUENCE_TOKEN"]

def validate_token():
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {CONFLUENCE_TOKEN}",
    }

    response = requests.get(
        "https://luxproject.luxoft.com/confluence/rest/api/user/current",
        headers=headers,
    )

    if response.json()['type'] == "anonymous": 
        print("ERROR: Confluence token 'CONFLUENCE_TOKEN' is invalid!")
        exit(-1)


# validate token on module's load
validate_token()


def _request_confluence_report(report_date: datetime) -> html.Element:
    url = "https://luxproject.luxoft.com/confluence/rest/api/content"

    confluence_report_date = report_date + timedelta(days=1) 

    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {CONFLUENCE_TOKEN}",
    }

    response = requests.get(
        f"{url}/?title=Status Report - {confluence_report_date.strftime('%Y-%m-%d')}&expand=body.storage",
        headers=headers,
    )

    # Taking lates available report
    days = 0
    while response.json()["size"] == 0 and days < 7:
        days += 1
        date = confluence_report_date - timedelta(days=days)
        response = requests.get(
            f"{url}/?title=Status Report - {date.strftime('%Y-%m-%d')}&expand=body.storage",
            headers=headers,
        )
    
    if days >= 7:
        print("ERROR: Confluence token is invalid!")
        exit(-1)

    page_content = response.json()["results"][0]["body"]["storage"]["value"]
    soup = BeautifulSoup(page_content, 'lxml')

    return soup


def get_project_status(report_date: datetime):
    page = _request_confluence_report(report_date)

    data = []
    table_body = page.find('tbody')
    rows = table_body.find_all('tr')
    for row in rows:
        cell1 = row.find(string=re.compile("SSDK:"))
        cell2 = row.find(string=re.compile("DONE"))
        try:
            cell_data1 = cell1.get_text()
            cell_data2 = cell2.get_text()
            data.append(cell_data1)
        except:
            continue

    return data


if __name__ == "__main__":
    summary = get_project_status(datetime.now())

    print("Summary:")
    print(json.dumps(summary, indent=4))

