import os
from datetime import datetime
from atlassian import Jira
from common import Issue
from typing import List
from urllib.parse import urljoin

JIRA_URL = "https://luxproject.luxoft.com/jira/"
JIRA_TOKEN = os.environ["LUXOFT_JIRA_TOKEN"]

jira_instance = Jira(
    # Url of jira server
    url=JIRA_URL,
    # password/token
    token=JIRA_TOKEN,
    cloud=False,
)

def validate_token():
    issues = jira_instance.jql("")
    if issues['total'] == 0:
        print("ERROR: Jira token 'JIRA_TOKEN' is invalid!")
        exit(-1)


# validate token on module's load
validate_token()

def get_issues() -> List[Issue]:
    jql_request = 'project = STVITT AND issuetype = Defect AND status in (Open, "In Progress", Suspended, Resolved, Deferred) AND labels = StreamingSDK'

    issues = jira_instance.jql(
        jql_request, fields="summary,customfield_12094,created"
    ).get("issues")

    issues = [
        Issue(
            key=issue["key"],
            summary=issue["fields"]["summary"],
            created_at=datetime.strptime(
                issue["fields"]["created"].split("T")[0], "%Y-%m-%d"
            ).strftime("%d/%b/%y"),
            severity=issue["fields"]["customfield_12094"]["value"].split(" ")[-1],
            url=urljoin(urljoin(JIRA_URL, "browse/"), issue['key'])
        )
        for issue in issues
    ]

    return issues


if __name__ == "__main__":
    issues = get_issues()
    print("Issues:")
    for issue in issues:
        print(f"{issue.key}: ({issue.severity}) : {issue.summary} [{issue.created_at}] ({issue.url})")
