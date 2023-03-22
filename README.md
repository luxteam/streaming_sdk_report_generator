# Streaming SDK report generator

## Installation
```
pip install -r requirements.txt
```

Set required environment variables:
- CONFLUENCE_TOKEN
- JENKINS_USERNAME
- JENKINS_TOKEN
- LUXOFT_JIRA_TOKEN


## Usage

### Generate report:
```
python3 ./gen_report.py
```

Result: `./report.docx`



### Generate emails:
```
python3 ./gen_emails.py
```

Result: `./Letter_1.oft` and `./Letter_2.oft`
