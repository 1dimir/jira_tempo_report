from getpass import getpass
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta
import xlsxwriter
import config

base_url = ('{jira_base_url}'
            'rest/tempo-timesheets/3/worklogs/?'
            'projectKey={project_key}&username={{user_name}}&'
            'dateFrom={{date_from}}&toDate={{date_till}}'
            ).format(
                jira_base_url=config.jira_base_url,
                project_key=config.jira_project_key
            )

issue_url = '{jira_base_url}rest/api/2/issue/{{issue_key}}'.format(jira_base_url=config.jira_base_url)


def get_issue_description(issue_key, auth):

    response = requests.get(
        issue_url.format(issue_key=issue_key),
        verify=False,
        auth=auth
    )

    if response.status_code != 200:
        return ''

    issue_json = response.json()
    fields = issue_json.get('fields')

    summary = fields.get('summary')

    if 'parent' in fields:
        result = u'{parent_key} ({issue_key}) {summary}'.format(
            parent_key=fields.get('parent').get('key'),
            issue_key=issue_key,
            summary=summary
        )
    else:
        result = u'{issue_key} {summary}'.format(
            issue_key=issue_key,
            summary=summary
        )
    return result


def process_work_logs(work_logs_json):
    time_spent = dict()
    issues_by_date = dict()
    for work_log in work_logs_json:
        log_date = datetime.strptime(work_log.get('dateStarted'), '%Y-%m-%dT%H:%M:%S.%f').date()
        seconds_worked = work_log.get('timeSpentSeconds', 0)

        if log_date in time_spent:
            time_spent[log_date] += seconds_worked
        else:
            time_spent[log_date] = seconds_worked

        if log_date not in issues_by_date:
            issues_by_date[log_date] = []

        issue_key = work_log.get('issue').get('key')
        if issue_key not in issues_by_date.get(log_date):
            issues_by_date.get(log_date).append(issue_key)
    return time_spent, issues_by_date


def time_spent_presentation(time_in_seconds):
    """

    :param time_in_seconds:
    :return: string, represents time in format 3h 45m 12s
    """

    result = ''

    spent_hours = time_in_seconds / 3600
    spent_minutes = (time_in_seconds % 3600) / 60
    spent_seconds = time_in_seconds % 60

    if spent_hours > 0:
        result += '{hours}h '.format(hours=spent_hours)
    if spent_minutes > 0:
        result += '{minutes}m '.format(minutes=spent_minutes)
    if spent_seconds > 0:
        result += '{seconds}s '.format(seconds=spent_seconds)

    return result.strip()


def usage():
    print 'report.py '


if __name__ == "__main__":

    user = raw_input('User:')
    password = getpass('Password for {user}:'.format(user=user))

    today = datetime.today().date()
    if today.weekday() <= 2:
        date_from = today - timedelta(7 + today.weekday())
        date_till = today - timedelta(today.weekday() + 1)
    else:
        date_from = today - timedelta(today.weekday())
        date_till = today + timedelta(6 - today.weekday())

    url = base_url.format(
        user_name=user,
        date_from=date_from.isoformat(),
        date_till=date_till.isoformat()
    )

    print url

    auth = HTTPBasicAuth(user, password)

    response = requests.get(url, verify=False, auth=auth)
    if response.status_code != 200:
        raise RuntimeError

    time_spent, issues_by_date = process_work_logs(response.json())

    workbook = xlsxwriter.Workbook('report.xlsx')
    worksheet = workbook.add_worksheet()

    index = 0
    for key in sorted(time_spent):

        worksheet.write(index, 0, key.isoformat())
        worksheet.write(index, 1, time_spent_presentation(time_spent.get(key)))

        comments = []
        for issue_key in issues_by_date.get(key):
            comments.append(get_issue_description(issue_key, auth))

        comments.sort()

        for comment in comments:
            worksheet.write(index, 2, comment)
            index += 1

        index += 1

    workbook.close()
