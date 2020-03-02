from requests_xml import XMLSession
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import urllib3
import xlrd
import pandas as pd
import sys
import datetime
import time

try:
    import xml.etree.cElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET

# disable the SSL warnings
urllib3.disable_warnings()

HTTP_PROXY = HTTPS_PROXY = 'http://proxy.toshiba-sol.co.jp:8080'
XML_STATUS = '{http://jvndb.jvn.jp/myjvn/Status}Status'
XML_ITEM = '{http://purl.org/rss/1.0/}item'
XML_IDENTIFIER = '{http://jvn.jp/rss/mod_sec/3.0/}identifier'
XML_LINK = '{http://purl.org/rss/1.0/}link'
XML_TITLE = '{http://purl.org/rss/1.0/}title'
XML_DETAIL_INFO = '{http://jvn.jp/vuldef/}Vulinfo'
XML_DETAIL_DATA = '{http://jvn.jp/vuldef/}VulinfoData'
XML_DETAIL_DATA_DES = '{http://jvn.jp/vuldef/}VulinfoDescription'
XML_DETAIL_DATA_AFFECTED = '{http://jvn.jp/vuldef/}Affected'
XML_DETAIL_DATA_IMPACT = '{http://jvn.jp/vuldef/}Impact'
XML_TOTAL_RES = 'totalRes'

COLUMN_NUMBER = 0
COLUMN_NAME = 'name'

DATE_PUBLIC_FROM_YEAR = '2018'  # 公表日 年
DATE_PUBLIC_FROM_MONTH = '09'  # 公表日 月

RESULT_TABLE_TITLE = ['oss', '脆弱性発見', 'タイトル', 'link', '概要', '深刻度', '影響']


# def find_files():
#     # file_dir = sys.path[0]
#     file_type = '.xls'
#
#     for files in os.listdir('.'):
#         if os.path.isfile(files) and (file_type in os.path.splitext(files)[1]):
#             return files
#
#     return ''


def get_xls(path):
    try:
        oss_sheet = xlrd.open_workbook(path).sheet_by_index(0)
        # get oss data by col
        return oss_sheet.col_values(COLUMN_NUMBER)

    except Exception as ex:
        print('file not exist', ex)


def get_csv(path):
    try:
        df = pd.read_csv(path)
        return df[COLUMN_NAME]

    except Exception as ex:
        print('file not exist', ex)


def get_oss_data(path):
    return get_csv(path)


def get_config_data():
    global DATE_PUBLIC_FROM_YEAR
    global DATE_PUBLIC_FROM_MONTH

    try:
        datetime.datetime.strptime(str(sys.argv[2] + sys.argv[3]), '%Y%m')
        DATE_PUBLIC_FROM_YEAR = sys.argv[2]
        DATE_PUBLIC_FROM_MONTH = sys.argv[3]
        return True

    except ValueError as err:
        print('公表日を再入力してください: YYYY MM')


def list_parse(lst, oss):
    # if has results data
    for row in lst:
        row.insert(0, oss)
        row.insert(1, '脆弱性あり')
    return lst


def overview_url(oss):
    return "https://jvndb.jvn.jp/myjvn?method=getVulnOverviewList&feed=hnd&keyword=%22" + oss + \
           "&dateFirstPublishedStartY=" + DATE_PUBLIC_FROM_YEAR + \
           "&dateFirstPublishedStartM=" + DATE_PUBLIC_FROM_MONTH + \
           "&rangeDatePublic=n&rangeDatePublished=n&rangeDateFirstPublished=n"


def detail_url(jvn_id):
    return "https://jvndb.jvn.jp/myjvn?method=getVulnDetailInfo&feed=hnd&vulnId=" + jvn_id


def get_xml(url):
    session = XMLSession()
    session.proxies = {'http': HTTP_PROXY, 'https': HTTPS_PROXY}
    retry = Retry(connect=3, backoff_factor=1)
    adapter = HTTPAdapter(max_retries=retry)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) '
                      'AppleWebKit/603.3.8 (KHTML, like Gecko) Version/10.1.2 Safari/603.3.8',
        'Accept-Encoding': 'gzip, deflate', 'Accept': '*/*', 'Connection': 'keep-alive'}
    session.mount('http://', adapter)
    r = session.get(url, headers=headers, verify=False)
    session.close()

    # results = []
    # if int(status.get('totalRes')) > 0:
    #     for elem in root.findall('{http://purl.org/rss/1.0/}item'):
    #         temp = []
    #         temp.append(elem.find('{http://purl.org/rss/1.0/}title').text)
    #         temp.append(elem.find('{http://purl.org/rss/1.0/}link').text)
    #         results.append(temp)
    return ET.fromstring(r.content)


def output_csv(data):
    titles = RESULT_TABLE_TITLE
    df = pd.DataFrame(data, columns=titles)
    file_path = 'output-' + time.strftime("%Y%m%d-%H%M%S", time.localtime()) + '.csv'

    try:
        df.to_csv(file_path, encoding='Shift_JIS', index=False)
        print("結果ファイルを確認ください: ", file_path)

    except Exception as err:
        print('Exception: ', err)


def main():
    pass


if __name__ == '__main__':
    main()
