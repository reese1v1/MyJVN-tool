from utility import my_utility
from queue import Queue
import threading
import time
import sys

DEBUG = False  # debug flag

exitFlag = 0  # thread task flag
queueLock = threading.Lock()  # queue lock
workQueue = Queue()  # thread task queue

results_table = []  # A global table for output results


class ProcessWorker(threading.Thread):
    def __init__(self, q):
        threading.Thread.__init__(self)
        # self.threadID = threadID
        # self.name = name
        self.q = q

    def run(self):
        process_data(self.q)


def process_data(q):
    while not exitFlag:
        queueLock.acquire()
        if not q.empty():
            data = str(q.get(block=False))
            queueLock.release()

            # if has oss data
            if len(data) > 0:
                try:
                    investigating(data)
                    print(data, "... ok")
                except Exception as err:
                    # print('Exception in: ', data, err)
                    # workQueue.put(data)
                    print('Error in: ', data)
                    results_table.extend([[data, 'エラー発生!', '-', '-', '-', '-', '-']])
                    time.sleep(1)
                    continue
                finally:
                    q.task_done()

            # if has not oss data(null)
            else:
                q.task_done()
        else:
            queueLock.release()
        # make a delay between each loop
        time.sleep(1)


def get_detail_list(detail_url):
    detail_data = my_utility.get_xml(detail_url) \
        .find(my_utility.XML_DETAIL_INFO) \
        .find(my_utility.XML_DETAIL_DATA)

    detail_list = [detail_data.find(my_utility.XML_DETAIL_DATA_DES),
                   detail_data.find(my_utility.XML_DETAIL_DATA_IMPACT),
                   detail_data.find(my_utility.XML_DETAIL_DATA_AFFECTED)]

    return [''.join(item.itertext()).strip() for item in detail_list]


def overview_map_func(list_item):
    detail_list = get_detail_list(my_utility.detail_url(list_item[2]))
    return [list_item[0], list_item[1], detail_list[0], detail_list[1], detail_list[2]]


def get_overview_list(overview_url):
    overview_root = my_utility.get_xml(overview_url)
    status = overview_root.find(my_utility.XML_STATUS)

    if int(status.get(my_utility.XML_TOTAL_RES)) > 0:
        overview_list = [[item.find(my_utility.XML_TITLE).text,
                          item.find(my_utility.XML_LINK).text,
                          item.find(my_utility.XML_IDENTIFIER).text]
                         for item in overview_root.findall(my_utility.XML_ITEM)]

        return list(map(overview_map_func, overview_list))


def investigating(oss):
    result_list = get_overview_list(my_utility.overview_url(oss))

    # if len(result_list) > 0:
    if result_list is not None:
        return results_table.extend(my_utility.list_parse(result_list, oss))

    else:
        return results_table.extend([[oss, '脆弱性なし', '-', '-', '-', '-', '-']])


def main():
    workers = 1 if DEBUG else 4  # the number of request threads

    if len(sys.argv) > 3:
        # read excel
        print('インプットファイル: ' + sys.argv[1] + '\n')
        # get oss data
        oss_data = my_utility.get_oss_data(sys.argv[1])

        if oss_data is not None:
            # get config data
            if my_utility.get_config_data():
                print('公表日: ' + my_utility.DATE_PUBLIC_FROM_YEAR + my_utility.DATE_PUBLIC_FROM_MONTH)
                print('脆弱性情報の検索を開始します。')
                ts = time.time()

                for x in range(int(workers)):
                    worker = ProcessWorker(workQueue)
                    worker.daemon = True
                    worker.start()

                # put into queue
                # queueLock.acquire()
                for row in oss_data:
                    workQueue.put(row)
                # queueLock.release()

                # waiting for queue's empty
                while not workQueue.empty():
                    pass

                # notifying the thread to quit
                global exitFlag
                exitFlag = 1

                # waiting for all threads to finish tasks
                workQueue.join()
                print("Exiting Main Thread")
                print("消費時間: {:.2f}s".format(time.time() - ts))

                if not DEBUG:
                    my_utility.output_csv(results_table)

    else:
        print('引数を入力してください: [ファイル名, 公表日 年, 公表日 月]')


if __name__ == '__main__':
    main()
