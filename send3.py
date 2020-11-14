import sys
import getopt
import requests
import time
import datetime
import smtplib
import random
import pytz
from bs4 import BeautifulSoup
from email.mime.text import MIMEText




# 从txt文件中读取数据
def get_info_from_txt():
    username = []
    password = []
    userEmail = []
    name = []
    sum = 0
    configFile = open('config.txt', 'r', encoding='utf-8')
    list = configFile.readlines()
    # print(list)
    for lists in list:
        lists = lists.strip()
        lists = lists.strip('\n')
        lists = lists.strip(',')
        lists = lists.split(',')
        # print(lists[0])
        username.append(lists[0])
        # print(lists[1])
        password.append(lists[1])
        # print(lists[2])
        userEmail.append(lists[2])
        # print(lists[3])
        name.append(lists[3])
        sum += 1
    # print('sum: ', sum)
    return username, password, userEmail, name, sum


# 发邮件通知结果
def send_result(text, fromEmail, passWorld, userEmail):
    if userEmail == 'null':
        print('用户指定不发送邮件')
        return
    nowServerTime = datetime.datetime.now().strftime('%Y-%m-%d')
    nowTime = datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d')
    msg_from = fromEmail
    passwd = passWorld
    msg_to = userEmail  # 目的邮箱

    subject = '安全上报结果'
    content = '服务器当前时间：' + nowServerTime + '\n' + '东八区当前时间：' + nowTime + '\n' + text
    msg = MIMEText(content)
    msg['Subject'] = subject
    msg['From'] = msg_from
    msg['To'] = msg_to

    s = smtplib.SMTP_SSL("smtp.qq.com", 465)
    s.login(msg_from, passwd)
    s.sendmail(msg_from, msg_to, msg.as_string())
    print("邮件发送成功")
    s.quit()


# 命令行参数
def parse_options(argv):
    fromEmail = ''
    pop3Key = ''
    helpMsg = "send.py\n" \
              "\tsend.py -f <fromEmail> -k <pop3Key>\n" \
              "\t-f\t\t\t\t邮件发送方邮箱地址\n" \
              "\t--fromEmail\n" \
              "\t-k\t\t\t\t邮箱POP3授权码\n" \
              "\t--pop3Key\n" \
              "\t-h\t\t\t\thelp\n" \
              "\t--help\n"

    if not ((('-f' in argv) or ('--fromEmail' in argv)) and (('-k' in argv) or ('--pop3Key' in argv))):
        print(helpMsg)
        sys.exit()

    argc = 0
    if ('-h' in argv) or ('--help' in argv):
        argc += 1
    if ('-f' in argv) or ('--fromEmail' in argv):
        argc += 2
    if ('-f' in argv) or ('--fromEmail' in argv):
        argc += 2
    if len(argv) != argc:
        print(helpMsg)
        sys.exit()

    try:
        opts, args = getopt.getopt(argv, "hf:k:", ["help", "fromEmail=", "pop3Key="])
    except getopt.GetoptError:
        print(helpMsg)
        sys.exit(2)

    if not len(opts):
        print(helpMsg)
        sys.exit()

    for option, value in opts:
        if option in ("-h", "--help"):
            print(helpMsg)
        elif option in ("-f", "--fromEmail"):
            if value is None:
                sys.exit()
            fromEmail = value
        elif option in ("-k", "--pop3Key"):
            if value is None:
                sys.exit()
            pop3Key = value
    return fromEmail, pop3Key


def main(argv):
    # 获取发送方邮箱地址和pop3授权码
    fromEmail, pop3Key = parse_options(argv)
    # 获取用户账号密码等信息
    username, password, userEmail, name, sum = get_info_from_txt()

    print('---------------------------')
    print(datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S'))
    print('账号密码读取成功，共', sum, '人')
    print('---------------------------')

    # 今日已上报人数
    finished = 0
    # 本次上报人数
    reported = 0
    # 发送给用户的邮件信息
    userString = []
    # 发送给admin的邮件信息
    adminString = []

    adminString.append('---------------------------')
    adminString.append(datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S'))
    adminString.append('账号密码读取成功，共' + str(sum) + '人')
    adminString.append('---------------------------')

    try:
        for i in range(sum):
            print('')

            headers = {
                'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Mobile Safari/537.36'
            }
            s = requests.Session()

            # 登录
            login_url = 'http://yiqing.ctgu.edu.cn/wx/index/loginSubmit.do'
            data = {'username': username[i], 'password': password[i]}
            r = s.post(url=login_url, headers=headers, data=data)
            r.encoding = 'utf-8'
            state = r.text

            print('No.' + str(i + 1) + '  ' + datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S'))

            adminString.append('No.' + str(i + 1) + '  ' + datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S'))

            if state == 'success':
                print('用户', name[i], username[i], '登录成功')
                # 记录用户登录成功，添加到userString和adminString
                text = '用户 ' + str(name[i]) + str(username[i]) + '登录成功'
                userString.append(text)
                adminString.append(text)
            else:
                print('用户', name[i], username[i], '密码错误，登录失败')
                # 记录用户登录失败，添加到userString和adminString
                text = '用户 ' + str(name[i]) + str(username[i]) + '密码错误，登录失败'
                userString.append(text + '\n')
                adminString.append(text + '\n')
                # 将登录失败信息 邮件发送给用户
                # send_result('\n'.join(userString), fromEmail, pop3Key, userEmail[i])
                # 调用clear()方法，清空列表，避免其出现在下一个用户的邮件中
                # print('userString:---------  : ', userString)
                userString.clear()
                continue

            # 获取带有data的html内容
            student_info_url = 'http://yiqing.ctgu.edu.cn/wx/health/toApply.do'
            r = s.get(url=student_info_url, headers=headers)
            r.encoding = 'utf-8'
            html = r.text

            time.sleep(random.random())

            try:
                # 解析html，获得data数据
                bs = BeautifulSoup(html, 'lxml')
                data = {
                    # ttoken
                    'ttoken': bs.find(attrs={'name': 'ttoken'})['value'],
                    # 省份
                    'province': bs.find(attrs={'name': 'province'})['value'],
                    # 城市
                    'city': bs.find(attrs={'name': 'city'})['value'],
                    # 县区
                    'district': bs.find(attrs={'name': 'district'})['value'],
                    # 地区代码：身份证号前六位
                    'adcode': bs.find(attrs={'name': 'adcode'})['value'],
                    # 经度
                    'longitude': bs.find(attrs={'name': 'longitude'})['value'],
                    # 纬度
                    'latitude': bs.find(attrs={'name': 'latitude'})['value'],
                    # 是否确诊新型肺炎
                    'sfqz': bs.find(attrs={'name': 'sfqz'})['value'],
                    # 是否疑似感染
                    'sfys': bs.find(attrs={'name': 'sfys'})['value'],
                    'sfzy': bs.find(attrs={'name': 'sfzy'})['value'],
                    # 是否隔离
                    'sfgl': bs.find(attrs={'name': 'sfgl'})['value'],
                    # 状态，statusName: "正常"
                    'status': bs.find(attrs={'name': 'status'})['value'],
                    'sfgr': bs.find(attrs={'name': 'sfgr'})['value'],
                    'szdz': bs.find(attrs={'name': 'szdz'})['value'],
                    # 手机号
                    'sjh': bs.find(attrs={'name': 'sjh'})['value'],
                    # 紧急联系人姓名
                    'lxrxm': bs.find(attrs={'name': 'lxrxm'})['value'],
                    # 紧急联系人手机号
                    'lxrsjh': bs.find(attrs={'name': 'lxrsjh'})['value'],
                    # 是否发热
                    'sffr': '否',
                    'sffy': '否',
                    'sfgr': '否',
                    'qzglsj': '',
                    'qzgldd': '',
                    'glyy': '',
                    'mqzz': '',
                    # 是否返校
                    'sffx': '否',
                    # 其它
                    'qt': '',
                }

                # 向服务器post数据
                sent_info_url = 'http://yiqing.ctgu.edu.cn/wx/health/saveApply.do'
                s.post(url=sent_info_url, headers=headers, data=data)

                # 记录用户上报成功，添加到userString和adminString
                text = '用户 ' + name[i] + '-' + str(username[i]) + ' 上报成功'

                print(text)

                userString.append(text)
                adminString.append(text + '\n')

                # 将用户上报成功信息，邮件发送给用户
                # send_result('\n'.join(userString), fromEmail, pop3Key, userEmail[i])
                # 调用clear()方法，清空列表，避免其出现在下一个用户的邮件中
                userString.clear()
                # 本次上报人数和今日已上报人数统计
                reported += 1
                finished += 1
            except:
                text = '用户 ' + name[i] + '-' + str(username[i]) + ' 今日已上报'
                print(text)
                adminString.append(text + '\n')
                # 今日已上报统计
                reported += 1
            # 调用clear()方法，清空列表，避免其出现在下一个用户的邮件中
            userString.clear()
            # 增大等待间隔，github访问速度慢
            time.sleep(random.random())
    finally:
        text = '应报:' + str(sum) + ' 本次上报:' + str(finished) + ' 今日已上报:' + str(reported)
        adminString.append('---------------------------')
        adminString.append(text)
        adminString.append('---------------------------')
        # print('\n'.join(adminString))
        print('---------------------------')
        print(text)
        send_result('\n'.join(adminString), fromEmail, pop3Key, 'yizhaosan@qq.com')
        print('---------------------------\n')
        # 调用clear()方法，清空列表，避免其出现在下次的邮件中
        # print('adminString:---------  : ', adminString)
        userString.clear()


if __name__ == "__main__":
    main(sys.argv[1:])
