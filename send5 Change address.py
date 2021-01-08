import pytz
import time
import datetime
from email.mime.text import MIMEText
import smtplib
import random
import sys
import getopt


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
def send_result(text, fromEmail, passworld, userEmail):
    if userEmail == 'null':
        print('用户指定不发送邮件')
        return
    # nowtime = datetime.datetime.now().strftime('%m-%d')
    # msg_from = ''  # 发送邮箱
    msg_from = fromEmail
    # passwd = ''  # 密码
    passwd = passworld
    msg_to = userEmail  # 目的邮箱

    subject = '安全上报结果'
    # content = text + '\n' + nowtime
    content = text
    msg = MIMEText(content)
    msg['Subject'] = subject
    msg['From'] = msg_from
    msg['To'] = msg_to

    s = smtplib.SMTP_SSL("smtp.qq.com", 465)
    s.login(msg_from, passwd)
    s.sendmail(msg_from, msg_to, msg.as_string())
    print("邮件发送成功")
    s.quit()


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
    fromEmail, pop3Key = parse_options(argv)
    username, password, userEmail, name, sum = get_info_from_txt()
    print('---------------------------')
    # 服务器当前时间
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    # 中国地区当前时间，PRC
    print(datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S'))
    print('账号密码读取成功，共', sum, '人')
    print('---------------------------')
    # 本次发送人数
    reported = 0
    # 发送给用户的邮件信息
    userString = []
    # 发送给admin的邮件信息
    adminString = []

    nowServerTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    nowTime = datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S')
    content = '服务器当前时间：' + nowServerTime + '\n' + '东八区当前时间：' + nowTime + '\n'
    adminString.append(content)

    adminString.append('---------------------------')
    adminString.append(datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S'))
    adminString.append('账号密码读取成功，共' + str(sum) + '人')
    adminString.append('---------------------------')
    try:
        for i in range(sum):
            print('')
            print('No.' + str(i + 1) + '  ' + datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S'))
            adminString.append('No.' + str(i + 1) + '  ' + datetime.datetime.now(pytz.timezone('PRC')).strftime('%Y-%m-%d %H:%M:%S'))

            text = '用户 ' + str(name[i]) + '-' + str(username[i])
            print(text)
            userString.append(text)
            adminString.append(text + '\n')

            text = '    由于代码使用GitHub Action功能来运行，' \
                   '服务器位于国外，加上学校安全上报网站服务器不稳定，' \
                   '存在上报失败的情况，还请大家谅解。' \
                   + '\n' + \
                   '    近期大家陆续都已离校，故上报信息中大家的所在地需要修改。' \
                   '只需要进入上报界面修改最近一天的位置信息，重新提交一次即可。' \
                   + '\n' + \
                   '    上报网址：http://yiqing.ctgu.edu.cn/ '

            userString.append(text)
            # 将用户信息，邮件发送给用户
            send_result('\n'.join(userString), fromEmail, pop3Key, userEmail[i])
            # 调用clear()方法，清空列表，避免其出现在下一个用户的邮件中
            userString.clear()
            # 本次发送人数统计
            reported += 1

        # 调用clear()方法，清空列表，避免其出现在下一个用户的邮件中
        userString.clear()
        # 增大等待间隔，github访问速度慢
        time.sleep(random.random(10, 20))

    finally:
        text = '应发送:' + str(sum) + ' 本次发送:' + str(reported)
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
        adminString.clear()


if __name__ == "__main__":
    main(sys.argv[1:])
