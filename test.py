import sys
import getopt


def main(argv):
    fromEmail = ''
    passworld = ''
    helpMsg = "test.py\n" \
              "\ttest.py -f <fromEmail> -k <pop3Key>\n" \
              "\t-f\t\t\t\t邮件发送方邮箱地址\n" \
              "\t--fromEmail\n" \
              "\t-k\t\t\t\t邮箱POP3授权码\n" \
              "\t--pop3Key\n" \
              "\t-h\t\t\t\thelp\n" \
              "\t--help\n"
    try:
       options, args = getopt.getopt(argv, "hf:k:", ["help", "fromEmail=", "pop3Key="])
    except getopt.GetoptError:
        print(helpMsg)
        sys.exit()
    if not len(options):
        print(helpMsg)
        sys.exit()
    for option, value in options:
        if option in ("-h", "--help"):
            print(helpMsg)
        elif option in ("-f", "--fromEmail"):
            fromEmail = value
            print('Email: ', fromEmail)
        elif option in ("-k", "--pop3Key"):
            passworld = value
            print('passworld: ', passworld)


if __name__ == "__main__":
    main(sys.argv[1:])