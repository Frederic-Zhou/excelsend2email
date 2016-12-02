package main

import (
	"bufio"
	"fmt"
	"net/smtp"
	"os"
	"regexp"

	"strings"

	"log"

	"flag"

	"github.com/tealeg/xlsx"
)

const (
	tableStyle = `
	<style type="text/css">
table.gridtable {
	font-family: verdana,arial,sans-serif;
	font-size:11px;
	color:#333333;
	border-width: 1px;
	border-color: #666666;
	border-collapse: collapse;
}
table.gridtable th {
	border-width: 1px;
	padding: 8px;
	border-style: solid;
	border-color: #666666;
	background-color: #dedede;
}
table.gridtable td {
	border-width: 1px;
	padding: 8px;
	border-style: solid;
	border-color: #666666;
	background-color: #ffffff;
}
</style>`
)

var filepath, emailuser, password, host, subject string

func main() {

	flag.StringVar(&filepath, "f", "./list.xlsx", "xlsx 文件名")
	flag.StringVar(&emailuser, "u", "", "邮件发送账号")
	flag.StringVar(&password, "p", "", "邮件发送账号的密码")
	flag.StringVar(&host, "h", "", "邮件服务器，含端口")
	flag.StringVar(&subject, "s", "邮件群发标题", "邮件标题")

	flag.Parse()

	xlFile, err := xlsx.OpenFile(filepath)
	if err != nil {
		log.Fatalln("err:", err.Error())
	}

	sendList := make(map[string]string)

	for _, sheet := range xlFile.Sheets {
		curMail := ""
		for _, row := range sheet.Rows {
			cells := getCellValues(row)
			//如果行包含电子邮件，创建一个新字典项
			if isEmail, emailStr := isEmailRow(cells); isEmail {
				curMail = emailStr
			} else {
				count := 0
				for _, c := range cells {
					if len(c) > 0 {
						count++
					}
				}

				if count > 1 {
					sendList[curMail] += fmt.Sprintf("<tr><td>%s</td></tr>", strings.Join(cells, "</td><td>"))
				} else {
					sendList[curMail] += fmt.Sprintf("<tr><td colspan='%d'>%s</td></tr>", len(cells), strings.Join(cells, ""))
				}

			}

		}
	}

	sendMail(sendList)
	fmt.Print("按下回车结束")
	bufio.NewReader(os.Stdin).ReadLine()

}

func getCellValues(r *xlsx.Row) (cells []string) {
	for _, cell := range r.Cells {
		txt := strings.Replace(strings.Replace(cell.Value, "\n", "", -1), " ", "", -1)
		cells = append(cells, txt)
	}
	return
}

func isEmailRow(r []string) (isEmail bool, email string) {
	// 查找连续的小写字母
	reg := regexp.MustCompile(`^[a-zA-Z_0-9.-]{1,64}@([a-zA-Z0-9-]{1,200}.){1,5}[a-zA-Z]{1,6}$`)
	for _, v := range r {
		if reg.MatchString(v) {
			return true, v
		}
	}
	return false, ""
}

func sendMail(sendList map[string]string) {

	fmt.Printf("共需要发送%d封邮件\n", len(sendList))
	index := 1
	for mail, content := range sendList {
		fmt.Printf("发送第%d封", index)
		if err := sendToMail(emailuser,
			password,
			host, // "smtp.mxhichina.com:25"
			mail,
			subject,
			fmt.Sprintf("%s<table class='gridtable'>%s</table>", tableStyle, content),
			"html"); err != nil {
			fmt.Printf(" ... 发送错误(X) %s %s \n", mail, err.Error())

		} else {
			fmt.Printf(" ... 发送成功(V) %s \n", mail)
		}
		index++
		//fmt.Printf("<table border='2'>%s</table> \n", content)
	}

}

func sendToMail(user, password, host, to, subject, body, mailtype string) error {
	auth := smtp.PlainAuth("", user, password, strings.Split(host, ":")[0])
	msg := []byte("To: " + to + "\r\nFrom: " + user + "\r\nSubject: " + subject + "\r\n" + "Content-Type: text/" + mailtype + "; charset=UTF-8" + "\r\n\r\n" + body)
	sendto := strings.Split(to, ";")
	err := smtp.SendMail(host, auth, user, sendto, msg)
	return err
}
