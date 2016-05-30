# -*- coding: utf-8 -*-
from xml.dom.minidom import parse, parseString # minidom 모듈의 파싱 함수를 임포트합니다.
from xml.etree import ElementTree
from http.client import HTTPConnection
from http.server import BaseHTTPRequestHandler, HTTPServer

##### global
loopFlag = 1
xmlFD = -1
ExamDoc = None

conn = None
regkey = "HK%2FAjOVzmL%2BcwzaN%2BLp4GK%2FeAPo%2BSnXoEUyQdtJAoif3MSkh94fmqLGGx%2BrtKEVY%2F8gz2BudoufnWL7Q6jXnzg%3D%3D"
server = "openapi.q-net.or.kr"

host = "smtp.gmail.com"
port = 587

#### Menu  implementation
def printMenu():
    print("========Menu==========")
    print("프로그램 종료 : q")
    print("기능사 시험 등록 일자 : t")
    print("기사/산업기사 시험 등록 일자 : a")
    print("E-Mail 전송 : m")
    print("======================")
    
def launcherFunction(menu):
    if menu == 'q':
        QuitExamMgr()
    elif menu == 'p':
        PrintDOMtoXML()
    elif menu == 't': 
        ret = getExamDataFromReg(1)
        print("start day is ", ret["StartDay"])
        print("end day is ", ret["EndDay"])
    elif menu == 'a': 
        ret = getExamDataFromReg(2)
        print("start day is ", ret["StartDay"])
        print("end day is ", ret["EndDay"])
    else:
        print ("error : unknow menu key")
        
def LoadXMLFromFile():
    global xmlFD
 
    try:
        xmlFD = open("tech.xml", encoding="utf-8")
    except IOError:
        print ("invalid file name or path")
        return None
    else:
        try:
            dom = parse(xmlFD)
        except Exception:
            print ("loading fail!!!")
        else:
            return dom
    return None

def QuitExamMgr():
    global loopFlag
    loopFlag = 0
    ExamFree()
    
def ExamFree():
    if checkDocument():
        ExamDoc.unlink()
        
def PrintDOMtoXML():
    if checkDocument():
        print(ExamDoc.toxml())
        
def PrintExamList(tags):
    global ExamDoc
    ExamDoc = LoadXMLFromFile()
    if not checkDocument():
        return None
    
    response = ExamDoc.childNodes
    body = response[0].childNodes
    items = body[1].childNodes
    item = items[0].childNodes
    for it in item:
        if it.nodeName == "item":
            subitem = it.childNodes
            for atom in subitem:
                if atom.nodeName in tags:
                    print("Doc Reg Start Day : ", atom.firstChild.nodeValue)
    
def checkDocument():
    global ExamDoc
    if ExamDoc == None:
        return False
    return True
    
def connectOpenAPIServer():
    global conn, server
    conn = HTTPConnection(server)
    
def getExamDataFromReg(num):
    global server, regKey, conn
    if conn == None :
        connectOpenAPIServer()
    uri1 = "http://openapi.q-net.or.kr/api/service/rest/InquiryTestInformationNTQSVC/getCList?ServiceKey=HK%2FAjOVzmL%2BcwzaN%2BLp4GK%2FeAPo%2BSnXoEUyQdtJAoif3MSkh94fmqLGGx%2BrtKEVY%2F8gz2BudoufnWL7Q6jXnzg%3D%3D"
    uri2 = "http://openapi.q-net.or.kr/api/service/rest/InquiryTestInformationNTQSVC/getEList?ServiceKey=HK%2FAjOVzmL%2BcwzaN%2BLp4GK%2FeAPo%2BSnXoEUyQdtJAoif3MSkh94fmqLGGx%2BrtKEVY%2F8gz2BudoufnWL7Q6jXnzg%3D%3D"
    if num == 1:
        conn.request("GET", uri1)
    elif num == 2:
        conn.request("GET", uri2)    
    
    req = conn.getresponse()
    if int(req.status) == 200 :
        return extractBookData(req.read())
    else:
        print ("OpenAPI request has been failed!! please retry")
        return None
        
def extractBookData(strXml):
    tree = ElementTree.fromstring(strXml)
   
    itemElements = tree.getiterator("item")  # return list type
    for item in itemElements:
        ExamStDay = item.find("docRegStartDt")
        ExamEnDay = item.find("docRegEndDt")
        return {"StartDay":ExamStDay.text,"EndDay":ExamEnDay.text}
        
def sendMain():
    global host, port
    html = ""
    title = str(input ('Title :'))
    senderAddr = str(input ('sender email address :'))
    recipientAddr = str(input ('recipient email address :'))
    msgtext = str(input ('write message :'))
    passwd = str(input (' input your password of gmail account :'))
    msgtext = str(input ('Do you want to include book data (y/n):'))
    if msgtext == 'y' :
        keyword = str(input ('input keyword to search:'))
        html = MakeHtmlDoc(SearchBookTitle(keyword))
    
    import mysmtplib
    # MIMEMultipart의 MIME을 생성합니다.
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    
    #Message container를 생성합니다.
    msg = MIMEMultipart('alternative')

    #set message
    msg['Subject'] = title
    msg['From'] = senderAddr
    msg['To'] = recipientAddr
    
    msgPart = MIMEText(msgtext, 'plain')
    bookPart = MIMEText(html, 'html', _charset = 'UTF-8')
    
    # 메세지에 생성한 MIME 문서를 첨부합니다.
    msg.attach(msgPart)
    msg.attach(bookPart)
    
    print ("connect smtp server ... ")
    s = mysmtplib.MySMTP(host,port)
    #s.set_debuglevel(1)
    s.ehlo()
    s.starttls()
    s.ehlo()
    s.login(senderAddr, passwd)    # 로긴을 합니다. 
    s.sendmail(senderAddr , [recipientAddr], msg.as_string())
    s.close()
    
    print ("Mail sending complete!!!")
  
##### run #####
while(loopFlag > 0):
    printMenu()
    menuKey = str(input ('select menu :'))
    launcherFunction(menuKey)
else:
    print ("Thank you! Good Bye")