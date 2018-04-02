#coding = utf-8

from chardet import detect

#缺陷结构体
class Issue(object):
    def __init__(self):
        #序号
        self.sort=''
        #代码路径
        self.codepath=''
        #代码行
        self.line=''
        #缺陷类型
        self.issuetype=''
        #缺陷准则
        self.norms=''
        #缺陷描述
        self.description=''
        #缺陷追踪
        self.traceability=''
        self.status=''
        self.state=''
        #缺陷等级
        self.level=''



######################################################
#函数名：processtxtdata
#输入参数：txt格式报告路径
#输出：缺陷结构体数组
#作者：zc
#时间：2018年3月26日15:25:34
#####################################################


    def processtxtdata(txtpath):
        #以二进制形式读取txt报告
        txtreport = open(txtpath, 'rb')
        txtdata = txtreport.read()
        txtreport.close()
        #判断txt报告编码格式
        txtdatainfo = detect(txtdata)
        txtdataencoding = txtdatainfo['encoding']
        #以明文方式读取txt报告
        filedata = open(txtpath, 'rU', encoding=txtdataencoding)
        txtdata = filedata.read()
        filedata.close()
        #缺陷分隔符号
        stringkey = '---------------------------------------------------------------------------'
        #缺陷数组
        issues = []
        #分割缺陷
        reissues = txtdata.split(stringkey)
        #print(reissues[1])
        reissuesnum = len(reissues)
        for i in range(1, reissuesnum):
            issueinfo = reissues[i].split('\n')
            #print(issueinfo)
            issueinfonum = len(issueinfo)
            issue = Issue()
            issue.sort = issueinfo[1].split(' ')[0]
            # 当编码格式为“ACHI”时，为Windows下的txt报告，为“UTF-8”时，为Linux下的txt报告
            #print(txtdataencoding)
            #print(txtdataencoding.find('asc'))
            if txtdataencoding.find('ascii') != -1:
                issue.codepath = issueinfo[1].split(':')[1]
                issue.line = issueinfo[1].split(' ')[2].split(':')[2]
            else:
                issue.line = issueinfo[1].split(' ')[2].split(':')[1]
                issue.codepath = issueinfo[1].split(' ')[2].split(':')[0]

            issue.norms = issueinfo[1].split(' ')[3]
            issue.level = issueinfo[1].split(' ')[4]
            issue.state = issueinfo[1].split(' ')[5]
            issue.description = issueinfo[2]
            if i == reissuesnum - 1:
                for j in range(3, issueinfonum - 5):
                    issue.traceability = issue.traceability + issueinfo[j] + '\n'
                    #print(issue.traceability)
                issue.status=issueinfo[issueinfonum - 5].split(' ')[2]
            else:
                for j in range(3, issueinfonum - 3):
                    issue.traceability = issue.traceability + issueinfo[j] + '\n'

                issue.status = issueinfo[issueinfonum - 3].split(' ')[2]
            #print(issue.traceability)
            #print(issueinfo[8])
            #print('1+1:' + issueinfo[issueinfonum - 3])

            #print(issue.status)
            issues.append(issue)

        return issues


