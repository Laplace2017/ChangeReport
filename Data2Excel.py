#coding=utf-8
import xlwt
import time

class Issue2Excel:
######################################################
#函数名：issues2excel
#输入参数：缺陷结构体数组
#输出：生成excel文件
#作者：zc
#时间：2018年3月27日11:09:37
#####################################################
    @staticmethod
    def issues2excel(issues, filepathname):
        #边框设置
        borders = xlwt.Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        borders.bottom_colour = 0x3A

        # “行、列”等描述性格式设置
        style = xlwt.XFStyle()
        font = xlwt.Font()
        pattern = xlwt.Pattern()
        style.pattern = pattern
        pattern.pattern = 1
        pattern.pattern_fore_colour = 44
        font.name = 'Times New Roman'
        font.height = 250
        font.bold = True
        style.font = font
        style.borders = borders

        # “静态分析报告单”格式设置
        style0 = xlwt.XFStyle()
        font0 = xlwt.Font()
        pattern0 = xlwt.Pattern()
        style0.pattern = pattern0
        pattern0.pattern = 1
        pattern0.pattern_fore_colour = 30
        font0.name = 'Arial'
        font0.colour_index = 1
        font0.height = 400
        font0.bold = True
        style0.font = font0
        al = xlwt.Alignment()
        al.horz = xlwt.Alignment.HORZ_CENTER
        al.vert = xlwt.Alignment.VERT_CENTER
        style0.alignment = al
        style0.borders = borders

        # 最下面的缺陷输入格式设置
        style2 = xlwt.XFStyle()
        al2 = xlwt.Alignment()
        al2.horz = xlwt.Alignment.HORZ_LEFT
        al2.vert = xlwt.Alignment.VERT_CENTER
        style2.alignment = al2
        style2.borders = borders

        # 日期和编制时间格式设置
        style1 = xlwt.XFStyle()
        font1 = xlwt.Font()
        pattern1 = xlwt.Pattern()
        style1.pattern = pattern1
        pattern1.pattern = 1
        pattern1.pattern_fore_colour = 30
        font1.name = 'Times New Roman'
        font1.colour_index = 1
        font1.bold = True
        style1.font = font1
        style1.borders = borders

        # 自动换行格式设置
        wrap_style = xlwt.easyxf('align:vert center, horiz left, wrap on;'
                                 'borders:left 1,top 1, bottom 1,right 1;'
                                 )
        # 行高格式设置
        tall_style = xlwt.easyxf('font:height 720;')

        #缺陷类型翻译字典
        Dict = {u"Android缺陷": ['ANDROID.RLK.MEDIAPLAYER', 'ANDROID.RLK.MEDIARECORDER', 'ANDROID.RLK.SQLCON',
                               'ANDROID.RLK.SQLOBJ', 'ANDROID.UF.BITMAP', 'ANDROID.UF.CAMERA', 'ANDROID.UF.MEDIAPLAYER',
                               'ANDROID.UF.MEDIARECORDER'],
                u"使用已释放的资源": ['UF.IMAGEIO', 'UF.IN', 'UF.JNDI', 'UF.MAIL', 'UF.MICRO', 'UF.NIO', 'UF.OUT', 'UF.SOCK',
                              'UF.SQLCON', 'UF.SQOBJ', 'UF.ZIP'],
                u"信息泄露": ['SV.IL.DEV', 'SV.IL.FILE'],
                u"弱加密": ['SV.PASSWD.HC', 'SV.PASSWD.HC.EMPTY', 'SV.PASSWD.PLAIN'],
                u"忽略返回值": ['RI.IGNOREDCALL', 'RI.IGNOREDNEW', 'RR.IGNORED'],
                u"拒绝服务": ['SV.INT_OVF'],
                u"数据注入": ['SV.DATA.DB', 'SV.LDAP', 'SV.SQL'],
                u"未经验证的用户输入": ['SV.EMAIL', 'SV.HTTP_SPLIT', 'SV.XPATH'],
                u"潜在的运行时缺陷": ['JD.UNMOD', 'NPE.COND', 'NPE.CONST', 'NPE.RET', 'NPE.RET.UTIL'],
                u"线程和同步缺陷": ['JD.LOCK'],
                u"资源泄露": ['RLK.AWT', 'RLK.HIBERNATE', 'RLK.IMAGEIO', 'RLK.IN', 'RLK.JNDI', 'RLK.MAIL', 'RLK.MICRO',
                          'RLK.NIO', 'RLK.OUT', 'RLK.SOCK', 'RLK.SQLCON', 'RLK.SQLOBJ', 'RLK.SWT', 'RLK.ZIP'],
                u"跨站点脚本攻击": ['SV.XSS.DB', 'SV.XSS.REF'],
                u"进程及路径注入": ['SV.EXEC', 'SV.EXEC.DIR', 'SV.EXEC.ENV'], }

        #根据输入的缺陷数组写入excel
        issuewb = xlwt.Workbook(encoding='gbk')
        issuest = issuewb.add_sheet('Sheet1')

        issuest.write_merge(0, 0, 0, 9, u"静态分析报告单" , style0)
        issuest.write_merge(1, 1, 0, 9, u"编制：可靠性与检测技术中心", style1)
        issuest.write_merge(2, 2, 0, 9, u"日期：" + time.strftime("%Y") + u"年" + time.strftime("%m") + u"月" + time.strftime(
            "%d") + u"日" + time.strftime("%H") + u"时" + time.strftime("%M") + u"分", style1)
        #抬头
        tag = [u"序号", u"路径", u"行", u"缺陷类型", u"检查准则", u"缺陷描述", u"缺陷追踪", u"status(缺陷状态)", u"state(陈述)", u"严重级别"]
        for j in range(10):
            issuest.write(3, j, tag[j], style)

        # 列表宽度
        for l in range(10):
            # 0序号
            if l == 0:
                issuest.col(l).width = 256 * 7
            # 1路径和5描述
            elif l == 1 or l == 5:
                issuest.col(l).width = 256 * 40
            # 7status、8state、9等级
            elif l == 7 or l == 8 or l == 9:
                issuest.col(l).width = 256 * 12
            # 4检查准则
            elif l == 4 or l == 3:
                issuest.col(l).width = 256 * 15
            # 2行
            elif l == 2:
                issuest.col(l).width = 256 * 7
            # 6追踪
            elif l == 6:
                issuest.col(l).width = 256 * 50

        #填写缺陷
        issuenum = len(issues)
        for i in range(0, issuenum):
            #序号
            issuest.write(i + 4, 0, issues[i].sort, wrap_style)
            #路径
            issuest.write(i + 4, 1, issues[i].codepath, wrap_style)
            #行
            issuest.write(i + 4, 2, issues[i].line, wrap_style)
            #缺陷类型（根据缺陷准则和字典Dict获取）
            typeitems = Dict.items()
            exit_flag = False
            for key, norms in typeitems:
                for norm in norms:
                    if norm == issues[i].norms:
                        exit_flag = True
                        issues[i].issuetype = key
                        issuest.write(i + 4, 3, issues[i].issuetype, wrap_style)
                        break
                if exit_flag:
                    break

            #缺陷准则
            issuest.write(i + 4, 4, issues[i].norms, wrap_style)
            #缺陷描述
            issuest.write(i + 4, 5, issues[i].description, wrap_style)
            #缺陷追踪
            issuest.write(i + 4, 6, issues[i].traceability, wrap_style)
            #缺陷状态
            issuest.write(i + 4, 7, issues[i].status, wrap_style)
            #陈述
            issuest.write(i + 4, 8, issues[i].state, wrap_style)
            #严重级别
            issuest.write(i + 4, 9, issues[i].level, wrap_style)

        #保存文件
        issuewb.save(filepathname)