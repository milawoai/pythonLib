# -*- coding: utf-8 -*-  
import socket
import urllib.request, urllib.error, urllib.parse
import re
import os
import win32gui
import win32com
import win32com.client
import pythoncom
from reportlab.pdfgen.canvas import Canvas  
from reportlab.pdfbase import pdfmetrics  
from reportlab.pdfbase.cidfonts import UnicodeCIDFont  
pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))  
from reportlab.lib.pagesizes import letter, A4  
from reportlab.lib.styles import ParagraphStyle,PropertySet  
from reportlab.platypus import Paragraph  
from reportlab.lib.enums import *  
from reportlab.lib.colors import *  
from reportlab.lib.styles import getSampleStyleSheet  
from reportlab.platypus import Paragraph, SimpleDocTemplate, PageBreak  
from reportlab.pdfbase.ttfonts import TTFont  
socket.setdefaulttimeout(30) 

Format = ["jpg", 'gif', 'png']
#----------- 处理页面上的各种标签 -----------
class HTML_Tool:
    # 用非 贪婪模式 匹配 \t 或者 \n 或者 空格 或者 超链接 或者 图片
    BgnCharToNoneRex = re.compile("(\t|\n| |<a.*?>|<img.*?>)")
    
    # 用非 贪婪模式 匹配 任意<>标签
    EndCharToNoneRex = re.compile("<.*?>")

    # 用非 贪婪模式 匹配 任意<p>标签
    BgnPartRex = re.compile("<p.*?>")
    CharToNewLineRex = re.compile("(<br/>|</p>|<tr>|<div>|</div>|<br>)")
    CharToNextTabRex = re.compile("<td>")

    # 将一些html的符号实体转变为原始符号
    replaceTab = [("&lt;","<"),("&gt;",">"),("&amp;","&"),("&amp;","\""),("&nbsp;"," ")]
    
    def Replace_Char(self,x):
        x = self.BgnCharToNoneRex.sub("",x)
        x = self.BgnPartRex.sub("\n    ",x)
        x = self.CharToNewLineRex.sub("\n",x)
        x = self.CharToNextTabRex.sub("\t",x)
        x = self.EndCharToNoneRex.sub("",x)

        for t in self.replaceTab:  
            x = x.replace(t[0],t[1])  
        return x  
    
class Baidu_Spider:
    # 申明相关的属性
    def __init__(self,url,FilePath = 'F:/ericPython/SPiderAndOther/ed/'):
        self.myUrl = url + '?see_lz=1'
        self.path = FilePath
        self.datas = []
        self.myTool = HTML_Tool()
        self.No = 0
        self.f = None
        self.titleName = None
        os.chdir(FilePath)
        print('已经启动百度贴吧爬虫，咔嚓咔嚓')
  
    # 初始化加载页面并将其转码储存
    def baidu_tieba(self):
        # 读取页面的原始信息并将其从gbk转码
        myPage = urllib.request.urlopen(self.myUrl).read().decode("gbk")
        # 计算楼主发布内容一共有多少页
        endPage = self.page_counter(myPage)
        # 获取该帖的标题
        title = self.find_title(myPage)
        print('文章名称：' + title)
        new_path = os.path.join(self.path, title)
        if not os.path.isdir(new_path):
            os.makedirs(new_path)
        
        # 获取最终的数据
        self.save_data(self.myUrl,title,endPage, new_path)
        return title

    #用来计算一共有多少页
    def page_counter(self,myPage):
        # 匹配 "共有<span class="red">12</span>页" 来获取一共有多少页
        myMatch = re.search(r'class="red">(\d+?)</span>', myPage, re.S)
        if myMatch:  
            endPage = int(myMatch.group(1))
            print('爬虫报告：发现楼主共有%d页的原创内容' % endPage)
        else:
            endPage = 0
            print('爬虫报告：无法计算楼主发布内容有多少页！')
        return endPage

    # 用来寻找该帖的标题
    def find_title(self,myPage):
        # 匹配 <h1 class="core_title_txt" title="">xxxxxxxxxx</h1> 找出标题
        myMatch = re.search(r'<h1.*?>(.*?)</h1>', myPage, re.S)
        title = '暂无标题'
        if myMatch:
            title  = myMatch.group(1)
        else:
            print('爬虫报告：无法加载文章标题！')
        # 文件名不能包含以下字符： \ / ： * ? " < > |
        title = title.replace('\\','').replace('/','').replace(':','').replace('*','').replace('?','').replace('"','').replace('>','').replace('<','').replace('|','')
        return title


    # 用来存储楼主发布的内容
    def save_data(self,url,title,endPage, new_path):
        # 加载页面数据到数组中
        os.chdir(new_path)
        self.get_data(url,endPage, title)
        # 打开本地文件
        print('爬虫报告：文件已下载到本地并打包成txt文件')
        print('请按任意键退出...')
        input();

    # 获取页面源码并将其存储到数组中
    def get_data(self,url,endPage, title):
        
        url = url + '&pn='
        self.f = open(title+'.txt','w+')
        self.titleName = title
        for i in range(1,endPage+1):
            print('爬虫报告：爬虫%d号正在加载中...' % i)
            print(url + str(i))
            try:
                myPage = urllib.request.urlopen(url + str(i)).read()
            except urllib.error.URLError:
                i = i -1
                continue
            except socket.timeout:
                i = i -1
                continue
            # 将myPage中的html代码处理并存储到datas里面
            #self.deal_data(myPage.decode('gbk'))
            self.deal_data(myPage.decode("gbk", 'ignore'))
        self.f.close()
        
    def catch_img1(self, Item):
        myItem = re.findall('<img.*?src=\"(.*?)\"',Item, re.S)
        if myItem == None:
            return None
        else:
            for item in myItem:
                dirset = item.split('.')[-1]
                if len(dirset)>3:
                    killAsk = re.compile("\?.*$")
                    dirset = killAsk.sub("", dirset)
                    
                self.No = self.No+1
                Name = '%(num)05d.%(dir)s' %{'num':self.No,'dir':dirset}
                src = r'<img.*?src="%s".*?>' %item
                Name = '\n'+Name+'\n'
                WSrc = r'<img.*?width="(\d+)".*?src="%s".*?>|<img.*?src="%s".*?width="(\d+)">' %(item, item)
                Width = re.match(WSrc,Item)
                HSrc = r'<img.*?height="(\d+)".*?src="%s".*?>|<img.*?src="%s".*?height="(\d+)">' %(item, item)
                Height = re.match(HSrc,Item)
                if Width != None:
                    print(Width.group(0))
                if Height != None:
                    print(Height.group(0))
                Item = re.sub(src, Name, Item)
            return Item
            
    def catch_img(self, Item):
        myItem = re.findall('(<img.*?src=\".*?\".*?>)',Item, re.S)
        if myItem == None:
            return None
        else:
            for item in myItem:
                #print("************************************************************")
                WSrc = r'<img.*?width="(\d+)".*?>' 
                Width = re.findall(WSrc,item)
                HSrc = r'<img.*?height="(\d+)".*?>'
                Height = re.findall(HSrc,item)
                if Width != None:
                    width = int(Width[0])
                if Height != None:
                   height = int(Height[0])
                if (Width == None and Height == None) or (width*height>=2500):
                    myItem1 =  re.findall('<img.*?src=\"(.*?)\"',item, re.S)
                    if myItem1 == None:
                         return None
                    else:
                        for item1 in myItem1:
                            dirset = item1.split('.')[-1]
                            if len(dirset)>3:
                                killAsk = re.compile("\?.*$")
                                dirset = killAsk.sub("", dirset)
                            self.No = self.No+1
                            Name = '%(num)05d.%(dir)s' %{'num':self.No,'dir':dirset}
                            src = r'<img.*?src="%s".*?>' %item1
                            if not os.path.isfile(Name):
                                print("****%(Name)s****" %{'Name':Name})
                                pf = open(Name,'wb')
                                try:
                                    data = urllib.request.urlopen(item1).read()                                    
                                except urllib.error.URLError:
                                    print(Name)
                                    continue
                                except socket.timeout:
                                    print(Name)
                                    continue
                                else:
                                    pf.write(data)
                                pf.close()
                            Name = '\n'+Name+'\n'
                            Item = re.sub(src, Name, Item)
            return Item

        
    # 将内容从页面代码中抠出来
    def deal_data(self,myPage):
        myItems = re.findall('<cc>.*?<div id="post_content.*?>(.*?)</cc>',myPage,re.S)
       
        for item in myItems:
            item = self.catch_img(item)
            data = self.myTool.Replace_Char(item.replace("\n",""))
            self.datas.append(data+'\n')
        self.f.writelines(self.datas)
        self.datas = []

    def pdfGeneration(self):
        c = canvas.Canvas(self.path+'\\'+self.titleName+'\\'+self.titleName+'.pdf')
        ReadFile = open(self.titleName+'.txt','r')
        for line in ReadFile:
            flag = False
            if line.split('.')[-1].strip() in Format:
                for file in os.listdir(os.getcwd()):
                    if file.strip() == line.strip():
                        flag = True
                        print(file)
                        break
                if flag == False:
                    wrange.InsertAfter(line)
                    wrange = newdoc.Range()
            else:
                c.drawString();
        c.save()
        return
    def wordGeneration(self):
        word = win32com.client.gencache.EnsureDispatch('Word.Application')
        #word = win32com.client.Dispatch('Word.Application')
        word.Visible = True
        newdoc = word.Documents.Add()
        newdoc.PageSetup.RightMargin = 20
        newdoc.PageSetup.LeftMargin = 20
        newdoc.PageSetup.Orientation = win32com.client.constants.wdOrientLandscape
        newdoc.PageSetup.PageWidth = 595
        newdoc.PageSetup.PageHeight = 842
        header_range= newdoc.Sections(1).Headers(win32com.client.constants.wdHeaderFooterPrimary).Range
        header_range.ParagraphFormat.Alignment = win32com.client.constants.wdAlignParagraphCenter
        header_range.Font.Bold = True
        header_range.Font.Size = 10
        header_range.Text = self.titleName
        
        total_column = 1
        total_row = 1
        
        ReadFile = open(self.titleName+'.txt','r')
        wrange = newdoc.Range()
        for line in ReadFile:
            flag = False
            if line.split('.')[-1].strip() in Format:
                for file in os.listdir(os.getcwd()):
                    if file.strip() == line.strip():
                        flag = True
##                        table = newdoc.Tables.Add(wrange,total_row, total_column)
##                        table.Borders.Enable = False
##                        cell_range= table.Cell(1, 1).Range
##                        cell_range.ParagraphFormat.LineSpacingRule =win32com.client.constants.wdLineSpaceSingle
##                        cell_range.ParagraphFormat.SpaceBefore = 0 
##                        cell_range.ParagraphFormat.SpaceAfter = 3
                        wrange = newdoc.Range(newdoc.Content.End -1 ,newdoc.Content.End)
                        wrange.InlineShapes.AddPicture(os.path.join(os.path.abspath("."), file))
                        wrange = newdoc.Range()
                        print(file)
                        #wrange.InsertAfter(line)
                        #wrange = newdoc.Range()
                        break
                if flag == False:
                    wrange.InsertAfter(line)
                    wrange = newdoc.Range()
            else:
                wrange.InsertAfter(line)
                wrange = newdoc.Range()
        newdoc.SaveAs(self.path+'\\'+self.titleName+'\\'+self.titleName+'.docx')
        newdoc.Close()
        word.Quit()
        

if __name__ == "__main__":
    print('请输入贴吧的地址:')
    bdurl = str(input())
    mySpider = Baidu_Spider(bdurl)
    name = mySpider.baidu_tieba()
    mySpider.wordGeneration()
