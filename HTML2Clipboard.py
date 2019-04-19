"""
Created on Apr 11, 2019

@author: HanqingTony(81810395@qq.com)

Requires pywin32

"""
import win32clipboard, win32con

#定义report函数用以调试，此处简单的定义为print，可以自己编写report函数实现更复杂的记录
report = print

class MainClass:
    '''
    主类
    '''
    def __init__(self):
        #前缀模板 长度97
        self.prefix = "Version:1.0\r\nStartHTML:00000097\r\nEndHTML:%8d\r\nStartFragment:00000191\r\nEndFragment:%8d\r\n<!DOCTYPE html><HTML><HEAD><TITLE>HTML To Clipboard!</TITLE></HEAD><BODY><!--StartFragment -->"
        #后缀模板 长度33 EndHTML = EndFragment + 33
        self.suffix = "<!--EndFragment --></BODY></HTML>"
        #html代码
        self.fragment = ''
        #self.selection = None
        #self.source = None
        #self.htmlClipboardVersion = None
        '''
        下面为注册"HTML Format"格式的过程。注册函数返回一个整数数字，该数字为剪贴板将内容识别为html的关键参数。在本人的运行环境中该数字固定为49388
        '''
        try:
            self.CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
            report("self.__init__:注册HTML格式成功！CF_HTML值为", self.CF_HTML)
        except:
            report("self.__init__:注册HTML格式失败！")

    
    def putIn(self, fragment):
        '''
        直接将html片段放入剪贴板的函数，输入html代码片段，运行逻辑为
        1. 读入目标HTML代码
        2. 计算目标代码长度
        3. 生成剪贴板内容
        4. 写入
        四个部分
        无返回值
        '''
        self.readFragment(fragment)
        self.calculateLength()
        self.constructContent()
        self.put()

    def readFragment(self, fragment):
        '''
        对传入的html字符串进行清洗，并保存在self.fragment中
        '''
        #检查fragment格式，未实现
        pass
        #将fragment存为self.fragment
        self.fragment = fragment

    def calculateLength(self):
        '''
        计算html片段的长度，先转换为二进制，再使用len函数计算
        '''
        self.fragmentLength = len(self.fragment.encode('utf-8'))
        report(self.fragmentLength)

    def constructContent(self):
        '''
        构造写入剪贴板内容
        1. 首先计算片段结束位置fragmentEnd，再据此计算出html结束位置也即整个Content的长度htmlEnd
        并代入之前的self.prefix中
        2. 然后将prefix，fragment，suffix合并，生成字符串contentText
        3. 最后生成二进制串contentBin
        '''
        fragmentEnd = 191 + self.fragmentLength
        htmlEnd = fragmentEnd + 33
        report("htmlend = ", htmlEnd, "fragmentEnd=", fragmentEnd)
        self.contentText = self.prefix%(htmlEnd, fragmentEnd) + self.fragment + self.suffix
        self.contentBin = self.contentText.encode('utf-8')
        report('self.constructContent:生成成功')
    
    def put(self):
        """
        将二进制数据放入剪贴板分三步
        1. 打开剪贴板
        2. 将数据放进去
        3. 关闭剪贴板
        """
        try:
            win32clipboard.OpenClipboard(0)
            win32clipboard.EmptyClipboard()
            #此步为关键，以html格式将二进制串写入剪贴板使用SetClipboardData函数，函数的两个参数一为之前注册的HTML format代码，一为src二进制串
            report(self.contentBin)
            win32clipboard.SetClipboardData(self.CF_HTML, self.contentBin)
        except:
            report('self.put():写入失败')
        finally:
            win32clipboard.CloseClipboard()
            report('self.put():写入成功')

    def dumpBin(self):
        '''
        dump出剪贴板二进制内容,返回二进制串
        '''
        src = None
        try:
            win32clipboard.OpenClipboard(0)
            src = win32clipboard.GetClipboardData(self.CF_HTML)
            #report(src)
        except:
            report('dump失败')
        finally:
            win32clipboard.CloseClipboard()
            self.contentBin = src
            return(self.contentBin)
            

    def dumpHTML(self):
        '''
        dump出字符串格式的剪贴板内容，并保存至self.fragment中
        '''
        try:
            self.fragment = self.dumpBin().decode('utf-8')
            #report(self.fragment)
            self.fragment = self.fragment.split('<!--EndFragment-->')[0].split('<!--StartFragment-->')[-1]
            #report(self.fragment)
            return(self.fragment)
            report('self.dumpHTML:dump成功!')

        except:
            report('self.dumpHTML:dump失败!')
#==============================
#工具函数
#==============================
def dumpText():
    '''
    dump剪贴板中的文字
    '''
    text = None
    try:
        win32clipboard.OpenClipboard()
        text = win32clipboard.GetClipboardData()
        #report(src)
    except:
        report('dumpText:dump失败')
    finally:
        win32clipboard.CloseClipboard()
        return(text)

def GetAvailableFormats():
    """
    返回可以支持的剪贴板格式
    """
    formats = []
    try:
        win32clipboard.OpenClipboard(0)
        cf = win32clipboard.EnumClipboardFormats(0)
        while (cf != 0):
            formats.append(cf)
            cf = win32clipboard.EnumClipboardFormats(cf)
    finally:
        win32clipboard.CloseClipboard()
    return(formats)

#=============================
# 以下为测试部分
# =========================            
def test_put(testHTML):
    a = MainClass()
    a.fragment = testHTML
    a.put()

def test_dump():
    a = MainClass()
    c = a.dumpOutBin()
    report(c.decode('utf-8'))

def test_putIn(html):
    a = MainClass()
    a.putIn(html)

def test_dumpHTML():
    a = MainClass()
    a.dumpHTML()
    print('=============================\n', a.fragment)

def test_getav():
    a = MainClass()
    print(a.GetAvailableFormats())

def test_dumpText():
    print(dumpText())

if __name__ == "__main__":
    #test_dump()
    #test_format()
    test_dumpText()