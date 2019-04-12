"""
Created on Apr 11, 2019

@author: HanqingTony

Requires pywin32

原理说明：
一般来讲，使用python操作win剪贴板，可以使用win32clipboard模块，使用其中的SetClipboardData函数进行写入。当SetClipboardData函数附带不同参数时可将剪贴板内容识别成不同格式，但是默认参数中没有对应html格式的参数，如果直接将html写入，将被识别为字符串。参考官方文档后，发现可以使用同模块中的RegisterClipboardFormat函数进行注册，取得对应html格式的参数int:CF_HTML，使用带此参数的SetClipboardData函数写入精心构造的二进制字符串，剪贴板可以正确将内容识别为html。

关于剪贴板中html格式，windows做了严格的规定，具体的构建方式可以参考windows官方文档：

其中有几个坑
- 编码问题：根据utf-8标准，一个中文字符占用3个ascii大小，一个半角空格占用1一个ascii，回车按照win系的管理/r/n, 占用2个字符。数数的时候一定要小心。之前老外开发的模块中文会是乱码，就是没有考虑中文占ascii码大小问题。
- 各参数的格式：一定要用0补满8位。windows官方文档的范例并不是二进制格式而是转码后的格式，不仅回车的格式没有体现，各位置参数也并没有补满八位。因此直接将范例写入，格式是错的。

相关文档：
关于剪贴板中的html内容格式：https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa767917(v=vs.85
关于pywin32剪贴板格式注册：http://timgolden.me.uk/pywin32-docs/win32clipboard__RegisterClipboardFormat_meth.html
"""

#import re
import time
#import random
import win32clipboard

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
        #用来存储html代码片段的属性
        self.fragment = ''
        #self.selection = None
        #self.source = None
        #self.htmlClipboardVersion = None
        '''
        注册"HTML Format"格式，返回一个整数数字，该数字为剪贴板将内容识别为html的关键参数。
        '''
        try:
            self.CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
            report("注册HTML格式成功！CF_HTML其值为", self.CF_HTML)
        except:
            report("self.__init__:注册HTML格式失败！")

    
    def putIn(self, fragment):
        '''
        直接将html片段放入剪贴板的函数，逻辑为
        1. 读入目标HTML代码
        2. 计算目标代码长度
        3. 生成剪贴板内容
        4. 写入
        四个部分
        '''
        self.readFragment(fragment)
        self.calculateLength()
        self.constructContent()
        self.put()

    def readFragment(self, fragment):
        '''
        对传入的html字符串进行清洗，并保存在self.fragment中
        '''
        #检查fragment格式
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


if __name__ == "__main__":
    #test_dump()
    #test_format()
    testHTML = '''<section data-role="outer" label="Powered by 135editor.com" style="font-size:16px;"><section class="_135editor" data-tools="135编辑器" data-id="94817" style="border: 0px none; box-sizing: border-box;"><section style="width: 100%;" data-width="100%"><section style="border-left: 6px solid #ffb497;padding-left:6px;box-sizing: border-box;"><section class="135brush" data-brushtype="text" style="background: rgb(255, 232, 223) none repeat scroll 0% 0%; display: inline-block; padding: 2px 12px; letter-spacing: 1.5px; color: rgb(51, 51, 51); font-size: 16px; box-sizing: border-box;">欣赏和珍惜</section></section><section style="padding-left: 5px; box-sizing: border-box;"><section style="border-color: currentcolor currentcolor rgb(208, 208, 208) rgb(208, 208, 208); border-style: none none solid solid; border-width: medium medium 1px 1px; border-image: none 100% / 1 / 0 stretch; box-sizing: border-box;"><section data-autoskip="1" class="135brush" style="padding: 1em; font-size: 14px; letter-spacing: 1.5px; line-height: 1.75em; color: rgb(51, 51, 51); text-align: justify; box-sizing: border-box;"><p>即使在清贫的岁月，也不能失去对幸福美好的向往，那些摆脱平庸的梦总能编制我们简单的生活，为我们简单的时光点缀希望。不能说我们总要多热爱生活，但总要有一颗懂得欣赏和珍惜的心。</p></section></section><section style="width: 1.2em;background: #ffb497;height:6px;float: right;"></section><section style="clear: both;"></section></section></section></section><section class="_135editor" data-role="paragraph" style="border: 0px none; box-sizing: border-box;"><p><br></p></section></section>'''
    test_dumpHTML()