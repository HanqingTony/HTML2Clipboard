# HTML2Clipboard v0.1
a python3 module to put HTML string into win clipboard with format  

将Html文本以可识别的格式放入win剪贴板中。剪贴板内容可以直接带格式粘贴在微信公众号后台等处。
## 使用方法：
```python3
import HTML2Clipboard

html2cb = HTML2Clipboard.MainClass()
html2cb.putIn('''
<div style = "color:blue;">
    <p>testHTML</p>
</div>
''')
```
## 原理说明:    
一般来讲，使用python操作win剪贴板，可以使用win32clipboard模块，使用其中的SetClipboardData函数进行写入。当SetClipboardData函数附带不同参数时可将剪贴板内容识别成不同格式，但是默认参数中没有对应html格式的参数，如果直接将html写入，将被识别为字符串。参考官方文档后，发现可以使用同模块中的RegisterClipboardFormat函数进行注册，取得对应html格式的参数int:CF_HTML，使用带此参数的SetClipboardData函数写入精心构造的二进制字符串，剪贴板可以正确将内容识别为html。

关于剪贴板中html二进制串的构造，windows做了严格的规定，具体的构造格式请参考下附的windows官方文档。  

## 几个坑:    
- 编码问题：根据utf-8标准，一个中文字符占用3个ascii大小，一个半角空格占用1一个ascii，回车按照win系的惯例/r/n, 占用2个字符。数数的时候一定要小心。之前老外开发的模块中文会是乱码，就是没有考虑中文占ascii码大小问题。

- 各参数的格式：一定要用0补满8位。windows官方文档的范例并不是二进制格式而是转码后的格式，不仅回车的格式没有体现，各位置参数也并没有补满八位。因此直接将范例写入，格式是错的。

## 相关文档:    
关于剪贴板中的html内容格式：https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa767917(v=vs.85)  
关于pywin32剪贴板格式注册：http://timgolden.me.uk/pywin32-docs/win32clipboard__RegisterClipboardFormat_meth.html
