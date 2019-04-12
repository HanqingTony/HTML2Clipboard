"""
Created on Sep 24, 2013

@author: jordans

Requires pywin32


original: http://code.activestate.com/recipes/474121/
    # HtmlClipboard
    # An interface to the "HTML Format" clipboard data format
    
    __author__ = "Phillip Piper (jppx1[at]bigfoot.com)"
    __date__ = "2006-02-21"
    __version__ = "0.1"

"""

import re
import time
import random
import win32clipboard

#---------------------------------------------------------------------------
#  Convenience functions to do the most common operation

def HasHtml():
    """
    Return True if there is a Html fragment in the clipboard..
    判断剪切板中是否有html代码，如果有 返回真
    """
    cb = HtmlClipboard()
    return cb.HasHtmlFormat()


def GetHtml():
    """
    Return the Html fragment from the clipboard or None if there is no Html in the clipboard.
    从剪切板中返回html片段，如果没有，返回None
    """
    cb = HtmlClipboard()
    if cb.HasHtmlFormat():
        return cb.GetFragment()
    else:
        return None


def PutHtml(fragment):
    """
    Put the given fragment into the clipboard.
    Convenience function to do the most common operation
    把html片段发送到剪贴板
    """
    cb = HtmlClipboard()
    cb.PutFragment(fragment)

#---------------------------------------------------------------------------

class HtmlClipboard:
    
    CF_HTML = None

    MARKER_BLOCK_OUTPUT = \
        "Version:1.0\r\n" \
        "StartHTML:%09d\r\n" \
        "EndHTML:%09d\r\n" \
        "StartFragment:%09d\r\n" \
        "EndFragment:%09d\r\n" \
        "StartSelection:%09d\r\n" \
        "EndSelection:%09d\r\n" \
        "SourceURL:%s\r\n"

    MARKER_BLOCK_EX = \
        "Version:(\S+)\s+" \
        "StartHTML:(\d+)\s+" \
        "EndHTML:(\d+)\s+" \
        "StartFragment:(\d+)\s+" \
        "EndFragment:(\d+)\s+" \
        "StartSelection:(\d+)\s+" \
        "EndSelection:(\d+)\s+" \
        "SourceURL:(\S+)"
    MARKER_BLOCK_EX_RE = re.compile(MARKER_BLOCK_EX)

    MARKER_BLOCK = \
        "Version:(\S+)\s+" \
        "StartHTML:(\d+)\s+" \
        "EndHTML:(\d+)\s+" \
        "StartFragment:(\d+)\s+" \
        "EndFragment:(\d+)\s+" \
           "SourceURL:(\S+)"
    MARKER_BLOCK_RE = re.compile(MARKER_BLOCK)

    DEFAULT_HTML_BODY = \
        "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">" \
        "<HTML><HEAD></HEAD><BODY><!--StartFragment-->%s<!--EndFragment--></BODY></HTML>"

    def __init__(self):
        self.html = None
        self.fragment = None
        self.selection = None
        self.source = None
        self.htmlClipboardVersion = None


    def GetCfHtml(self):
        """
        Return the FORMATID of the HTML format
        """
        if self.CF_HTML is None:
            self.CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")

        return self.CF_HTML


    def GetAvailableFormats(self):
        """
        Return a possibly empty list of formats available on the clipboard
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

        return formats


    def HasHtmlFormat(self):
        """
        Return a boolean indicating if the clipboard has data in HTML format
        """
        return (self.GetCfHtml() in self.GetAvailableFormats())


    def GetFromClipboard(self):
        """
        Read and decode the HTML from the clipboard
        """

        # implement fix from: http://teachthe.net/?p=1137

        cbOpened = False
        while not cbOpened:
            try:
                win32clipboard.OpenClipboard(0)
                src = win32clipboard.GetClipboardData(self.GetCfHtml())
                src = src.decode("utf-8")
                #print(src)
                self.DecodeClipboardSource(src)
                cbOpened = True
                win32clipboard.CloseClipboard()
            except Exception as err:
                # If access is denied, that means that the clipboard is in use.
                # Keep trying until it's available.
                if err.winerror == 5:  # Access Denied
                    pass
                    # wait on clipboard because something else has it. we're waiting a
                    # random amount of time before we try again so we don't collide again
                    time.sleep( random.random()/50 )
                elif err.winerror == 1418:  # doesn't have board open
                    pass
                elif err.winerror == 0:  # open failure
                    pass
                else:
                    print( 'ERROR in Clipboard section of readcomments: %s' % err)

                    pass        
    def DecodeClipboardSource(self, src):
        """
        Decode the given string to figure out the details of the HTML that's on the string
        """
        # Try the extended format first (which has an explicit selection)
        matches = self.MARKER_BLOCK_EX_RE.match(src)
        if matches:
            self.prefix = matches.group(0)
            self.htmlClipboardVersion = matches.group(1)
            self.html = src[int(matches.group(2)):int(matches.group(3))]
            self.fragment = src[int(matches.group(4)):int(matches.group(5))]
            self.selection = src[int(matches.group(6)):int(matches.group(7))]
            self.source = matches.group(8)
        else:
            # Failing that, try the version without a selection
            matches = self.MARKER_BLOCK_RE.match(src)
            if matches:
                self.prefix = matches.group(0)
                self.htmlClipboardVersion = matches.group(1)
                self.html = src[int(matches.group(2)):int(matches.group(3))]
                self.fragment = src[int(matches.group(4)):int(matches.group(5))]
                self.source = matches.group(6)
                self.selection = self.fragment


    def GetHtml(self, refresh=False):
        """
        Return the entire Html document
        """
        if not self.html or refresh:
            self.GetFromClipboard()
        return self.html


    def GetFragment(self, refresh=False):
        """
        Return the Html fragment. A fragment is well-formated HTML enclosing the selected text
        """
        if not self.fragment or refresh:
            self.GetFromClipboard()
        return self.fragment


    def GetSelection(self, refresh=False):
        """
        Return the part of the HTML that was selected. It might not be well-formed.
        """
        if not self.selection or refresh:
            self.GetFromClipboard()
        return self.selection


    def GetSource(self, refresh=False):
        """
        Return the URL of the source of this HTML
        """
        if not self.selection or refresh:
            self.GetFromClipboard()
        return self.source


    def PutFragment(self, fragment, selection=None, html=None, source=None):
        """
        Put the given well-formed fragment of Html into the clipboard.

        selection, if given, must be a literal string within fragment.
        html, if given, must be a well-formed Html document that textually
        contains fragment and its required markers.
        """
        if selection is None:
            selection = fragment
        if html is None:
            html = self.DEFAULT_HTML_BODY % fragment
        if source is None:
            source = "file://HtmlClipboard.py"

        fragmentStart = html.index(fragment)
        fragmentEnd = fragmentStart + len(fragment)
        selectionStart = html.index(selection)
        selectionEnd = selectionStart + len(selection)
        self.PutToClipboard(html, fragmentStart, fragmentEnd, selectionStart, selectionEnd, source)


    def PutToClipboard(self, html, fragmentStart, fragmentEnd, selectionStart, selectionEnd, source="None"):
        """
        Replace the Clipboard contents with the given html information.
        """

        try:
            win32clipboard.OpenClipboard(0)
            win32clipboard.EmptyClipboard()
            src = self.EncodeClipboardSource(html, fragmentStart, fragmentEnd, selectionStart, selectionEnd, source)
            src = src.encode("UTF-8")
            #print(src)
            #此步为关键，以html格式将二进制串写入剪贴板使用SetClipboardData函数，函数的两个参数一为之前注册的HTML format代码，一为src二进制串
            win32clipboard.SetClipboardData(self.GetCfHtml(), src)
        finally:
            win32clipboard.CloseClipboard()


    def EncodeClipboardSource(self, html, fragmentStart, fragmentEnd, selectionStart, selectionEnd, source):
        """
        Join all our bits of information into a string formatted as per the HTML format specs.
        """
        # How long is the prefix going to be?
        dummyPrefix = self.MARKER_BLOCK_OUTPUT % (0, 0, 0, 0, 0, 0, source)
        lenPrefix = len(dummyPrefix)

        prefix = self.MARKER_BLOCK_OUTPUT % (lenPrefix, len(html)+lenPrefix,
                        fragmentStart+lenPrefix, fragmentEnd+lenPrefix,
                        selectionStart+lenPrefix, selectionEnd+lenPrefix,
                        source)
        return (prefix + html)


def DumpHtml():

    cb = HtmlClipboard()
    print("GetAvailableFormats()=%s" % str(cb.GetAvailableFormats()))
    print("HasHtmlFormat()=%s" % str(cb.HasHtmlFormat()))
    if cb.HasHtmlFormat():
        cb.GetFromClipboard()
        print("prefix=>>>%s<<<END" % cb.prefix)
        print("htmlClipboardVersion=>>>%s<<<END" % cb.htmlClipboardVersion)
        print("GetSelection()=>>>%s<<<END" % cb.GetSelection())
        print("GetFragment()=>>>%s<<<END" % cb.GetFragment())
        print("GetHtml()=>>>%s<<<END" % cb.GetHtml())
        print("GetSource()=>>>%s<<<END" % cb.GetSource())



if __name__ == '__main__':

    #data = "<p>Writing to the clipboard is <strong>easy</strong> with this code.</p>"
    data = '''<section data-role="outer" label="Powered by 135editor.com" style="font-size:16px;"><section class="_135editor" data-tools="135编辑器" data-id="94785" style="border: 0px none; box-sizing: border-box;"><section style="padding: 10px 10px 10px 20px; box-sizing: border-box;"><section style="margin-left: -15px; margin-bottom: -24px; overflow: hidden; padding: 0px; display: flex; align-items: center; box-sizing: border-box;"><section style="width: 2.2em; height: 2.2em; background: rgb(228, 244, 253) none repeat scroll 0% 0%; border-radius: 100%; text-align: center; line-height: 2.2em; font-size: 18px; box-sizing: border-box;"><br></section><section style="margin-left: -10px; display: inline-block; font-size: 1em; font-family: inherit; font-weight: inherit; text-align: inherit; text-decoration: inherit; color: rgb(254, 254, 254); border-color: rgb(239, 112, 96); background-color: transparent; box-sizing: border-box;"><section class="135brush" data-bcless="darken" data-brushtype="text" style="letter-spacing: 1.5px; border-left-width: 8px; border-left-style: solid; border-color:#e4f4fd; display: inline-block; line-height: 1.4em; padding: 5px 10px; height: 32px; vertical-align: top; font-size: 16px; font-family: inherit;float: left; color: inherit; box-sizing: border-box !important; background:#e4f4fd;color: #333;">“问君哪得怕如此,为有源头射线来”</section><section style="width: 0.5em; display: inline-block; height: 32px; vertical-align: top; border-bottom-width: 1em; border-bottom-style: solid; border-bottom-color:#e4f4fd; border-top-width: 1em; border-top-style: solid; border-top-color:#e4f4fd; font-size: 16px; border-left-color: #e4f4fd; color: inherit; box-sizing: border-box !important; border-right-width: 1em !important; border-right-style: solid !important; border-right-color: transparent !important;"></section></section></section><section style="border-radius:1px ;border:1px #e6f5fd solid;box-sizing: border-box;background: #f5fafd;border-radius:6px ;"><section style="border-radius:1px ;border:1px #e6f5fd solid;margin:5px -5px -5px 5px;border-radius:6px ;box-sizing: border-box;background-color:#fefefe"><section data-autoskip="1" class="135brush" style="font-size: 14px; text-align: justify; letter-spacing: 1.5px; line-height: 1.75em; color: rgb(63, 63, 63); padding: 1.5em 1em 1em; box-sizing: border-box;"><p><br></p><p style="text-indent: 2em;">自然界中的一切物体，一般都会以电磁波的形式时刻不停地向外传送热量，这种传送能量的方式称为辐射。简单说来就是从某种物质中发射出来的波或粒子（热辐射、核辐射等）。其实我们天天和辐射打交道，只是我们自己并不一定会意识到，太阳光、紫外线、热、声等这些都是辐射。但当人们谈论辐射时，往往想到的却是相关放射性的这一类辐射。</p></section></section></section></section></section><section class="_135editor" data-role="paragraph" style="border: 0px none; box-sizing: border-box;"><p><br></p></section></section>
    '''
    def test_SimpleGetPutHtml(data):
        PutHtml(data)
        if GetHtml() == data:
            print("passed")
        else:
            print("failed")
    
    def test_get(data):
        PutHtml(data)
        print(GetHtml())


    def test_Put(fragment):
        """
        Put the given fragment into the clipboard.
        Convenience function to do the most common operation
        把html片段发送到剪贴板
        """
    print(data)
    PutHtml(data)
    
    #test_SimpleGetPutHtml()
    #cont = cont.encode('utf-8')
    ##PutHtml(cont)
    #DumpHtml()
    