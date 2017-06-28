VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XmlScraping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Scraping v0.1.1 Alpha
' (c) Victor Zevallos - https://github.com/vba-dev/vba-scraping
'
' Library used: Microsoft XML, v6.0, Microsoft HTML Object Library
'
' @class XmlScraping
' @author victorzevallos@protonmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

Private Doc As New MSHTML.HTMLDocument

Private Element As MSHTML.IHTMLElement

Private Elements As MSHTML.IHTMLElementCollection

Public Sub GotoPage(Url As String)
    Dim XMLPage As New MSXML2.XMLHTTP60

    XMLPage.Open "GET", Url, False
    XMLPage.send

    Doc.body.innerHTML = XMLPage.responseText
End Sub

Public Function Css(Selector As String) As XmlScraping
    Set Element = Doc.querySelector(Selector)
    
    Set Css = Me
End Function

Public Function Text() As String
    Text = Element.innerText
End Function

Public Function Html() As String
    Html = Element.innerHTML
End Function

