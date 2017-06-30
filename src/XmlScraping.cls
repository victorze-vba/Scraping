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

Private doc As New MSHTML.HTMLDocument

Private element As MSHTML.IHTMLElement

Private elements As MSHTML.IHTMLElementCollection

Private children As MSHTML.IHTMLDOMChildrenCollection

Public Sub gotoPage(url As String)
    Dim XMLPage As New MSXML2.XMLHTTP60

    XMLPage.Open "GET", url, False
    XMLPage.send

    doc.body.innerHTML = XMLPage.responseText
End Sub

Public Function css(selector As String) As XmlScraping
    Set children = doc.querySelectorAll(selector)
    
    Set css = Me
End Function

Public Function ind(i As Integer) As XmlScraping
    Set element = children(i)
    
    Set ind = Me
End Function

Public Function at_css(selector As String) As XmlScraping
    Set element = doc.querySelector(selector)
    
    Set at_css = Me
End Function

Public Function Class(selector As String) As XmlScraping
    Set elements = doc.getElementsByClassName(selector)
    
    Set Class = Me
End Function

Public Function attr(nameAtribute As String) As String
    attr = element.getAttribute(nameAtribute)
End Function

Public Function index(i As Integer) As XmlScraping
    If element Is Nothing Then
        Set element = children(i)
    Else
        Set element = children(i)
    End If
    
    Set index = Me
End Function

Public Function text() As String
    text = element.innerText
End Function

Public Function html() As String
    html = element.innerHTML
End Function

