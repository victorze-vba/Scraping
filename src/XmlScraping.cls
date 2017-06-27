VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XmlScraping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

