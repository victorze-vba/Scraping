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

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Private Doc As New MSHTML.HTMLDocument

Private Element As MSHTML.IHTMLElement

Private Elements As MSHTML.IHTMLElementCollection

Private Children As MSHTML.IHTMLDOMChildrenCollection

' --------------------------------------------- '
' Public Methods
' --------------------------------------------- '

''
' Visit the url
'
' @param {String} url
' @return void
''
Public Sub gotoPage(url As String)
    Dim XMLPage As New MSXML2.XMLHTTP60

    XMLPage.Open "GET", url, False
    XMLPage.send

    Doc.body.innerHTML = XMLPage.responseText
End Sub

''
' Select a collection of items
'
' @param {String} selector
' @return XmlScraping
'
' Example selector  id: "#nameId" | class: ".nameClas" | tag: "nameTag"
''
Public Function css(selector As String) As XmlScraping
    Set Children = Doc.querySelectorAll(selector)
    
    Set css = Me
End Function

''
' Select the first item in the collection
'
' @param {String} selector
' @return XmlScraping
''
Public Function at_css(selector As String) As XmlScraping
    Set Element = Doc.querySelector(selector)
    
    Set at_css = Me
End Function

''
' Select a collection of elements with the class nameClass
'
' @param {String} nameClass
' @return XmlScraping
''
Public Function Class(nameClass As String) As XmlScraping
    Set Elements = Doc.getElementsByClassName(nameClass)
    
    Set Class = Me
End Function

''
' Select an elements with the specified id
'
' @param {String} nameId
' @return XmlScraping
''
Public Function id(nameId As String) As XmlScraping
    Set Element = Doc.getElementById(nameId)
    
    Set id = Me
End Function

''
' Returns the value of the specified attribute
'
' @param {String} nameAtribute
' @return String
''
Public Function attr(nameAtribute As String) As String
    attr = Element.getAttribute(nameAtribute)
End Function

''
' Select an item from an item collection
'
' @param {Integer} i
' @return XmlScraping
''
Public Function index(i As Integer) As XmlScraping
    If Element Is Nothing Then
        Set Element = Children(i)
    Else
        Set Element = Children(i)
    End If
    
    Set index = Me
End Function

''
' Returns the text of the selected element
'
' @return String
''
Public Function text() As String
    text = Element.innerText
End Function


''
' Returns the html of the selected element
'
' @return String
''
Public Function html() As String
    html = Element.innerHTML
End Function

