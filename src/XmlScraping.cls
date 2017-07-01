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

Private doc As New MSHTML.HTMLDocument

Private element As MSHTML.IHTMLElement

Private elements As MSHTML.IHTMLElementCollection

Private children As MSHTML.IHTMLDOMChildrenCollection

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

    doc.body.innerHTML = XMLPage.responseText
End Sub

''
' Select a collection of elements
'
' @param {String} selector
' @return XmlScraping
'
' Example selector  id: "#nameId" | class: ".nameClas" | tag: "nameTag"
''
Public Function css(selector As String) As XmlScraping
    Set children = doc.querySelectorAll(selector)
    
    Set css = Me
End Function

''
' Select the first element in the collection
'
' @param {String} selector
' @return XmlScraping
''
Public Function at_css(selector As String) As XmlScraping
    Set element = doc.querySelector(selector)
    
    Set at_css = Me
End Function

''
' Select a collection of elements with the class className
'
' @param {String} className
' @return XmlScraping
''
Public Function Class(className As String) As XmlScraping
    Set elements = doc.getElementsByClassName(className)
    
    Set Class = Me
End Function

''
' Select an item from an elements collection
'
' @param {Integer} i
' @return XmlScraping
''
Public Function index(i As Integer) As XmlScraping
    Set element = children(i)

    Set index = Me
End Function

''
' Returns the number of items in the collection
'
' @return Integer
''
Public Function count() As Integer
    count = children.Length
End Function


''
' Select an elements with the specified id
'
' @param {String} idName
' @return XmlScraping
''
Public Function id(idName As String) As XmlScraping
    Set element = doc.getElementById(idName)
    
    Set id = Me
End Function

''
' Returns the text of the selected element
'
' @return String
''
Public Function text() As String
    text = element.innerText
End Function


''
' Returns the html of the selected element
'
' @return String
''
Public Function html() As String
    html = element.innerHTML
End Function

''
' Returns the value of the specified attribute
'
' @param {String} atributeName
' @return String
''
Public Function attr(atributeName As String) As String
    attr = element.getAttribute(atributeName)
End Function

