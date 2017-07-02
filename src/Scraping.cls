VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scraping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' Scraping v0.1.1 Alpha
' (c) Victor Zevallos - https://github.com/vba-dev/vba-scraping
'
' Library used: Microsoft Internet Controls,
'               Microsoft HTML Object Library
'
' @class Scraping
' @author victorzevallos@protonmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Private ie As SHDocVw.InternetExplorer

Private doc As New MSHTML.HTMLDocument

Private element As MSHTML.IHTMLElement

Private children As MSHTML.IHTMLDOMChildrenCollection

' --------------------------------------------- '
' Public Methods
' --------------------------------------------- '

''
' Visit a url and you can choose to view the browser
'
' @param {String} url
' @param {Boolean} visibleBrowser
' @return void
''
Public Sub gotoPage(url As String, Optional visibleBrowser As Boolean = False)
    Set ie = New SHDocVw.InternetExplorer

    ie.visible = visibleBrowser

    ie.Navigate url
    
    loadBrowser
    
    Set doc = ie.Document
End Sub

''
' Select a collection of elements
'
' @param {String} selector
' @return Scraping
'
' Example selector  id: "#nameId" | class: ".nameClas" | tag: "nameTag"
''
Public Function css(selector As String) As Scraping
    Set children = doc.querySelectorAll(selector)
    
    Set css = Me
End Function

''
' Select the first element in the collection
'
' @param {String} selector
' @return Scraping
''
Public Function at_css(selector As String) As Scraping
    Set element = doc.querySelector(selector)
    
    Set at_css = Me
End Function

''
' Select an item from an elements collection
'
' @param {Integer} i
' @return Scraping
''
Public Function index(i As Integer) As Scraping
    Set element = children(i)

    Set index = Me
End Function

''
' Return to an object to work with tables
'
' @return MSHTML.HTMLTableRow
''
Public Function rowTable() As MSHTML.HTMLTableRow
    Set rowTable = element
End Function

''
' Return the number of items in the collection
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
' @return Scraping
''
Public Function id(idName As String) As Scraping
    Set element = doc.getElementById(idName)

    Set id = Me
End Function

''
' Return the text of the selected element
'
' @return String
''
Public Function text() As String
    text = element.innerText
End Function

''
' Return the html of the selected element
'
' @return String
''
Public Function html() As String
    html = element.innerHTML
End Function

''
' Return the value of the specified attribute
'
' @param {String} atributeName
' @return String
''
Public Function attr(atributeName As String) As String
    attr = element.getAttribute(atributeName)
End Function

''
' Click an element
'
' @return void
''
Public Sub click(Optional sleep As Byte = 3)
    element.click
    
    loadBrowser
    
    pause sleep
    
    Set doc = ie.Document
End Sub

''
' Assigns a value to a text or select element
'
' @return void
''
Public Function fieldValue(str As String) As Scraping
    element.value = str

    Set fieldValue = Me
End Function

' --------------------------------------------- '
' Private Methods
' --------------------------------------------- '

''
' Wait for the browser to load
'
' @return void
''
Private Sub loadBrowser()
    Do While ie.ReadyState <> READYSTATE_COMPLETE Or ie.Busy
    Loop
End Sub

''
' Pause in seconds
'
' @return void
''
Private Sub pause(Seconds As Byte)
    Application.wait (Now + TimeValue("0:00:0" + CStr(Seconds)))
End Sub

''
' Runs when object is no longer used
'
' @return void
''
Private Sub Class_Terminate()
    If Not ie.visible Then ie.quit
    
    Set ie = Nothing
End Sub
