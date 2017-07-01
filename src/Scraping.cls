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
' Library used: Microsoft Internet Controls, Microsoft HTML Object Library
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

Private elements As MSHTML.IHTMLElementCollection

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
' Select a collection of elements with the class className
'
' @param {String} className
' @return Scraping
''
Public Function Class(className As String) As Scraping
    Set elements = doc.getElementsByClassName(className)
    
    Set Class = Me
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
' Returns the number of items in the collection
'
' @return Integer
''
Public Function count() As Integer
    count = children.Length
End Function

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
' Click an element
'
' @return void
''
Public Sub click()
    element.click

    loadBrowser

    sleep 3

    Set click = Me
End Sub

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
Private Sub sleep(Seconds As Byte)
    Application.Wait (Now + TimeValue("0:00:0" + CStr(Seconds)))
End Sub

''
' Runs when instantiating on object
'
' @return void
''
Private Sub Class_Initialize()
    
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



