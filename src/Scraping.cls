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

Private IE As SHDocVw.InternetExplorer

Private Doc As New MSHTML.HTMLDocument

Private Element As MSHTML.IHTMLElement

Private Elements As MSHTML.IHTMLElementCollection

Public Sub GotoPage(Url As String, Optional VisibleBrowser As Boolean = False)
    Set IE = New SHDocVw.InternetExplorer

    IE.visible = VisibleBrowser

    IE.Navigate Url
    
    LoadBrowser
    
    Set Doc = IE.Document
End Sub

Public Function Id(IdName As String) As Scraping
    Set Element = Doc.getElementById(IdName)

    Set Id = Me
End Function

Public Function Class(ClassName As String) As Scraping
    Set Elements = Doc.getElementsByClassName(ClassName)
    
    Set Class = Me
End Function

Public Function Index(indexColl As Integer) As Scraping
    Set Element = Elements(indexColl)

    Set Class = Me
End Function

Public Function Css(Selector As String) As Scraping
    Set Element = Doc.querySelector(Selector)
    
    Set Css = Me
End Function

Public Function Text() As String
    Text = Element.innerText
End Function

Public Function Html() As String
    Html = Element.innerHTML
End Function

Public Function Click() As Scraping
    Element.Click
    
    LoadBrowser

    Sleep 3
    
    Set Click = Me
End Function

Public Sub Home()
    IE.GoHome
End Sub

Public Sub Quit()
    If IE.visible Then IE.Quit
End Sub

Public Function Url()
    Url = IE.LocationURL
End Function

' --------------------------------------------- '
' Private Methods
' --------------------------------------------- '

Private Sub LoadBrowser()
    Do While IE.ReadyState <> READYSTATE_COMPLETE Or IE.Busy
    Loop
End Sub

Private Sub Sleep(Seconds As Byte)
    Application.Wait (Now + TimeValue("0:00:0" + CStr(Seconds)))
End Sub
 
Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    If Not IE.visible Then IE.Quit
    
    Set IE = Nothing
End Sub



