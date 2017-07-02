Attribute VB_Name = "examples"
Option Explicit

Sub do_a_search_on_wikipedia()

    Dim doc As New Scraping
    Dim search As String
    
    search = "document object model"
    
    doc.gotoPage "https://en.wikipedia.org/wiki/Main_Page", True 'browser visible

    doc.id("searchInput").fieldValue search
    doc.id("searchButton").click

End Sub

Sub extract_the_titles_of_the_questions_in_stackoverflow()

    Dim i As Integer
    Dim doc As New XmlScraping
    Dim numberTitles As Integer
    
    doc.gotoPage "https://stackoverflow.com/"
    
    numberTitles = doc.css(".summary h3 a").count
    
    Workbooks.Add
    
    For i = 0 To numberTitles - 1
        Cells(i + 1, 1) = doc.css(".summary h3 a").index(i).text
    Next i
    
End Sub

Sub find_a_member()

    Dim miembro As String
    Dim numeroMatricula As String
    Dim condicion As String
    Dim doc As New Scraping
    
    doc.gotoPage "http://www.ccppuno.pe/web/index.php/colegiados/miembros-ordinarios"
    
    numeroMatricula = "0333"
    
    doc.id("search-searchword").fieldValue numeroMatricula
    doc.id("searchphrasematricula").click 'select a radio element
    doc.css(".btn").index(0).click
    
    miembro = doc.css(".categoryResults tbody tr td").index(1).text
    condicion = doc.css(".categoryResults tbody tr td").index(2).text
    
    Debug.Print "Miembro del colegio de contadores de Puno"
    Debug.Print "Num Matrícula: " & numeroMatricula
    Debug.Print "Name: " & miembro
    Debug.Print "Condición: " & condicion
    
End Sub

' Busca a todas las personas que tengan un nombre específico
Sub Seeks_members_of_the_school_of_accountants_by_name()

    Dim numMember As String
    Dim doc As New Scraping
    Dim i As Integer
    Dim row As MSHTML.HTMLTableRow

    doc.gotoPage "http://www.ccppuno.pe/web/index.php/colegiados/miembros-ordinarios"
    
    doc.id("search-searchword").fieldValue "Victor" 'pon tu nombre
    doc.css(".btn").index(0).click
    
    numMember = doc.css(".categoryResults tbody tr").count
    
    Workbooks.Add
    
    For i = 0 To numMember - 1
        Set row = doc.css(".categoryResults tbody tr").index(i).rowTable
        Cells(i + 1, 1) = row.Cells.item(1).innerText
    Next i

End Sub

' Extrae el tipo de cambio de la Sunat
Sub extract_exchange_rates()

    On Error Resume Next

    Dim i As Integer
    Dim f As Integer
    Dim doc As New Scraping
    Dim row As MSHTML.HTMLTableRow

    doc.gotoPage "http://www.sunat.gob.pe/cl-at-ittipcam/tcS01Alias"

    doc.css("select").index(0).fieldValue "03"    'month (mes)
    doc.css("select").index(1).fieldValue "2017"  'year (año)
    doc.css(".button").index(0).click sleep:=1

    Workbooks.Add
    Cells(1, 1) = doc.css("h3").index(0).text
    Cells(3, 1) = "Día"
    Cells(3, 2) = "Compra"
    Cells(3, 3) = "Venta"

    f = 4
    For i = 2 To 7
        Set row = doc.css("form table>tbody>tr").index(i).rowTable

        Cells(f, 1) = row.Cells.item(0).innerText
        Cells(f, 2) = row.Cells.item(1).innerText
        Cells(f, 3) = row.Cells.item(2).innerText
        f = f + 1
        Cells(f, 1) = row.Cells.item(3).innerText
        Cells(f, 2) = row.Cells.item(4).innerText
        Cells(f, 3) = row.Cells.item(5).innerText
        f = f + 1
        Cells(f, 1) = row.Cells.item(6).innerText
        Cells(f, 2) = row.Cells.item(7).innerText
        Cells(f, 3) = row.Cells.item(8).innerText
        f = f + 1
        Cells(f, 1) = row.Cells.item(9).innerText
        Cells(f, 2) = row.Cells.item(10).innerText
        Cells(f, 3) = row.Cells.item(11).innerText
        f = f + 1
    Next i

End Sub

Sub select_a_checkbox_type_element()

    Dim doc As New Scraping
    
    doc.gotoPage "https://styde.net/login/", True

    doc.id("rememberme").click

End Sub
