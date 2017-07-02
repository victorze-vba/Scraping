Attribute VB_Name = "Test_Scraping"
Option Explicit

Private Sub test_at_css_and_id()
    Dim Specs As New SpecSuite
    Dim doc As New Scraping

    doc.gotoPage "https://stackoverflow.com/"
    
    With Specs.It("Extract the text of the element with id 'nav-questions'")
        .Expect(doc.at_css("#nav-questions").text).ToEqual "Questions"
    End With
    
    With Specs.It("Extract the text of the element with id 'nav-questions'")
        .Expect(doc.id("nav-questions").text).ToEqual "Questions"
    End With
    
    InlineRunner.RunSuite Specs
End Sub

Private Sub test_css_index_and_count()
    Dim Specs As New SpecSuite
    Dim doc As New Scraping

    doc.gotoPage "https://stackoverflow.com/"
    
    With Specs.It("Count the number of questions on the main page of these stackoverflow")
        .Expect(doc.css(".summary h3 a").count).ToEqual 96
    End With
    
    'Extract the first three questions from stasadsf
    Debug.Print doc.css(".summary h3 a").index(0).text
    Debug.Print doc.css(".summary h3 a").index(1).text
    Debug.Print doc.css(".summary h3 a").index(2).text
    
    InlineRunner.RunSuite Specs
End Sub

Private Sub test_html()
    Dim Specs As New SpecSuite
    Dim doc As New Scraping

    doc.gotoPage "https://stackoverflow.com/"
    
    Debug.Print doc.css(".js-gps-track").index(0).html
    
    'Works but can not be tested
'    With Specs.It("Get html of element")
'        .Expect(doc.css(".js-gps-track").index(0).html).ToEqual "<span class=-img>Stack Overflow</span>"
'    End With
    
    InlineRunner.RunSuite Specs
End Sub

Private Sub test_send_form_and_extract_data_from_a_table()
    Dim Specs As New SpecSuite
    Dim doc As New Scraping

    doc.gotoPage "http://www.sunat.gob.pe/cl-at-ittipcam/tcS01Alias"
    
    doc.css("select").index(0).fieldValue "03"
    doc.css("select").index(1).fieldValue "2016"
    doc.css(".button").index(0).click sleep:=1
    
    With Specs.It("Extract the date")
        .Expect(doc.css("h3").index(0).text).ToEqual "Marzo - 2016"
    End With
    
    InlineRunner.RunSuite Specs
End Sub

Private Sub test_select_radio_item()
    Dim Specs As New SpecSuite
    Dim doc As New Scraping
    Dim name As String

    doc.gotoPage "http://www.ccppuno.pe/web/index.php/colegiados/miembros-ordinarios"
    
    doc.id("search-searchword").fieldValue "1520"
    doc.id("searchphrasematricula").click
    doc.css(".btn").index(0).click
    
    name = doc.css(".categoryResults tbody tr td").index(1).text
    
    With Specs.It("Extract the name")
        .Expect(name).ToEqual "ZEVALLOS SUAÑA, LUIS"
    End With
    
    InlineRunner.RunSuite Specs
End Sub



