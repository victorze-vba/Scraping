Attribute VB_Name = "Test_XmlScraping"
Option Explicit

Sub XmlScrapingTest()
    Dim Specs As New SpecSuite
    Dim Browser As New XmlScraping

    Browser.gotoPage "https://stackoverflow.com/"
    
    With Specs.It("Get content of element")
        .Expect(Browser.css("#nav-questions").text).ToEqual "Questions"
    End With
    
    With Specs.It("Get html of element")
        .Expect(Browser.css(".js-gps-track").html).ToEqual "<SPAN class=-img>Stack Overflow</SPAN> "
    End With

    InlineRunner.RunSuite Specs
End Sub

Sub XmlIndexCollection()
    Dim Specs As New SpecSuite
    Dim doc As New XmlScraping

    doc.gotoPage "https://vba-dev.github.io/vba-scraping/"
    
    With Specs.It("Item 0 of Collection")
        .Expect(doc.Class("description").index(0).text).ToEqual ""
    End With

    InlineRunner.RunSuite Specs
End Sub

Sub XmlCollection()
    Dim Specs As New SpecSuite
    Dim doc As New XmlScraping

    doc.gotoPage "https://vba-dev.github.io/vba-scraping/"

    With Specs.It("Select tag")
        .Expect(doc.css("span").index(0).text).ToEqual "<li>"
        .Expect(doc.css("span").index(2).html).ToEqual "&lt;title&gt;"
    End With
Debug.Print doc.css("span").index(0).text
    With Specs.It("Select class")
        .Expect(doc.at_css(".title").text).ToEqual "VBA Scraping"
        .Expect(doc.css(".title").index(0).text).ToEqual "VBA Scraping"
        .Expect(doc.Class("title").index(0).text).ToEqual "VBA Scraping"
    End With
    
    With Specs.It("Select class")
        .Expect(doc.css(".download a").index(0).attr("href")).ToEqual "about:scraping.xlsm"
    End With
    
    InlineRunner.RunSuite Specs
End Sub
