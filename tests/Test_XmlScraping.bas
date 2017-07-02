Attribute VB_Name = "Test_XmlScraping"
Option Explicit

Sub TestTextHtml()
    Dim Specs As New SpecSuite
    Dim doc As New XmlScraping

    doc.gotoPage "https://stackoverflow.com/"
    
    With Specs.It("Extract the text of an element with id 'nav-questions'")
        .Expect(doc.id("nav-questions").text).ToEqual "Questions"
    End With
    
    With Specs.It("Extract to the html of the first element of a collection with class .js-gps-track")
        .Expect(doc.css(".js-gps-track").index(0).html).ToEqual "<SPAN class=-img>Stack Overflow</SPAN> "
    End With

    InlineRunner.RunSuite Specs
End Sub

Sub TestCollection()
    Dim Specs As New SpecSuite
    Dim doc As New XmlScraping

    doc.gotoPage "https://vba-dev.github.io/vba-scraping/"

    With Specs.It("Select tag")
        .Expect(doc.css("span").index(0).text).ToEqual "<a>"
        .Expect(doc.css("span").index(2).html).ToEqual "Submit"
    End With

    With Specs.It("Select class")
        .Expect(doc.at_css(".title").text).ToEqual "VBA Scraping"
        .Expect(doc.css(".title").index(0).text).ToEqual "VBA Scraping"
    End With
    
    With Specs.It("Select class")
        .Expect(doc.css(".download a").index(0).attr("href")).ToEqual "about:examples.xlsm"
    End With
    
    InlineRunner.RunSuite Specs
End Sub

