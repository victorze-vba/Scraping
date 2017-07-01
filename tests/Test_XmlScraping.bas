Attribute VB_Name = "Test_XmlScraping"
Option Explicit

Sub TestTextHtml()
    Dim Specs As New SpecSuite
    Dim Doc As New XmlScraping

    Doc.gotoPage "https://stackoverflow.com/"
    
    With Specs.It("Extract the text of an element with id 'nav-questions'")
        .Expect(Doc.id("nav-questions").text).ToEqual "Questions"
    End With
    
    With Specs.It("Extract to the html of the first element of a collection with class .js-gps-track")
        .Expect(Doc.css(".js-gps-track").index(0).html).ToEqual "<SPAN class=-img>Stack Overflow</SPAN> "
    End With

    InlineRunner.RunSuite Specs
End Sub

Sub TestCollection()
    Dim Specs As New SpecSuite
    Dim Doc As New XmlScraping

    Doc.gotoPage "https://vba-dev.github.io/vba-scraping/"

    With Specs.It("Select tag")
        .Expect(Doc.css("span").index(0).text).ToEqual "<li>"
        .Expect(Doc.css("span").index(2).html).ToEqual "&lt;title&gt;"
    End With

    With Specs.It("Select class")
        .Expect(Doc.at_css(".title").text).ToEqual "VBA Scraping"
        .Expect(Doc.css(".title").index(0).text).ToEqual "VBA Scraping"
        .Expect(Doc.Class("title").index(0).text).ToEqual "VBA Scraping"
    End With
    
    With Specs.It("Select class")
        .Expect(Doc.css(".download a").index(0).attr("href")).ToEqual "about:examples.xlsm"
    End With
    
    InlineRunner.RunSuite Specs
End Sub

