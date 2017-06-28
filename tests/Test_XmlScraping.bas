Attribute VB_Name = "Test_XmlScraping"
Option Explicit

Sub XmlScrapingTest()
    Dim Specs As New SpecSuite
    Dim Browser As New XmlScraping

    Browser.GotoPage "https://stackoverflow.com/"
    
    With Specs.It("Get content of element")
        .Expect(Browser.Css("#nav-questions").Text).ToEqual "Questions"
    End With
    
    With Specs.It("Get html of element")
        .Expect(Browser.Css(".js-gps-track").Html).ToEqual "<SPAN class=-img>Stack Overflow</SPAN> "
    End With

    InlineRunner.RunSuite Specs
End Sub
