Attribute VB_Name = "Test_XmlScraping"
Option Explicit

Sub XmlScrapingTest()
    Dim Specs As New SpecSuite
    Dim browser As New XmlScraping

    browser.GotoPage "https://stackoverflow.com/"
    
    With Specs.It("Get content of element")
        .Expect(browser.Css("#nav-questions").Text).ToEqual "Questions"
    End With
    
    With Specs.It("Get html of element")
        .Expect(browser.Css(".js-gps-track").Html).ToEqual "<SPAN class=-img>Stack Overflow</SPAN> "
    End With

    InlineRunner.RunSuite Specs
End Sub
