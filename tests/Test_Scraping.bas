Attribute VB_Name = "Test_Scraping"
Option Explicit

Sub Test()
    'Call test_visita_paginas_web
    Call test_llenar_y_enviar_formulario
End Sub

Sub ScrapingTest()
    Dim Specs As New SpecSuite

    Dim Browser As New Scraping

    Browser.GotoPage "https://stackoverflow.com/"
    
    With Specs.It("Get content of element")
        .Expect(Browser.Css("#nav-questions").Text).ToEqual "Questions"
    End With

    With Specs.It("Get html of element")
        .Expect(Browser.Css(".js-gps-track").Html).ToEqual "<span class=-img>Stack Overflow</span>"
    End With
    
    Browser.GotoPage "https://stackoverflow.com/", True
    Browser.Id("nav-questions").Click.Quit
    
    
    InlineRunner.RunSuite Specs
End Sub

Sub ColletionTest()
    Dim Specs As New SpecSuite

    Dim Browser As New Scraping

    Browser.GotoPage "https://stackoverflow.com/"
End Sub

