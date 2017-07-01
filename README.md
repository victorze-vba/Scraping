# Scraping web
 Library to extract data from websites
 
 Librería para facilitar la extracción de datos de internet

## Issues
 If you find an error or want to suggest an improvement [write to me](https://github.com/vba-dev/vba-scraping/issues)
 
 Si encuentras un error o quieres sugerir una mejora [escribeme](https://github.com/vba-dev/vba-scraping/issues)

## Tests
 To perform the tests is using the library [vba-tdd](https://github.com/VBA-tools/VBA-TDD)
 
 Para realizar las pruebas se está utilizando la librería [vba-tdd](https://github.com/VBA-tools/VBA-TDD)
 
 ## Examples
 ### Class XmlScraping
 ```vb
Sub Example()

    Dim doc As New XmlScraping

    doc.gotoPage "https://vba-dev.github.io/vba-scraping/"

    Debug.Print doc.css("span").index(5).html
    
    Debug.Print doc.class("title").index(0).text
    
    Debug.Print doc.css(".download a").index(0).attr("href")

End Sub
 
 ```
 
 ### CLass Scraping
 ```vb
 Sub ScrapingTest()

    Dim browser As New Scraping

    browser.gotoPage "https://stackoverflow.com/"
    
    browser.gotoPage "https://stackoverflow.com/", True
    
    browser.Id("nav-questions").Click.Quit

End Sub
```

