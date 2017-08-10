# Web Scraping
 Extract data from websites easily.

 Extrae datos de sitios web facilmente.

## Issues
 If you find an error or want to suggest an improvement [write to me](https://github.com/vba-dev/vba-scraping/issues)

 Si encuentras un error o quieres sugerir una mejora [escríbeme](https://github.com/vba-dev/vba-scraping/issues)

## Tests
 To perform the tests is using the library [vba-tdd](https://github.com/VBA-tools/VBA-TDD)

 Para realizar las pruebas estoy utilizando la librería [vba-tdd](https://github.com/VBA-tools/VBA-TDD)

 ## Examples
```vb
Sub do_a_search_on_wikipedia()

    Dim doc As New Scraping
    Dim search As String

    search = "document object model"

    doc.gotoPage "https://en.wikipedia.org/wiki/Main_Page", True 'browser visible

    doc.id("searchInput").fieldValue search
    doc.id("searchButton").click

End Sub
```

```vb
Sub extract_the_titles_of_the_questions_in_stackoverflow()

    Dim i As Integer
    Dim doc As New XmlScraping
    Dim numberTitles As Integer

    doc.gotoPage "https://stackoverflow.com/"

    numberTitles = doc.css(".summary h3 a").count

    For i = 0 To numberTitles - 1
        Cells(i + 1, 1) = doc.css(".summary h3 a").index(i).text
    Next i

End Sub
```
