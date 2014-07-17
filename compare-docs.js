var wdPaneRevisions = 18

if (WScript.Arguments.Length != 2) {
    WScript.Echo("Drag & drop two Microsoft Word files.")
}

var app = WScript.CreateObject("Word.Application")
var doc1 = app.Documents.Open(WScript.Arguments(0))
var doc2 = app.Documents.Open(WScript.Arguments(1))

app.Application.CompareDocuments(doc1, doc2)

doc1.Close()
doc2.Close()

do {
    app.ActiveWindow.View.SplitSpecial = wdPaneRevisions
} while (app.ActiveWindow.View.SplitSpecial != wdPaneRevisions)

app.Visible = true
