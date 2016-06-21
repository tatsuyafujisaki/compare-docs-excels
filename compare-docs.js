if (WScript.Arguments.length !== 2) {
  WScript.Echo("Please drag & drop two Microsoft Word files.");
  WScript.Quit();
}

var app = WScript.CreateObject("Word.Application");
var doc1 = app.Documents.Open(WScript.Arguments.Item(0));
var doc2 = app.Documents.Open(WScript.Arguments.Item(1));

app.CompareDocuments(doc1, doc2);

doc1.Close();
doc2.Close();

var wdPaneRevisions = 18;

app.ActiveWindow.View.SplitSpecial = wdPaneRevisions;
app.Visible = true;