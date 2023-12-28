Sub Click(Source As Button)

	Dim sess As New NotesSession
	Dim db As NotesDatabase
	Dim coll As NotesDocumentCollection
	Dim doc As NotesDocument
	Dim subj As Variant
	Dim csvFileName As String
	Dim csvFileNum As Integer
	Dim header As String
	Dim csvLine As String
	Dim bodyItem As NotesRichTextItem
	Dim plainText As String
	Dim itemValue  As Variant
  Dim form As NotesForm
  Dim view As NotesView

	Set db = sess.CurrentDatabase
	Set coll = db.AllDocuments
	Set doc = coll.GetFirstDocument()

  formName$ = "Document"
	Set form = db.GetForm(formName$)

	csvFileName = "C:\export_path\export.csv"
	csvFileNum = Freefile
	Open csvFileName For Output As csvFileNum

  ' Support for even English
	Print #csvFileNum, Chr(239) & Chr(187) & Chr(191)

	Forall field In form.Fields
		header = header & field & ","
	End Forall

	header = Left(header, Len(header) -1)
	Print #csvFileNum, header


	While Not(doc Is Nothing)
		csvLine = ""

		' Loop through all items in the document
		Forall field In form.Fields
			itemName = field
			If doc.HasItem(itemName) Then
				Set subjItem = doc.GetFirstItem(itemName)
				If(subjItem.Type = RICHTEXT) Then
					Print "RichText: " & itemName
					plainText = subjItem.GetFormattedText(False, 0)
					plainText = Replace(plainText, Chr(13) & Chr(10), "")
					csvLine = csvLine & plainText & ","
				Else

					itemValue = doc.GetItemValue(itemName)
					If Isarray(itemValue) Then
						csvLine = csvLine & itemValue(0) & ","
					Else
						csvLine = csvLine & subjItem.Text & ","
					End If
				End If

			End If
		End Forall

    ' Write in csv with new line
		Print #csvFileNum, csvLine

    ' Add next doc for the loop
		Set doc = coll.GetNextDocument(doc)
	Wend

	Close csvFileNum

End Sub
