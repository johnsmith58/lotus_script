Sub Click(Source As Button)

	Dim sess As New NotesSession
	Dim db As NotesDatabase
	Dim coll As NotesDocumentCollection
	Dim doc As NotesDocument
	Dim attachment As NotesEmbeddedObject
	Dim attachmentCount As Integer
	Dim exportPath As String
	Dim rtitem As Variant
	Dim body As NotesRichTextItem
	Dim rtnav As NotesRichTextNavigator
	Dim subject As String
	Dim exportRootPath As String

	Set db = sess.CurrentDatabase
	Set coll = db.AllDocuments
	Set doc = coll.GetFirstDocument()

	exportRootPath$ =  "C:\export_path\attachments\"

	While Not(doc Is Nothing)

		Set body = doc.GetFirstItem("Body")
		Set rtnav = body.CreateNavigator

		subject = Replace(doc.GetFirstItem("Subject").Text, " ", "_")
		subject = Replace(subject, "/", "_")

    ' Define sub directory using subject value
		exportPath$ = exportRootPath$ & subject & "\"

    ' Create sub directory for specific NotesDocument because have a more than one in one NotesDocument
		Mkdir exportPath

		If rtnav.FindFirstElement(RTELEM_TYPE_FILEATTACHMENT) Then
			Do
        ' Get attachment file element
				Set att = rtnav.GetElement()
				filepath$ = exportPath$ & att.Source
        
        ' Store in sub directory
				Call att.ExtractFile(filepath$)
			Loop While rtnav.FindNextElement()
		End If

		Set doc = coll.GetNextDocument(doc)

	Wend

End Sub
