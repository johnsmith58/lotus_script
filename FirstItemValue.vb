Sub Click(Source As Button)

	Dim sess As New NotesSession
	Dim db As NotesDatabase
	Dim coll As NotesDocumentCollection
	Dim doc As NotesDocument

	Set db = sess.CurrentDatabase
	Set coll = db.AllDocuments
	Set doc = coll.GetFirstDocument()

	Print "Subject: " & doc.GetItemValue("Subject")(0)

End Sub
