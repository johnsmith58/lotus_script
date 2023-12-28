## Export Lotus Notes Attachments using Lotus Script

##### This script is export Lotus Notes DB with csv and embedded object using Lotus Script.

I used the Button click on Document for easy to use, because I am not familiar with Lotus Notes and also I am only supposed to export data with CSV and embedded objects using this script.

I created files for specific function, I wish that will be helpful.

Can be define more than one database.
> dbs(1) = "db1.nsf"

Then don't forget to change your database count.
> Dim dbs(1 To {Int}) As String

I used Document form all of item value, if you used differernt I hope you change here.
> formName$ = "Document"

Create custom sub directory under specific folder
> Mkdir exportAttachmentPath$

</br>
Full Code
</br>
</br>

````
Sub Click(Source As Button)

	Dim sess As New NotesSession
	Dim db As NotesDatabase
	Dim coll As NotesDocumentCollection
	Dim doc As NotesDocument
	Dim tmpStr As String

	' db list
	Dim dbs(1 To 58) As String
	dbs(1) = "db.nsf"

	' define attribute
	dbServer = "dbserver"
	formName$ = "Document"
	rootPath$ = "C:\path\"

	Dim i As Integer
	For i = Lbound(dbs) To Ubound(dbs)
		' Init db
		Set db = sess.GetDataBase(dbServer, dbs(i), False)

		If Not db Is Nothing Then

			' Get all docu
			Set coll = db.AllDocuments
			Set doc = coll.GetFirstDocument()

			' Create Root Dir
			dbTitle = db.Title
			dbTitle = Replace(dbTitle, "/", "_")
			dbTitle = Replace(dbTitle, "\", "_")

			exportRootPath$ = rootPath$ & dbTitle & "\"
			Mkdir exportRootPath$

			' Create attachment folder for store file
			Mkdir exportRootPath$ & "attachments" & "\"

			' Create csv file
			csvFileName = dbTitle & ".csv"
			csvFilePath = exportRootPath$ & csvFileName

			' Open csv file for write
			csvFileNum = Freefile
			Open csvFilePath For Output As csvFileNum
			' Support multiple lang
			Print #csvFileNum, Chr(239) & Chr(187) & Chr(191)

			' Define form field for csv header
			Dim form As NotesForm
			Set form = db.GetForm(formName$)

			' Write csv header
			Dim header As String
			header = ""
			header = header & "NoteID" & ","
			Forall field In form.Fields
				header = header & field & ","
			End Forall

			header = header & "Created Date" & ","
			header = header & "Created By" & ","
			header = header & "Lat Modified" & ","

			header = Left(header, Len(header) -1)
			Print #csvFileNum, header

			' Loop doc
			While Not(doc Is Nothing)

				uuid = doc.NoteID
				If uuid <> "" Then

					' Write csv body
					csvLine = ""
					csvLine = csvLine & uuid & ","

					Forall field In form.Fields
						itemName = field
						If doc.HasItem(itemName) Then
							Set subjItem = doc.GetFirstItem(itemName)
							tmpStr = ""
							If(subjItem.Type = RICHTEXT) Then
								tmpStr = subjItem.GetFormattedText(False, 0)
							Else
								itemValue = doc.GetItemValue(itemName)
								If Isarray(itemValue) Then
									tmpStr = itemValue(0)
								Else
									tmpStr = subjItem.Text
								End If
							End If
							tmpStr = Replace(tmpStr, Chr(13) & Chr(10), "\n")
							tmpStr = Replace(tmpStr, ",", "\ca")
							csvLine = csvLine & tmpStr
						Else
							csvLine = csvLine & "NULL" & ","
						End If
					End Forall

					csvLine = csvLine & doc.Created & ","
					csvLine = csvLine & doc.Authors(0) & ","
					csvLine = csvLine & doc.LastModified & ","

					Print #csvFileNum, csvLine

					' Attachments
					' Create folder for each document
					uuid = Replace(uuid, "/", "_")
					uuid = Replace(uuid, "\", "_")

					exportAttachmentPath$ = exportRootPath$ & "attachments" & "\" & uuid & "\"
					Mkdir exportAttachmentPath$

					' Get file place
					Dim body As NotesRichTextItem
					Dim rtnav As NotesRichTextNavigator

					Set body = doc.GetFirstItem("Body")

					If Not body Is Nothing Then
						Set rtnav = body.CreateNavigator

						' Store attachment files
						If rtnav.FindFirstElement(RTELEM_TYPE_FILEATTACHMENT) Then
							Do
								Set att = rtnav.GetElement()
								exportFilepath$ = exportAttachmentPath$ & att.Source
								Call att.ExtractFile(exportFilepath$)
							Loop While rtnav.FindNextElement()
						End If

					End If

				Else
					Print "Found!"
					Print "uuid: "& uuid
				End If

				Set doc = coll.GetNextDocument(doc)

			Wend

			' Close csv file
			Close csvFileNum

		Else
			Print "db not found: " & dbs(i)
		End If
	Next

End Sub
