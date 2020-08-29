 'Call ToPDF(finalDest, Item.SenderName, Item.body, Item.Subject)
        

Sub ToPDF(strChemin As String, sender As String, body As String, object As String)

    
    Set wordapp = CreateObject("Word.Application")
    Set wordDoc = wordapp.Documents.Add
    'Set wordDoc = wordapp.Documents.Open(strChemin & "\" & "Temp.doc")
    
        wordapp.Visible = True
 
        With wordapp.Selection
        .Paragraphs.Space1
        .Paragraphs.LineSpacingRule = 0
        .MoveDown , Count:=0
        .TypeText Text:=Chr(0)
        .TypeParagraph
        .TypeText Text:=sender
        .TypeParagraph
        .ParagraphFormat.SpaceAfter = 0
        .TypeText Text:=body
        .TypeParagraph
        End With

        Set pdfjob = CreateObject("PDFCreator.clsPDFCreator")
        With pdfjob
            If .cstart("/NoProcessingAtStartup") = False Then
                MsgBox "Initialisation de PDFCreator impossible", vbCritical + vbOKOnly, "PrtPDFCreator"
                
                Exit Sub
            End If
        .cOption("UseAutosave") = 1
        .cOption("UseAutosaveDirectory") = 1
        .cOption("AutosaveDirectory") = strChemin
        .cOption("AutosaveFilename") = object
        .cOption("AutosaveFormat") = 0
        .cClearCache
        End With
        ActivePrinter = "PDFCreator"
        ActiveWindow.SelectedSheets.PrintOut Copies:=1
        'Application.PrintOut copies:=1
        Do Until pdfjob.cCountOfPrintjobs = 1
        DoEvents
        Loop
        pdfjob.cPrinterStop = False
        Do Until pdfjob.cCountOfPrintjobs = 0
        DoEvents
        Loop
        With pdfjob
        .cDefaultprinter = DefaultPrinter
        .cClearCache
        .cClose
        End With
End Sub
