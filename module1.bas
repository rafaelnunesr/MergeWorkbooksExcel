Public LastColumnToMerge As String
Public LastRowToMerge As Integer
Public JumpEmptyRow As Integer
Public NameSheetImported As String
Public MergingDestination As String
Public TotalSheetsImported As Integer
Public strRangeToCheck As String
Public Flag As Variant

Function CopyPasteFirstWorkSheet()

    Sheets(NameSheetImported).Range(strRangeToCheck).Copy
    Sheets(MergingDestination).Range(strRangeToCheck).PasteSpecial xlPasteValues

End Function

Function SanitizeWorksheets()

    Dim iRow As Long
    Dim rowIndex, maxEmptyCell As Integer
    Dim row As String
    
    rowIndex = 0
    maxEmptyCell = 500
    
    Sheet = Worksheets(NameSheetImported).Range(strRangeToCheck)
    
    For iRow = LBound(Sheet, 1) To UBound(Sheet, 1)
        
        If rowIndex < maxEmptyCell Then
    
            If Sheet(iRow, 3) <> "" Then
                rowIndex = 0
                If Sheet(iRow, 2) <> NameSheetImported Then
                    row = "A" & iRow & ":" & LastColumnToMerge & iRow
                    Worksheets(NameSheetImported).Range(row).EntireRow.ClearContents
                End If
                
            Else
                rowIndex = rowIndex + 1
                
            End If
        
        Else
            Exit For
        
        End If

    Next iRow

End Function

Function CleanDestinationSheet()
    
    Dim RangeToClean As String
    RangeToClean = "A2:" & LastColumnToMerge & LastRowToMerge
    
    Worksheets(MergingDestination).Range(RangeToClean).ClearContents

End Function

Function DeleteImportedSheets()
    
    Application.DisplayAlerts = False
    Worksheets(NameSheetImported).Delete
    Application.DisplayAlerts = True

End Function


Function MergeValues()
    
    Dim iRow As Long
    Dim iCol As Long
    
    Destination = Worksheets(MergingDestination).Range(strRangeToCheck)
    Sheet = Worksheets(NameSheetImported).Range(strRangeToCheck)
    
    For iRow = LBound(Sheet, 1) To UBound(Sheet, 1)
        
        For iCol = LBound(Sheet, 2) To UBound(Sheet, 2)
    
            If Sheet(iRow, 2) <> "" Then
                If Destination(iRow, 2) = NameSheetImported Or Destination(iRow, 2) = "" Then
                    Worksheets(NameSheetImported).Activate
                        
                    Range("A1").Cells(iRow, iCol).Copy
                    Sheets(MergingDestination).Range("A1").Cells(iRow, iCol).PasteSpecial xlPasteValues
                    
                 Else
                 
                    CleanDestinationSheet
                    
                    Dim Msg, Style, Title, Response, SheetA, SheetB
                    currentAdressChecking = Worksheets(MergingDestination).Range(strRangeToCheck).Cells(iRow, 2).Address(False, False)
                            
                    SheetA = Destination(iRow, 2)
                    SheetB = NameSheetImported
                    Flag = True
                                            
                    Msg = "As planilhas " & SheetA & " e " & SheetB _
                        & " possuem conflitos de vendedores na célula " & currentAdressChecking & " ." _
                        & vbNewLine _
                        & vbNewLine _
                        & vbNewLine _
                        & "               " & currentAdressChecking & ": " & Destination(iRow, 2) _
                        & vbNewLine _
                        & "               " & currentAdressChecking & ": " & Sheet(iRow, 2) _
                        & vbNewLine _
                        & vbNewLine _
                        & vbNewLine _
                        & "Por favor, corriga as divergências entre estas duas planilhas." _
                        & vbNewLine _
                        & "Todas as importações foram canceladas."
    
        
                        Title = "Valores divergentes"
                             
                        Style = vbOKOnly + vbExclamation
                        Response = MsgBox(Msg, Style, Title)
                        
                        Exit For
                End If
                
            End If
            
        Next iCol

    Next iRow

End Function


Function ImportWorksheets()
    Dim fnameList, fnameCurFile As Variant
    Dim wksCurSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook
    Dim RenamedSheet As String
 
    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Choose Excel files to merge", MultiSelect:=True)
    
    Set WorkSheetNamesImported = CreateObject("System.Collections.ArrayList")
 
    If (vbBoolean <> VarType(fnameList)) Then
 
        If (UBound(fnameList) > 0) Then
 
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
 
            Set wbkCurBook = ActiveWorkbook
 
            For Each fnameCurFile In fnameList
                
                If Flag = False Then
 
                    Set wbkSrcBook = Workbooks.Open(Filename:=fnameCurFile)
    
                    For Each wksCurSheet In wbkSrcBook.Sheets
                        If wksCurSheet.Name = MergingDestination Then
    
                            wksCurSheet.Copy After:=wbkCurBook.Sheets(wbkCurBook.Sheets.Count)
                            
                            RenamedSheet = Left(wbkSrcBook.Name, InStr(wbkSrcBook.Name, ".") - 1)
                            
                            NameSheetImported = RenamedSheet
                            ActiveSheet.Name = RenamedSheet
                            WorkSheetNamesImported.Add RenamedSheet
                            TotalSheetsImported = TotalSheetsImported + 1
                            Worksheets(RenamedSheet).Visible = xlSheetHidden
                            
                            If TotalSheetsImported = 1 Then
                                CopyPasteFirstWorkSheet
                            Else
                                SanitizeWorksheets
                                MergeValues
                            End If
                            
                        End If
                        
                    Next
         
                    wbkSrcBook.Close SaveChanges:=False
                    DeleteImportedSheets
                    
                Else
                    CleanDestinationSheet
                    wbkSrcBook.Close SaveChanges:=False
                    Application.ActiveWorkbook.Close SaveChanges:=False
                End If
 
            Next
 
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
 
        End If
    
    Else
        Flag = True
        
    End If
       
End Function

Function SaveNewWorkbook()

    Dim fname As String
    Dim path As String
    
    path = Application.ActiveWorkbook.path
    fname = "Funil de Vendas - Carteira Guarulhos"
    
    ActiveSheet.Buttons.Delete
    
    Application.DisplayAlerts = False
    Application.ActiveWorkbook.SaveAs Filename:=path & "\" & fname, _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True

End Function

Sub Main()

    LastColumnToMerge = "I"
    LastRowToMerge = 5000
    JumpEmptyRow = 50
    TotalSheetsImported = 0
    Flag = False
    

    MergingDestination = "Base Funil"
    strRangeToCheck = "A1:" & LastColumnToMerge & LastRowToMerge
    
    ImportWorksheets
    
    If Flag = False Then
    
        SaveNewWorkbook
        MsgBox "As planilhas selecionadas foram importadas e fundidas com sucesso! A nova planilha está salva na mesma pasta desta aqui. :)", vbOKOnly + vbInformation
    
    Else
        MsgBox "Infelizmente não foi possível importar as planilhas neste momento. Por favor, verifique se você selecionou alguma planilha ou se não existem dados conflitantes no campo 'VENDEDOR'."
        
    End If
        
    Application.ActiveWorkbook.Close SaveChanges:=False
    
    Application.Quit
    
End Sub


