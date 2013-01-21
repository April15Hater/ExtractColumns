Attribute VB_Name = "modExtract"
'****************************************************************************************
'*  Name        :   modExtract                                                          *
'*  Authors     :   Joseph Solomon, CPA                                                 *
'*  Purpose     :   Holds routines neccesary for extraction of the columns.             *
'*  Last Update :   2013/01/21                                                          *
'****************************************************************************************

Option Explicit

Public Sub ExtractColumns( _
ByRef prlstWorksheets As MSForms.ListBox, _
ByRef prlstExtractColumns As MSForms.ListBox)
    Const lngcTargetHeaderRow As Long = 1   'Row to write the headers on the target wks
    Dim lngA As Long                        'Listbox counter variable
    Dim wbkTarget As Workbook               'Workbook parent object to wksTarget
    Dim wksTarget As Worksheet              'Worksheet holding the aggregated columns
    Dim uobjWorksheet As clsWorksheet       'Class object holding one user-selected
                                                'instance with a workbook and worksheet
                                                'object.
    Dim uobjWorksheets As clsWorksheets     'Custom collection class holding each user-
                                                'selected wbk/wks instance.
    Dim colExtractHeaders As Collection     'User-selected column headers to extract
                                                
    'Collect the user-selected lisbox data prior to the new workbook addition. Reasoning
    'is the addition of the target workbook is going to trigger events that are going to
    'clear the selections in the listboxes, hence the listboxes can not be relied upon for
    'storing the user-selected data.
    Set uobjWorksheets = New clsWorksheets
    uobjWorksheets.FillFromListbox _
        prlstWbkWks:=prlstWorksheets, _
        pvlngWBCol:=1, _
        pvlngWSCol:=2
    Set colExtractHeaders = New Collection
    For lngA = 0 To prlstExtractColumns.ListCount - 1
        If prlstExtractColumns.Selected(lngA) Then
            colExtractHeaders.Add prlstExtractColumns.Column(0, lngA), CStr(lngA)
        End If 'prlstExtractColumns.Selected(lngA)
    Next lngA
    
    'Setup and define target wbk/wks headers, format, and properties.
    Set wbkTarget = Workbooks.Add
    For Each wksTarget In wbkTarget.Sheets
        If wbkTarget.Sheets.Count > 1 Then
            'Downsize workbook to 1 worksheet
            Application.DisplayAlerts = False
            wksTarget.Delete
            Application.DisplayAlerts = True
        ElseIf wbkTarget.Sheets.Count = 1 Then
            'Rename and clear any of the contents of the target worksheet
            With wksTarget
                .Name = "Extracted Columns"
                .Cells.ClearContents
                For lngA = 1 To colExtractHeaders.Count
                    With .Cells(lngcTargetHeaderRow, lngA)
                        .Value = colExtractHeaders.Item(lngA)
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                    End With '.Cells(1, lngA)
                Next lngA
            End With 'wksTarget
            Exit For
        End If
    Next wksTarget
    
    'Walk user-selected data, and copy Source Data to Target Worksheet
    For Each uobjWorksheet In uobjWorksheets
        Dim lngSourceColumn As Long
        Dim lngSourceLastRow As Long
        Dim lngTargetColumn As Long
        Dim lngTargetLastRow As Long
        
        'Find last rows of source and target to define ranges
        lngSourceLastRow = uobjWorksheet.gwksWorksheet.UsedRange _
            .Find( _
            What:="*", _
            After:=[A1], _
            SearchDirection:=xlPrevious) _
            .Row
        lngTargetLastRow = wksTarget.UsedRange _
            .Find( _
            What:="*", _
            After:=[A1], _
            SearchDirection:=xlPrevious) _
            .Row
        For lngA = 1 To colExtractHeaders.Count
            Dim strHeaderName As String
            Dim rngSourceColumn As Range
            Dim rngTargetColumn As Range
            
            strHeaderName = colExtractHeaders.Item(lngA)
            
            'Find Source column and define range
            With uobjWorksheet.gwksWorksheet
                lngSourceColumn = .Rows(glngcHEADERROW) _
                    .Find( _
                    What:=strHeaderName, _
                    After:=[A1], _
                    lookat:=xlWhole, _
                    SearchDirection:=xlNext) _
                    .Column
                Set rngSourceColumn = .Range( _
                    .Cells(glngcHEADERROW + 1, lngSourceColumn), _
                    .Cells(lngSourceLastRow, lngSourceColumn))
            End With 'uobjWorksheet.gwksWorksheet
            
            'Find Target column and define range
            With wksTarget
                lngTargetColumn = .Rows(lngcTargetHeaderRow) _
                    .Find( _
                    What:=strHeaderName, _
                    After:=[A1], _
                    lookat:=xlWhole, _
                    SearchDirection:=xlNext) _
                    .Column
                Set rngTargetColumn = .Range( _
                    .Cells(lngTargetLastRow + 1, lngTargetColumn), _
                    .Cells(lngTargetLastRow + rngSourceColumn.Count, lngTargetColumn))
            End With 'wksTarget
            
            'Copy Data
            rngTargetColumn.Value = rngSourceColumn.Value
        Next lngA
    Next uobjWorksheet
    
    'Reset listbox to user-selected values
    For Each uobjWorksheet In uobjWorksheets
        For lngA = 0 To prlstWorksheets.ListCount - 1
            If uobjWorksheet.gstrWbkWksName = _
                prlstWorksheets.Column(1, lngA) & "|" & _
                prlstWorksheets.Column(2, lngA) Then
                prlstWorksheets.Selected(lngA) = True
            End If
        Next lngA
    Next uobjWorksheet
End Sub


