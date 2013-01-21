VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExtractColumns 
   Caption         =   "Extract Columns"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   OleObjectBlob   =   "frmExtractColumns.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExtractColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1

'****************************************************************************************
'*  Methods                     Scope       Description                                 *
'*  ~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~           *
'*  AddOpenWorkbookstoListbox   Private     Populates passed listbox with open listboxes*
'*                                              from ListOpenWorkbooks function         *
'*  mvarListOpenWorkbooks       Private     Function that returns string array of open  *
'*                                              workbooks on the machine. Function will *
'*                                              return non-array variant on failure     *
'*  AddColumnstoListbox         Private     Adds all columns from workbooks selected to *
'*                                              the columns listbox                     *
'****************************************************************************************
Private Sub AddOpenWorkbookstoListbox( _
ByVal pvlstOpenWbks As MSForms.ListBox)
    'pvlstOpenWbks                      'Listbox Object that will get populated by this
    '                                       subroutine
    Dim wksTarget As Worksheet          'Worksheets to be walked
    Dim uobjWorksheet As clsWorksheet   'Class object holding one user-selected
                                            'instance with a workbook and worksheet
                                            'object.
    Dim uobjWorksheets As clsWorksheets 'Custom collection class holding each user-
                                            'selected wbk/wks instance.
    Dim varaNameOpenWorkbook As Variant 'Array of open workbooks
    Dim lngA As Long                    'Generic Counter
    
    'Store User-Selections from listbox
    Set uobjWorksheets = New clsWorksheets
    uobjWorksheets.FillFromListbox _
        prlstWbkWks:=pvlstOpenWbks, _
        pvlngWBCol:=1, _
        pvlngWSCol:=2
    
    'Get open workbooks
    varaNameOpenWorkbook = mvarListOpenWorkbooks
    
    'Clear out old values from listbox
    pvlstOpenWbks.Clear
    If Not IsArray(varaNameOpenWorkbook) Then Exit Sub
    
    'Loop array, add values to passed listbox control
    For lngA = LBound(varaNameOpenWorkbook) To UBound(varaNameOpenWorkbook)
        For Each wksTarget In Workbooks(varaNameOpenWorkbook(lngA)).Sheets
            With pvlstOpenWbks
                .AddItem
                .Column(1, .ListCount - 1) = varaNameOpenWorkbook(lngA)
                .Column(2, .ListCount - 1) = wksTarget.Name
            End With
        Next wksTarget
    Next lngA
    
    'Reset listbox to user-selected values
    For Each uobjWorksheet In uobjWorksheets
        For lngA = 0 To pvlstOpenWbks.ListCount - 1
            If uobjWorksheet.gstrWbkWksName = _
                pvlstOpenWbks.Column(1, lngA) & "|" & _
                pvlstOpenWbks.Column(2, lngA) Then
                pvlstOpenWbks.Selected(lngA) = True
            End If
        Next lngA
    Next uobjWorksheet
End Sub

Private Function mvarListOpenWorkbooks() As Variant
    Dim wbkTarget As Workbook               'Workbook being walked
    Dim varaNameOpenWorkbooks() As Variant  'Holds names of the open workbooks
    Dim lngCountOpenWorkbooks As Long       'Holds count of open workbooks to set
    '                                           arrays upper bound
    Dim lngA As Long                        'Counter to walk open workbooks
    On Error GoTo mvarListOpenWorkbooks_Err

    'If no workbooks open, return function as false
    If Application.Workbooks.Count = 0 Then
        mvarListOpenWorkbooks = False
        Exit Function
    End If

    'Redimension the array to hold the amount of open workbooks
    ReDim varaNameOpenWorkbooks(1 To Application.Workbooks.Count)

    'Walk the open workbooks and store each name in an array
    lngA = 1
    For Each wbkTarget In Application.Workbooks
         varaNameOpenWorkbooks(lngA) = wbkTarget.Name
         lngA = lngA + 1
    Next wbkTarget
    
    mvarListOpenWorkbooks = varaNameOpenWorkbooks
    
mvarListOpenWorkbooks_Exit:
    Set wbkTarget = Nothing
    Exit Function

mvarListOpenWorkbooks_Err:
    mvarListOpenWorkbooks = False
    GoTo mvarListOpenWorkbooks_Exit
End Function

Private Sub AddColumnstoListbox( _
ByVal pvlstColumns As MSForms.ListBox, _
ByVal pvlstOpenWbks As MSForms.ListBox, _
ByVal pvlngTitleRow As Long)
'pvlstColumns                   'Listbox Control Object to list available columns
'pvlstOpenWbks                  'Listbox Control Object to list open workbooks
'pvlngTitleRow                  'Title Row of Excel spreadsheets
Dim lngLoopCount As Long        'Counter
Dim lngLoopCount2 As Long       'counter
Dim rngTitleRange As Range      'Range defining the column titles for a worksheet
Dim rngTitleCell As Range       'Individual column title cell within column title range
Dim lngCountRowHdrs As Long     'Count of row headers
Dim lngCountWbksSel As Long     'Number of workbooks selected
Dim blnTitleFound As Boolean    'Used to prevent duplicate titles
Dim colTitles As Collection     'Collection to hold column titles
    On Error GoTo AddColumnstoListbox_Err
    pvlstColumns.Clear
    Set colTitles = New Collection
    
    'Count Selections
    For lngLoopCount = 0 To pvlstOpenWbks.ListCount - 1
        If pvlstOpenWbks.Selected(lngLoopCount) Then
            lngCountWbksSel = lngCountWbksSel + 1
        End If
    Next lngLoopCount
    If lngCountWbksSel = 0 Then Exit Sub

    'Look at all column headers in all worksheets, and add uniques to the collection
    lngCountRowHdrs = 0
    For lngLoopCount = 0 To pvlstOpenWbks.ListCount - 1
        If pvlstOpenWbks.Selected(lngLoopCount) Then
            With Workbooks(pvlstOpenWbks.Column(1, lngLoopCount)) _
                .Worksheets(pvlstOpenWbks.Column(2, lngLoopCount))
                
                Set rngTitleRange = Application.Intersect( _
                    .Cells(pvlngTitleRow, 1).EntireRow, .UsedRange)
                    
                For Each rngTitleCell In rngTitleRange
                    For lngLoopCount2 = 1 To colTitles.Count
                        If colTitles.Item(lngLoopCount2) = rngTitleCell.Value Then
                            blnTitleFound = True
                            Exit For
                        End If
                    Next lngLoopCount2
                    If Not blnTitleFound Then colTitles.Add rngTitleCell.Value
                    blnTitleFound = False
                Next rngTitleCell
            End With
        End If
    Next lngLoopCount

    'Add collection values to the Columns Listbox
    For lngLoopCount = 0 To colTitles.Count - 1
        lstExtractColumns.AddItem colTitles.Item(lngLoopCount + 1), lngLoopCount
    Next lngLoopCount
    
AddColumnstoListbox_Exit:

AddColumnstoListbox_Err:
    Select Case Err.Number
        Case Is = 9 'Subscript out of range.  Can trigger if lstWorkbook didn't update
            AddOpenWorkbookstoListbox lstWorkbook
            Resume
    End Select
End Sub

Private Sub app_NewWorkbook(ByVal Wb As Workbook)
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub app_SheetActivate(ByVal Sh As Object)
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub app_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    If Success Then
        AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
    End If
End Sub

Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub cmdExtract_Click()
    ExtractColumns Me.lstWorkbook, Me.lstExtractColumns
End Sub

Private Sub cmdOpenWorkbook_Click()
    Dim lngA As Long
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .FilterIndex = 2
        .Show

        ' Display paths of each file selected
        For lngA = 1 To .SelectedItems.Count
            Application.Workbooks.Open .SelectedItems(lngA)
        Next lngA
    End With
End Sub

Private Sub cmdRefresh_Click()
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub lstWorkbook_AfterUpdate()
    AddColumnstoListbox lstExtractColumns, lstWorkbook, txtTitleRow
End Sub

Private Sub lstWorkbook_Change()
    AddColumnstoListbox lstExtractColumns, lstWorkbook, txtTitleRow
End Sub

Private Sub lstWorkbook_Enter()
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub UserForm_Initialize()
    Set app = ThisWorkbook.Parent
End Sub

