VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExtractColumns 
   Caption         =   "Extract Columns"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   OleObjectBlob   =   "frmExtractColumns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExtractColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
'****************************************************************************************
'*  Name        :   frmExtractColumns                                                   *
'*  Authors     :   Joseph Solomon, CPA; Richard Austin; Scott Parrett                  *
'*  Conventions :   Reddick-Gray Naming, RVBA Coding                                    *
'*  Next Steps  :   1)  Add functionality to Column Title Row Text Box                  *
'*                  2)  Code to populate Columns Listbox                                *
'*                  3)  Code to extract selected columns from the workbook into a new   *
'*                          workbook.                                                   *
'*                  4)  Add functionality to Open Workbook Command                      *
'*                  5)  Set Userform properties and convert to xlam (2007+ Addin)       *
'*                                                                                      *
'*  Wish List   :   1)  Loop Worksheets in Workbook                                     *                                                                           *
'*                  2)  Convert Fixed-Width Text Files.                                 *
'*                                                                                      *
'*  Rev Date        Description                                                         *
'*  ~~~ ~~~~~~~~~~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~                           *
'*  1   2012/12/14  Created form and various controls.  Coded open workbooks listbox to *
'*                      update on various Application events, created/updated           *
'*                      documentation.                                                  *
'****************************************************************************************

'*                                                                                      *
'*  Methods                     Scope       Description                                 *
'*  ~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~           *
'*  AddOpenWorkbookstoListbox   Private     Populates passed listbox with open listboxes*
'*                                              from ListOpenWorkbooks function         *
'*  mvarListOpenWorkbooks       Private     Function that returns string array of open  *
'*                                              workbooks on the machine. Function will *
'*                                              return non-array variant on failure     *                                                                                      *
'*
Private Sub AddOpenWorkbookstoListbox( _
ByVal pvlstOpenWbks As MSForms.ListBox)
    'pvlstOpenWbks                      'Listbox Object that will get populated by this
    '                                       subroutine
    Dim wksTarget As Worksheet          'Worksheets to be walked
    Dim varaNameOpenWorkbook As Variant 'Array of open workbooks
    Dim lngA As Long                    'Generic Counter
    
    'Get open workbooks
    varaNameOpenWorkbook = mvarListOpenWorkbooks
    
    'Clear out old values from listbox
    pvlstOpenWbks.Clear
    If Not IsArray(varaNameOpenWorkbook) Then Exit Sub
    
    'Loop array, add values to passed listbox control
    For lngA = LBound(varaNameOpenWorkbook) To UBound(varaNameOpenWorkbook)
        For Each wksTarget In Workbooks(varaNameOpenWorkbook(lngA)).Sheets
            pvlstOpenWbks.AddItem 1, varaNameOpenWorkbook(lngA)
            pvlstOpenWbks.AddItem 2, wksTarget.Name
        Next lngB
    Next lngA
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

Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkboo
End Sub

Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    If Success Then
        AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
    End If
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub App_WorkbookSync(ByVal Wb As Workbook, ByVal SyncEventType As Office.MsoSyncEventType)
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub cmdRefresh_Click()
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub

Private Sub lstWorkbook_Enter()
    AddOpenWorkbookstoListbox pvlstOpenWbks:=Me.lstWorkbook
End Sub
