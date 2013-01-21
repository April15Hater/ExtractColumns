VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorksheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolWorksheets As Collection

Private Sub Class_Initialize()
    Set mcolWorksheets = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolWorksheets = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolWorksheets.[_NewEnum]
End Property

Public Sub Add( _
ByVal pvclsWorksheet As clsWorksheet, _
ByVal pvstrWbkWks As String)
    mcolWorksheets.Add pvclsWorksheet, pvclsWorksheet.gstrWbkWksName
End Sub

Public Sub Remove(Index As Variant)
    mcolWorksheets.Remove Index
End Sub

Public Property Get Item(Index As Variant) As clsWorksheet
Attribute Item.VB_UserMemId = 0
    Set Item = mcolWorksheets.Item(Index)
End Property

Property Get Count() As Long
    Count = mcolWorksheets.Count
End Property

Public Sub FillFromListbox( _
ByRef prlstWbkWks As MSForms.ListBox, _
ByVal pvlngWBCol As Long, _
ByVal pvlngWSCol As Long)
    Dim lngA As Long
    Dim uobjWorksheet As clsWorksheet
    For lngA = 0 To prlstWbkWks.ListCount - 1
        Set uobjWorksheet = New clsWorksheet
        If prlstWbkWks.Selected(lngA) Then
            Dim wbkSource As Workbook
            Dim wksSource As Worksheet
            Set uobjWorksheet = New clsWorksheet
            Set wbkSource = Workbooks(prlstWbkWks.Column(pvlngWBCol, lngA))
            Set wksSource = wbkSource _
                .Worksheets(prlstWbkWks.Column(pvlngWSCol, lngA))
            Set uobjWorksheet.gwbkWorkbook = wbkSource
            Set uobjWorksheet.gwksWorksheet = wksSource
            Me.Add uobjWorksheet, uobjWorksheet.gstrWbkWksName
        End If 'prlstWbkWks.Selected(lngA)
        Set uobjWorksheet = Nothing
    Next lngA
End Sub
