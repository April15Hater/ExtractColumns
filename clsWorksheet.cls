VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mwbkWorkbook As Workbook
Private WithEvents mwksWorksheet As Worksheet

Public Property Get gwbkWorkbook() As Workbook
    Set gwbkWorkbook = mwbkWorkbook
End Property

Public Property Set gwbkWorkbook(ByVal pvwbkWorkbook As Workbook)
    Set mwbkWorkbook = pvwbkWorkbook
End Property

Public Property Get gwksWorksheet() As Worksheet
    Set gwksWorksheet = mwksWorksheet
End Property

Public Property Set gwksWorksheet(ByVal pvwksWorksheet As Worksheet)
    Set mwksWorksheet = pvwksWorksheet
End Property

Public Property Get gstrWbkWksName() As Variant
    gstrWbkWksName = mwbkWorkbook.Name & "|" & mwksWorksheet.Name
End Property

