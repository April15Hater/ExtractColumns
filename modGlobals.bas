Attribute VB_Name = "modGlobals"
Option Explicit

'****************************************************************************************
'*  Name        :   modGlobals                                                          *
'*  Authors     :   Joseph Solomon, CPA; Richard Austin; Scott Parrett, MBA             *
'*  Purpose     :   Module to declare Public and Global variables and constants.        *
'*  Last Update :   2013/01/21                                                          *
'****************************************************************************************

Public Const gblncTesting As Boolean = True
Public Const gstrcAPPMENU = "Extract Column Tool"
Public Const glngcHEADERROW As Long = 1         'Header Row constant.  Variable header
                                                'row functionality to be added in the
                                                'future, so constant used to make the
                                                'instances where the Header Row is
                                                'referenced easily to find & search.
Public glngCOLUMNEXTRACT As Long

Public Sub ShowForm()
    If Not frmExtractColumns.Visible Then frmExtractColumns.Show
End Sub
