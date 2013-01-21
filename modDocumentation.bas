Attribute VB_Name = "modDocumentation"
'****************************************************************************************
'*  Project Name:   Extract Columns                                                     *
'*  Authors     :   Joseph Solomon, CPA; Richard Austin; Scott Parrett, MBA             *
'*  Conventions :   Reddick-Gray Naming, RVBA Coding                                    *
'*  Next Steps  :   1)  Set Userform properties and convert to xlam (2007+ Addin)       *
'*                  2)  Refresh for new worksheets                                      *
'*                                                                                      *
'*  Wish List   :   1)  Convert Fixed-Width Text Files.                                 *
'*                  2)  Replace listbox wbk/wks information with custom collection      *
'*                                                                                      *
'*  Modules             Description                                                     *
'*  ~~~~~~~~            ~~~~~~~~~~~                                                     *
'*  modExtract          Holds routines neccesary for extraction of the columns.         *
'*  modGlobals          Module to declare Public and Global variables and constants.    *
'*  modDocumentation    This Module; entirely comments holding Application-Wide info    *
'*  clsWorksheet        Object that holds workbook and worksheet objects as properties  *
'*                          and a string property of the wbk/wks name separated by a    *
'*                          "|" character.                                              *
'*  clsWorksheets       Custom collection of wbk/wks objects.  Wraps clsWorksheet.      *
'*  frmExtractColumn    Main form used to drive user-selections.                        *
'*                                                                                      *
'*  Rev Date        Description                                                         *
'*  ~~~ ~~~~~~~~~~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~                           *
'*  1   2012/12/14  Created form and various controls.  Coded open workbooks listbox to *
'*                      update on various Application events, created/updated           *
'*                      documentation.                                                  *
'*  2   2012/12/21  Created AddColumnstoListbox method.  Fixed bug for populating the   *
'*                      lstWorkbook listbox.  Added event for change on Column List.    *
'*                      changed workbooks and column lists to multiselect.              *
'*  3   2013/01/21  Created custom collection class for workbook/worksheet objects;     *
'*                      built column extraction routine; added Header Row constant for  *
'*                      easier implementation of variable header row functionality;     *
'*                      programmed Open Workbook command button.                        *
'*                  New Modules: modExtract, clsWorksheet, clsWorksheets,               *
'*                      modDocumentation                                                *
'****************************************************************************************

