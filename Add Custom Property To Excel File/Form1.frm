VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " How to add Custom Property to Excel File"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtValue 
      Height          =   330
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Property To Custom Tab In Excel "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Custom Property Value"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   780
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Custom Property Name"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   300
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExcelSheet              As Excel.Application    'Excel Object
Dim ExcelFileProperty       As DocumentProperties   'Excel File Property

Dim Property_Name           As String               'Property Name
Dim Property_Value          As Variant              'Property Value
Dim PropertyNameDataType    As Long                 'Data Type of the property


' ******************************************************************************
' Routine:           PropertyData
' Description:       collect the data of the property from the user
' Created by:        Gil1
' Machine:           GIL
' Date-Time:         4/10/2001-11:03:18 PM
' Last modification: last_modification_info_here
' ******************************************************************************
Private Function PropertyData()
        
On Error GoTo ErrHandler
        
Property_Name = txtName.Text
Property_Value = CStr(txtValue.Text)

PropertyNameDataType = 4 '4 for text type
                         '1 for number type
                         '3 for date type
                         '2 for boolean type
                         
'--- open new Excel File ---
Call MakeExcelFile

'---Add the new property to the custom tab in Excel file ---
Call AddCustomProperty(Property_Name, PropertyNameDataType, Property_Value)

Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

' ******************************************************************************
' Routine:           MakeExcelFile
' Description:       open new excel file
' Created by:        Gil1
' Machine:           GIL
' Date-Time:         4/10/2001-11:05:48 PM
' Last modification: last_modification_info_here
' ******************************************************************************
Public Function MakeExcelFile()

On Error GoTo ErrHandler

    Set ExcelSheet = New Excel.Application
    
    '--- add new workbook and make it visible ---
    ExcelSheet.Workbooks.Add
    ExcelSheet.Visible = True
    

Exit Function
ErrHandler:
    MsgBox Err.Description
End Function
' ******************************************************************************
' Routine:           AddCustomProperty
' Description:       Add the property to the custom tab
' Created by:        Gil1
' Machine:           GIL
' Date-Time:         4/10/2001-11:08:33 PM
' Last modification: last_modification_info_here
' ******************************************************************************
Function AddCustomProperty(PropertyName As String, _
                            PropertyType As Long, _
                            Optional PropertyValue As Variant = "", _
                            Optional bLinkToContent As Boolean = False, _
                            Optional LinkSource As Variant = "")
                                   
On Error GoTo ErrHandler
                                       
    Call DeleteIfExisting(PropertyName)
    
    '--- add the property to the custom tab ---
    '--- if LinkToContent = True then the new custom ---
    '--- property is linked to the location specified by
    '--- LinkSource , a value must be provided ---
    '--- unless the property is linked ---
    
    Set ExcelFileProperty = ExcelSheet.ActiveWorkbook.CustomDocumentProperties. _
        Add(PropertyName, bLinkToContent, PropertyType, PropertyValue, LinkSource)
'    ExcelFileProperty.Add Property_Name, bLinkToContent, PropertyType, PropertyValue, LinkSource

Exit Function
ErrHandler:
    MsgBox Err.Description
End Function
' ******************************************************************************
' Routine:           DeleteIfExisting
' Description:       check if this property already exsit and deleteit
' Created by:        Gil1
' Machine:           GIL
' Date-Time:         4/10/2001-11:07:03 PM
' Last modification: last_modification_info_here
' ******************************************************************************
Function DeleteIfExisting(PropertyName As String)

On Error GoTo ErrHandler

    Dim ExcelFileProperty As DocumentProperty
    
    On Error Resume Next
    
    Set ExcelFileProperty = ActiveWorkbook.CustomDocumentProperties(PropertyName)
    
    If Err.Number = 0 Then
        ExcelFileProperty.Delete
    End If
    
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

Private Sub Command1_Click()

On Error GoTo ErrHandler

If Len(txtName.Text) = 0 Or Len(txtValue.Text) = 0 Then
    MsgBox "You must write the name and value of the property before send it to Excel", vbCritical, "Error"
Else
    Call PropertyData
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub
