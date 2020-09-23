VERSION 5.00
Begin VB.Form frm_routeprint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Route Print By Erwan L."
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FRM_RO~1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Copy To Clipboard"
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   120
      Picture         =   "FRM_RO~1.frx":0260
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   5160
      Picture         =   "FRM_RO~1.frx":056A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Enum"
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "frm_routeprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Route Print
'Author : Erwan L.
'mail:erwan.l@free.fr
'use iphlpapi (may be this can be done with winsock)
'use loadlibrary to check iphlpapi
'dynamic load of listview
'copy to clipboard function

'

Option Explicit
Public WithEvents listview1  As VBControlExtender
Attribute listview1.VB_VarHelpID = -1

Private Sub Command1_Click()
  '
    Dim hDll As Long
    Dim hMod As Long
'
    hDll = LoadLibrary("iphlpapi.dll")
    If hDll <> 0 Then
        hMod = GetProcAddress(hDll, "GetTcpTable")
        If hMod = 0 Then
            MsgBox "could not getprocaddress of GetTcpTable" & vbCrLf & _
            Err.LastDllError
            FreeLibrary hDll
            Exit Sub
        End If
    Else
        MsgBox "could not load iphlpapi library" & vbCrLf & _
        Err.LastDllError
        Exit Sub
    End If
    

Dim infos() As String
ReDim infos(0)
routeprint infos()
Dim j As Integer
Dim tblinfos
Dim lvitem
listview1.object.listitems.Clear
For j = 1 To UBound(infos)
    Err.Clear
    tblinfos = Split(infos(j), vbTab)
    'il y a un ou plsrs connectes
    If Err.Number = 0 Then
    Set lvitem = listview1.object.listitems.Add(, , tblinfos(0))
                    lvitem.subitems(1) = tblinfos(1)
                    lvitem.subitems(2) = tblinfos(2)
                    lvitem.subitems(3) = tblinfos(3)
    End If
    Next j
End Sub

Private Sub Command3_Click()
Dim j As Integer
Dim i As Integer
Dim str As String
Dim lvitem
For i = 1 To listview1.object.ColumnHeaders.Count
    str = str & listview1.object.ColumnHeaders(i) & ";"
Next i
str = str & vbCrLf
For j = 1 To listview1.object.listitems.Count
    Set lvitem = listview1.object.listitems(j)
    str = str & lvitem & ";"
        For i = 1 To listview1.object.ColumnHeaders.Count - 1
            str = str & lvitem.subitems(i) & ";"
        Next i
    str = str & vbCrLf
    Next j
Clipboard.SetText str
str = ""
End Sub

Private Sub Form_Activate()
Debug.Print "Form_Activate"
Debug.Print Me.Width & ":" & Me.Height
End Sub

Private Sub Form_GotFocus()
Debug.Print "Form_GotFocus"
Debug.Print Me.Width & ":" & Me.Height
End Sub

Private Sub Form_Initialize()
Debug.Print "Form_Initialize"
Debug.Print Me.Width & ":" & Me.Height
End Sub

Private Sub Form_Load()
Debug.Print "Form_Load"
Debug.Print Me.Width & ":" & Me.Height
'
On Error Resume Next
Licenses.Add "MSComctlLib.listviewctrl"
Err.Clear
Set Me.listview1 = Me.Controls.Add("MSComctlLib.listviewctrl", "listview1", Me)
If Err <> 0 Then
    MsgBox Err.Description & vbCrLf & "cant use MSCOMCTL.OCX : MSComctlLib.listviewctrl"
    Unload Me
    End If
On Error GoTo 0

Me.listview1.Top = 10
Me.listview1.Width = 5850
Me.listview1.Height = 2535
Me.listview1.Left = 120
Me.listview1.Visible = True
'
listview1.object.ColumnHeaders.Add 1, , "Network Dest."
listview1.object.ColumnHeaders.Add 2, , "Mask"
listview1.object.ColumnHeaders.Add 3, , "Gateway"
listview1.object.ColumnHeaders.Add 4, , "Interface"


listview1.object.View = 3 'lvwreport
'

End Sub

Private Sub Form_Resize()
Debug.Print "Form_Resize"
Debug.Print Me.Width & ":" & Me.Height
End Sub
