VERSION 5.00
Begin VB.Form Forma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File List"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "Forma.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDir 
      Caption         =   "&Open"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
   Begin VB.CheckBox chkFecha 
      Caption         =   "Date"
      Height          =   200
      Left            =   2020
      TabIndex        =   5
      Top             =   930
      Width           =   855
   End
   Begin VB.CheckBox chkTipo 
      Caption         =   "Type"
      Height          =   200
      Left            =   1060
      TabIndex        =   4
      Top             =   930
      Width           =   855
   End
   Begin VB.CheckBox chkSize 
      Caption         =   "Size"
      Height          =   200
      Left            =   100
      TabIndex        =   3
      Top             =   930
      Width           =   880
   End
   Begin VB.ComboBox cmbExt 
      Height          =   315
      ItemData        =   "Forma.frx":0442
      Left            =   2400
      List            =   "Forma.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Exit"
      Height          =   390
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1865
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Get List"
      Default         =   -1  'True
      Height          =   390
      Left            =   70
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   600
      MaxLength       =   255
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Info"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2880
      TabIndex        =   10
      Top             =   900
      Width           =   1000
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   3960
      Y1              =   1150
      Y2              =   1150
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   3960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line4 
      X1              =   3960
      X2              =   0
      Y1              =   1630
      Y2              =   1630
   End
   Begin VB.Line Line3 
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   1630
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1630
   End
   Begin VB.Label Label3 
      Caption         =   "File Extencion :"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Forma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~¤~ Programmed By Felipe Arteaga webmaster@pitic.net
Dim FSO, Txt, SolExt, Directorio
Private Lista()
Private Sub cmdCerrar_Click()
    Unload Me
    End
End Sub
Private Sub cmdDir_Click()
    Load Dir
    Dir.Show vbModal
End Sub
Private Sub cmdGenerar_Click()
On Error GoTo Control
If Len(Me.txtDir.Text) <= 2 Then
    MsgBox "Directorio no valido", vbExclamation, App.Title
    Exit Sub
End If
Directorio = Me.txtDir.Text
Me.cmbExt.Enabled = False
Me.cmdCerrar.Enabled = False
Me.cmdGenerar.Enabled = False
Me.Label1.Enabled = False
Me.Label3.Enabled = False
Me.chkFecha.Enabled = False
Me.chkSize.Enabled = False
Me.chkTipo.Enabled = False
Tempo = Me.txtDir.Text
Me.txtDir.Text = "Working....."
Me.Refresh
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Txt = FSO.OpenTextFile("C:\Lista.txt", 2, True)
SolExt = LCase(Me.cmbExt.Text)
Archivos (Directorio)
FolderList Directorio
Txt.Close
a = MsgBox("Done, The file C:\Lista.txt has been created " & Chr(13) & "              ¿Do you wish to open the file?", vbQuestion + vbYesNo, "Presto!")
If a = 6 Then Shell "notepad.exe C:\Lista.txt", vbNormalFocus
Me.cmbExt.Enabled = True
Me.cmdCerrar.Enabled = True
Me.cmdGenerar.Enabled = True
Me.Label1.Enabled = True
Me.Label3.Enabled = True
Me.chkFecha.Enabled = True
Me.chkSize.Enabled = True
Me.chkTipo.Enabled = True
Me.txtDir.Text = Tempo
Me.Refresh
Exit Sub
Control:
Select Case Err.Number
    Case 76
        MsgBox "Error" & Chr(13) & _
        "Directorio no valido.", vbCritical, App.Title
    Case Else
        MsgBox "Error" & Chr(13) & _
        "Numero: " & Err.Number & Chr(13) & _
        "Descripcion: " & Err.Description, vbExclamation, App.Title
End Select
Me.cmbExt.Enabled = True
Me.cmdCerrar.Enabled = True
Me.cmdGenerar.Enabled = True
Me.Label1.Enabled = True
Me.Label3.Enabled = True
Me.chkFecha.Enabled = True
Me.chkSize.Enabled = True
Me.chkTipo.Enabled = True
Me.txtDir.Text = Tempo
Me.Refresh
End Sub
Sub Archivos(Direc)
Set fx1 = FSO.GetFolder(Direc)
Set fx2 = fx1.Files
JR = 0
For Each f1 In fx2
    Ext = FSO.GetExtensionName(f1.Path)
    extT = LCase(Ext)
    If extT = SolExt Then
        If JR = 0 Then
            Txt.writeline " "
            Txt.writeline "Folder: " & fx1
            Txt.writeline " "
        End If
        If Me.chkSize.Value = 1 And Me.chkTipo.Value = 1 And Me.chkFecha.Value = 1 Then
            Info = f1.Name & _
            Chr(9) & "Size: " & f1.Size & " Bytes. " & _
            Chr(9) & "Type: " & f1.Type & _
            Chr(9) & "Modify: " & f1.DateLastModified
        ElseIf Me.chkSize.Value = 1 And Me.chkTipo.Value = 1 And Me.chkFecha.Value = 0 Then
            Info = f1.Name & _
            Chr(9) & "Size: " & f1.Size & " Bytes. " & _
            Chr(9) & "Type: " & f1.Type
        ElseIf Me.chkSize.Value = 1 And Me.chkTipo.Value = 0 And Me.chkFecha.Value = 0 Then
            Info = f1.Name & _
            Chr(9) & "Size: " & f1.Size & " Bytes. "
        ElseIf Me.chkSize.Value = 0 And Me.chkTipo.Value = 1 And Me.chkFecha.Value = 0 Then
            Info = f1.Name & _
            Chr(9) & "Type: " & f1.Type
        ElseIf Me.chkSize.Value = 0 And Me.chkTipo.Value = 0 And Me.chkFecha.Value = 1 Then
            Info = f1.Name & _
            Chr(9) & "Modify: " & f1.DateLastModified
        ElseIf Me.chkSize.Value = 0 And Me.chkTipo.Value = 1 And Me.chkFecha.Value = 1 Then
            Info = f1.Name & _
            Chr(9) & "Type: " & f1.Type & _
            Chr(9) & "Modify: " & f1.DateLastModified
        ElseIf Me.chkSize.Value = 0 And Me.chkTipo.Value = 0 And Me.chkFecha.Value = 0 Then
            Info = f1.Name
        End If
        Txt.writeline Info
        Info = ""
        JR = JR + 1
    End If
Next
End Sub
Sub FolderList(Directo)
Dim f, f1, sf
Set f = FSO.GetFolder(Directo)
Set sf = f.SubFolders
For Each f1 In sf
    Archivos (f1.Path)
    FolderList (f1.Path)
Next
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub Form_Load()
With Me.cmbExt
    .AddItem "Mp3": .AddItem "Mp2": .AddItem "Wav": .AddItem "Avi": .AddItem "Mpg": .AddItem "Txt"
    .AddItem "Jpg": .AddItem "Gif": .AddItem "Bmp": .AddItem "Asp": .AddItem "Dbf": .AddItem "Mdb"
    .AddItem "Xls": .AddItem "Doc": .AddItem "Ppt": .AddItem "Dll": .AddItem "Ocx": .AddItem "Vbs"
    .AddItem "Vbp": .AddItem "Htm": .AddItem "Exe": .AddItem "Com": .AddItem "Bat": .AddItem "Sys"
    .ListIndex = 0
End With
Me.txtDir.SelStart = 3
End Sub
Private Sub lblInfo_Click()
MsgBox "                             File List" & Chr(13) & Chr(13) & _
       "~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤~¤" & Chr(13) & _
       "                      webmaster@pitic.net" & Chr(13) & _
       "                       http://www.pitic.net" _
       , vbInformation, "PiticNet"
End Sub
