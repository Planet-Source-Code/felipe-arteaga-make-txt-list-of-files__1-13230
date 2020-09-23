VERSION 5.00
Begin VB.Form Dir 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "Dir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDir 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   80
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "Dir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~¤~ Programmed By Felipe Arteaga webmaster@pitic.net
Private Sub cmdAceptar_Click()
    Forma.txtDir.Text = Dir1.Path
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Dir1_Change()
    txtDir.Text = Dir1.Path
    Me.Caption = Dir1.Path
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub Form_Load()
    Dir1.Path = "C:\"
    Drive1.Drive = "C:\"
    Me.Caption = Dir1.Path
End Sub
