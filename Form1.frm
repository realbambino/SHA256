VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command 
      Caption         =   "Hash"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "_HASH_"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
Dim cSHA256 As New clsSHA256
    
    Label.Caption = cSHA256.SHA256(Text.Text)
    
    Set cSHA256 = Nothing
End Sub

