VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find SNDFNC"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "TRY"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O K"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SNDFNC=0x00474780"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub VScroll1_Change()
Label2.Caption = VScroll1.Value
End Sub

Private Sub Command1_Click()
Me.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Select Case Label2.Caption
Case "0"
KO_SND_FNC = SndFnc(1)
Notice (Hex(SndFnc(1)))
Label2.Caption = "1"
Case "1"
KO_SND_FNC = SndFnc(2)
Notice (Hex(SndFnc(2)))
Label2.Caption = "2"
Case "2"
KO_SND_FNC = SndFnc(3)
Notice (Hex(SndFnc(3)))
Label2.Caption = "3"
Case "3"
KO_SND_FNC = SndFnc(4)
Notice (Hex(SndFnc(4)))
Label2.Caption = "4"
Case "4"
KO_SND_FNC = SndFnc(5)
Notice (Hex(SndFnc(5)))
Label2.Caption = "5"
Case "5"
KO_SND_FNC = SndFnc(6)
Notice (Hex(SndFnc(6)))
Label2.Caption = "6"
Case "6"
KO_SND_FNC = SndFnc(7)
Notice (Hex(SndFnc(7)))
Label2.Caption = "7"
Case "7"
KO_SND_FNC = SndFnc(8)
Notice (Hex(SndFnc(8)))
Label2.Caption = "8"
Case "8"
KO_SND_FNC = SndFnc(9)
Notice (Hex(SndFnc(9)))
Label2.Caption = "9"
Case "9"
KO_SND_FNC = SndFnc(10)
Notice (Hex(SndFnc(10)))
Label2.Caption = "10"
Case Else
End Select
End Sub
