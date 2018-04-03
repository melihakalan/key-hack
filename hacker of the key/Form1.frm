VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "QuestHack 1.3"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   2805
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4471
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Key"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "a1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "a2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "a3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "a4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "a5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "a6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "a7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "List1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Other"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Command3"
      Tab(1).Control(2)=   "Command4"
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(4)=   "Command6"
      Tab(1).Control(5)=   "Command7"
      Tab(1).Control(6)=   "Command8"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Kar"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command9"
      Tab(2).Control(1)=   "List2"
      Tab(2).Control(2)=   "Command10"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Elm"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command11"
      Tab(3).Control(1)=   "List3"
      Tab(3).Control(2)=   "Command12"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Mrdn"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command13"
      Tab(4).Control(1)=   "List4"
      Tab(4).ControlCount=   2
      Begin VB.ListBox List4 
         Height          =   1620
         ItemData        =   "Form1.frx":008C
         Left            =   -74880
         List            =   "Form1.frx":00A5
         TabIndex        =   28
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command13 
         Caption         =   ">>>"
         Height          =   255
         Left            =   -74160
         TabIndex        =   27
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Go EMC / Luferson / CZ"
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   420
         Width           =   2535
      End
      Begin VB.ListBox List3 
         Height          =   1425
         ItemData        =   "Form1.frx":0123
         Left            =   -74880
         List            =   "Form1.frx":0148
         TabIndex        =   25
         Top             =   780
         Width           =   2535
      End
      Begin VB.CommandButton Command11 
         Caption         =   ">>>"
         Height          =   255
         Left            =   -74160
         TabIndex        =   24
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   ">>>"
         Height          =   255
         Left            =   -74160
         TabIndex        =   23
         Top             =   2220
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   1425
         ItemData        =   "Form1.frx":0231
         Left            =   -74880
         List            =   "Form1.frx":0256
         TabIndex        =   22
         Top             =   780
         Width           =   2535
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Go EMC / Luferson / CZ"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   420
         Width           =   2535
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Captain Falkwine"
         Height          =   255
         Left            =   -74160
         TabIndex        =   20
         Top             =   2100
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Captain Fargo"
         Height          =   255
         Left            =   -74160
         TabIndex        =   19
         Top             =   1860
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Buy Poison Knife"
         Height          =   255
         Left            =   -74160
         TabIndex        =   18
         Top             =   1500
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Blue Potion"
         Height          =   255
         Left            =   -74160
         TabIndex        =   10
         Top             =   1140
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Red Potion"
         Height          =   255
         Left            =   -74160
         TabIndex        =   9
         Top             =   900
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Receive Certificate"
         Height          =   255
         Left            =   -74160
         TabIndex        =   8
         Top             =   420
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   1815
         ItemData        =   "Form1.frx":0340
         Left            =   25
         List            =   "Form1.frx":035F
         TabIndex        =   5
         Top             =   420
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<< Run"
         Height          =   1815
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   885
      End
      Begin VB.Label a7 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   2580
         Width           =   90
      End
      Begin VB.Label a6 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   1320
         TabIndex        =   16
         Top             =   2580
         Width           =   90
      End
      Begin VB.Label a5 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   2580
         Width           =   90
      End
      Begin VB.Label a4 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   2580
         Width           =   90
      End
      Begin VB.Label a3 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Left            =   960
         TabIndex        =   13
         Top             =   2580
         Width           =   90
      End
      Begin VB.Label a2 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   2580
         Width           =   90
      End
      Begin VB.Label a1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   2580
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cleaver:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   7
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "answers:"
         Height          =   195
         Left            =   45
         TabIndex        =   6
         Top             =   2580
         Width           =   660
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Attach"
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Knight OnLine Client"
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "KoJD ~ onlinehile.com, snoxd.net"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   2805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
LoadOffsets
LoadPointers
If AttachKO = False Then
Exit Sub
End If
Me.Show
KO_ADR_CHR = ReadLong(KO_PTR_CHR)
KO_ADR_DLG = ReadLong(KO_PTR_DLG)
FIXSNDFNC
End Sub

Private Sub Command10_Click()
Dim pBytes() As Byte
Dim pStr As String
Select Case List2.ListIndex
Case 0
pStr = "6407a710"
Case 1
pStr = "6407f620"
Case 2
pStr = "6407f820"
Case 3
pStr = "6407fa20"
Case 4
pStr = "64074b23"
Case 5
pStr = "64070f"
Case 6
pStr = "640777"
Case 7
pStr = "64075a2"
Case 8
pStr = "6407612"
Case 9
pStr = "6407642"
Case 10
pStr = "6407c801"
End Select
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command11_Click()
Dim pBytes() As Byte
Dim pStr As String
Select Case List3.ListIndex
Case 0
pStr = "6407a810"
Case 1
pStr = "6407f720"
Case 2
pStr = "6407f920"
Case 3
pStr = "6407fb20"
Case 4
pStr = "64075123"
Case 5
pStr = "640710"
Case 6
pStr = "640774"
Case 7
pStr = "6407bf01"
Case 8
pStr = "6407c601"
Case 9
pStr = "6407ca01"
Case 10
pStr = "6407cb01"
End Select
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command12_Click()
Dim pBytes() As Byte
Dim pStr As String
pStr = "6407bf02"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command13_Click()
Dim pBytes() As Byte
Dim pStr As String
Select Case List4.ListIndex
Case 0
pStr = "6407e826"
Case 1
pStr = "64072d"
Case 2
pStr = "640735"
Case 3
pStr = "640744"
Case 4
pStr = "6407fd"
Case 5
pStr = "64077104"
Case 6
pStr = "64072310"
End Select
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command2_Click()
Dim pBytes() As Byte
Dim pStr As String
Dim nat As Integer

nat = ReadLong(KO_ADR_CHR + nation)
Select Case nat
Case 1
Select Case List1.ListIndex
Case 0
pStr = "64077710"
Case 1
pStr = "64077810"
Case 2
pStr = "64077910"
Case 3
pStr = "64077a10"
Case 4
pStr = "64077b10"
Case 5
pStr = "64077c10"
Case 6
pStr = "64077d10"
Case 7
pStr = "6407d101"
Case 8
pStr = "6407b501"
Case Else
pStr = "FF"
End Select

Case 2

Select Case List1.ListIndex
Case 0
pStr = "64078510"
Case 1
pStr = "64078610"
Case 2
pStr = "64078710"
Case 3
pStr = "64078810"
Case 4
pStr = "64078910"
Case 5
pStr = "64078a10"
Case 6
pStr = "64078b10"
Case 7
pStr = "6407d101"
Case 8
pStr = "6407b501"
Case Else
pStr = "FF"
End Select

End Select

ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()

End Sub

Private Sub Command3_Click()
Dim pBytes() As Byte
Dim pStr As String
pStr = "6407c717"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command4_Click()
Dim pBytes() As Byte
Dim pStr As String
pStr = "64077204"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command5_Click()
Dim pBytes() As Byte
Dim pStr As String
pStr = "64077604"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command6_Click()
Dim pBytes() As Byte
Dim pStr As String
pStr = "64077a04"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command7_Click()
Dim pBytes() As Byte
Dim pStr As String
pStr = "6407d101"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command8_Click()
Dim pBytes() As Byte
Dim pStr As String
pStr = "6407b501"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Command9_Click()
Dim pBytes() As Byte
Dim pStr As String
pStr = "6407be02"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub Form_Load()
If Dir(App.Path & "/load.ini") <> "" Then
Me.Caption = ReadIni(App.Path & "/load.ini", "MAIN", "WINDOW")
Else
MsgBox "could not find load.ini"
End
End If
End Sub

Private Sub List1_Click()

a1.ForeColor = vbBlack
a2.ForeColor = vbBlack
a3.ForeColor = vbBlack
a4.ForeColor = vbBlack
a5.ForeColor = vbBlack
a6.ForeColor = vbBlack
a7.ForeColor = vbBlack

a1.BackColor = &H8000000F
a2.BackColor = &H8000000F
a3.BackColor = &H8000000F
a4.BackColor = &H8000000F
a5.BackColor = &H8000000F
a6.BackColor = &H8000000F
a7.BackColor = &H8000000F

Select Case List1.ListIndex
Case 0
a1.ForeColor = vbYellow
a1.BackColor = vbBlack
Case 1
a2.ForeColor = vbYellow
a2.BackColor = vbBlack
Case 2
a3.ForeColor = vbYellow
a3.BackColor = vbBlack
Case 3
a4.ForeColor = vbYellow
a4.BackColor = vbBlack
Case 4
a5.ForeColor = vbYellow
a5.BackColor = vbBlack
Case 5
a6.ForeColor = vbYellow
a6.BackColor = vbBlack
Case 6
a7.ForeColor = vbYellow
a7.BackColor = vbBlack
End Select

End Sub

Public Sub LoadPointers()
On Error GoTo errr
If Dir(App.Path & "/load.ini") <> "" Then
Dim pointers As String
pointers = ReadIni(App.Path & "/load.ini", "MAIN", "LOADPOINTERS")

KO_PTR_CHR = ReadIni(App.Path & "/load.ini", pointers, "KO_PTR_CHR")
KO_PTR_DLG = ReadIni(App.Path & "/load.ini", pointers, "KO_PTR_DLG")
KO_PTR_PKT = ReadIni(App.Path & "/load.ini", pointers, "KO_PTR_PKT")
KO_SND_FNC = ReadIni(App.Path & "/load.ini", pointers, "KO_SND_FNC")

For i = 1 To 10
If i < 10 Then
SndFnc(i) = ReadIni(App.Path & "/load.ini", "SNDARRAY_" & pointers, "SND0" & i)
ElseIf i = 10 Then
SndFnc(i) = ReadIni(App.Path & "/load.ini", "SNDARRAY_" & pointers, "SND" & i)
End If
Next
Else

MsgBox "could not find load.ini, pointers haven't been updated.(1751)"

End If
Exit Sub
errr:
MsgBox "Could not load pointers!!! error occurred! pointers haven't been updated.(1751)"
LoadOffsets
End Sub

Public Sub FIXSNDFNC()
'1931689751 => &H474A20
'3531683606 => &H474780
WriteLong &HB8AB00, 1931689751
KO_SND_FNC = &H474A20
End Sub
