VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cm1 
      Left            =   8880
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Map"
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   9120
      Width           =   1095
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   2
      Left            =   8520
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   3
      Top             =   2640
      Width           =   780
      Visible         =   0   'False
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   1
      Left            =   8520
      Picture         =   "Form1.frx":1B42
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   2
      Top             =   1920
      Width           =   780
      Visible         =   0   'False
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   0
      Left            =   8520
      Picture         =   "Form1.frx":3684
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   1
      Top             =   1200
      Width           =   780
      Visible         =   0   'False
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   120
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tiles(100) As String, LineRead(10) As String
Dim pX As Integer, pY As Integer, Border As Integer

Private Sub cmdLoad_Click()
Border = 48 * 9
cm1.Filter = "Map/Text files |*.txt|"
cm1.Action = 1
If cm1.Filename = "" Then Exit Sub
LoadMap cm1.Filename
End Sub

Private Sub Form_Load()
pX = 0
pY = 0

End Sub
Function LoadMap(Filename As String)
pX = 0
pY = 0
picMap.Cls
Dim x As Integer, a As Integer

    Open Filename For Input As #1
        Do Until EOF(1)
        x = x + 1
        Input #1, LineRead(x)
        Loop
    Close #1
    
    Tiles(0) = Split(LineRead(1), "|")(1)
    Tiles(1) = Split(LineRead(1), "|")(2)
    Tiles(2) = Split(LineRead(1), "|")(3)
    Tiles(3) = Split(LineRead(1), "|")(4)
    Tiles(4) = Split(LineRead(1), "|")(5)
    Tiles(5) = Split(LineRead(1), "|")(6)
    Tiles(6) = Split(LineRead(1), "|")(7)
    Tiles(7) = Split(LineRead(1), "|")(8)
    Tiles(8) = Split(LineRead(1), "|")(9)
    Tiles(9) = Split(LineRead(1), "|")(10)
    Tiles(10) = Split(LineRead(2), "|")(1)
    Tiles(11) = Split(LineRead(2), "|")(2)
    Tiles(12) = Split(LineRead(2), "|")(3)
    Tiles(13) = Split(LineRead(2), "|")(4)
    Tiles(14) = Split(LineRead(2), "|")(5)
    Tiles(15) = Split(LineRead(2), "|")(6)
    Tiles(16) = Split(LineRead(2), "|")(7)
    Tiles(17) = Split(LineRead(2), "|")(8)
    Tiles(18) = Split(LineRead(2), "|")(9)
    Tiles(19) = Split(LineRead(2), "|")(10)
    Tiles(20) = Split(LineRead(3), "|")(1)
    Tiles(21) = Split(LineRead(3), "|")(2)
    Tiles(22) = Split(LineRead(3), "|")(3)
    Tiles(23) = Split(LineRead(3), "|")(4)
    Tiles(24) = Split(LineRead(3), "|")(5)
    Tiles(25) = Split(LineRead(3), "|")(6)
    Tiles(26) = Split(LineRead(3), "|")(7)
    Tiles(27) = Split(LineRead(3), "|")(8)
    Tiles(28) = Split(LineRead(3), "|")(9)
    Tiles(29) = Split(LineRead(3), "|")(10)
    Tiles(30) = Split(LineRead(4), "|")(1)
    Tiles(31) = Split(LineRead(4), "|")(2)
    Tiles(32) = Split(LineRead(4), "|")(3)
    Tiles(33) = Split(LineRead(4), "|")(4)
    Tiles(34) = Split(LineRead(4), "|")(5)
    Tiles(35) = Split(LineRead(4), "|")(6)
    Tiles(36) = Split(LineRead(4), "|")(7)
    Tiles(37) = Split(LineRead(4), "|")(8)
    Tiles(38) = Split(LineRead(4), "|")(9)
    Tiles(39) = Split(LineRead(5), "|")(10)
    Tiles(40) = Split(LineRead(5), "|")(1)
    Tiles(41) = Split(LineRead(5), "|")(2)
    Tiles(42) = Split(LineRead(5), "|")(3)
    Tiles(43) = Split(LineRead(5), "|")(4)
    Tiles(44) = Split(LineRead(5), "|")(5)
    Tiles(45) = Split(LineRead(5), "|")(6)
    Tiles(46) = Split(LineRead(5), "|")(7)
    Tiles(47) = Split(LineRead(5), "|")(8)
    Tiles(48) = Split(LineRead(5), "|")(9)
    Tiles(49) = Split(LineRead(5), "|")(10)
    Tiles(50) = Split(LineRead(6), "|")(1)
    Tiles(51) = Split(LineRead(6), "|")(2)
    Tiles(52) = Split(LineRead(6), "|")(3)
    Tiles(53) = Split(LineRead(6), "|")(4)
    Tiles(54) = Split(LineRead(6), "|")(5)
    Tiles(55) = Split(LineRead(6), "|")(6)
    Tiles(56) = Split(LineRead(6), "|")(7)
    Tiles(57) = Split(LineRead(6), "|")(8)
    Tiles(58) = Split(LineRead(6), "|")(9)
    Tiles(59) = Split(LineRead(6), "|")(10)
    Tiles(60) = Split(LineRead(7), "|")(1)
    Tiles(61) = Split(LineRead(7), "|")(2)
    Tiles(62) = Split(LineRead(7), "|")(3)
    Tiles(63) = Split(LineRead(7), "|")(4)
    Tiles(64) = Split(LineRead(7), "|")(5)
    Tiles(65) = Split(LineRead(7), "|")(6)
    Tiles(66) = Split(LineRead(7), "|")(7)
    Tiles(67) = Split(LineRead(7), "|")(8)
    Tiles(68) = Split(LineRead(7), "|")(9)
    Tiles(69) = Split(LineRead(7), "|")(10)
    Tiles(70) = Split(LineRead(8), "|")(1)
    Tiles(71) = Split(LineRead(8), "|")(2)
    Tiles(72) = Split(LineRead(8), "|")(3)
    Tiles(73) = Split(LineRead(8), "|")(4)
    Tiles(74) = Split(LineRead(8), "|")(5)
    Tiles(75) = Split(LineRead(8), "|")(6)
    Tiles(76) = Split(LineRead(8), "|")(7)
    Tiles(77) = Split(LineRead(8), "|")(8)
    Tiles(78) = Split(LineRead(8), "|")(9)
    Tiles(79) = Split(LineRead(8), "|")(10)
    Tiles(80) = Split(LineRead(9), "|")(1)
    Tiles(81) = Split(LineRead(9), "|")(2)
    Tiles(82) = Split(LineRead(9), "|")(3)
    Tiles(83) = Split(LineRead(9), "|")(4)
    Tiles(84) = Split(LineRead(9), "|")(5)
    Tiles(85) = Split(LineRead(9), "|")(6)
    Tiles(86) = Split(LineRead(9), "|")(7)
    Tiles(87) = Split(LineRead(9), "|")(8)
    Tiles(88) = Split(LineRead(9), "|")(9)
    Tiles(89) = Split(LineRead(9), "|")(10)
    Tiles(90) = Split(LineRead(10), "|")(1)
    Tiles(91) = Split(LineRead(10), "|")(2)
    Tiles(92) = Split(LineRead(10), "|")(3)
    Tiles(93) = Split(LineRead(10), "|")(4)
    Tiles(94) = Split(LineRead(10), "|")(5)
    Tiles(95) = Split(LineRead(10), "|")(6)
    Tiles(96) = Split(LineRead(10), "|")(7)
    Tiles(97) = Split(LineRead(10), "|")(8)
    Tiles(98) = Split(LineRead(10), "|")(9)
    Tiles(99) = Split(LineRead(10), "|")(10)
    
        While a <= 99
        
    If Tiles(a) = "g" Then Call BitBlt(picMap.hDC, pX, pY, 780, 780, picTiles(2).hDC, 0, 0, SRCCOPY)
    If Tiles(a) = "d" Then Call BitBlt(picMap.hDC, pX, pY, 780, 780, picTiles(0).hDC, 0, 0, SRCCOPY)
    If Tiles(a) = "w" Then Call BitBlt(picMap.hDC, pX, pY, 780, 780, picTiles(1).hDC, 0, 0, SRCCOPY)
    If Tiles(a) = "b" Then Call BitBlt(picMap.hDC, pX, pY, 780, 780, picTiles(1).hDC, 0, 0, BLACKNESS)

            picMap.Refresh
            If pX >= Border Then
        pX = 0
        pY = pY + picTiles(0).ScaleHeight
        Else
    pX = pX + picTiles(0).ScaleWidth

    End If
    a = a + 1
    Wend
    
End Function
Function Add(Text As String)
Text1.SelStart = Len(Text1.Text)
Text1.Text = Text1.Text & Text & vbCrLf
Text1.SelStart = Len(Text1.Text)
End Function
