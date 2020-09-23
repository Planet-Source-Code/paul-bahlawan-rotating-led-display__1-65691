VERSION 5.00
Begin VB.Form frmRotLED 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Rotating LED display"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrRotate 
      Interval        =   40
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   600
      MaxLength       =   18
      TabIndex        =   0
      Text            =   "Your Message Here"
      Top             =   2160
      Width           =   2655
   End
End
Attribute VB_Name = "frmRotLED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Rotating LED Display
'' Paul Bahlawan
'' June 15, 2006
''
'' Built on Simon Lynn's 3D Rotating DNA project:
'' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=13973
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Ratio As Single, Gap As Single, size As Single
Private Dfont(479) As Byte          'Dot Matrix font will be stored here
Private message As String
Private interAng As Single
Const DEG = 1.74532925199433E-02    '(pi/180) for converting degrees to radians

Private Sub DrawLED()
Dim Ang As Single
Dim i As Long
Dim j As Long
Dim colr As Long
Dim chrPos As Long
    
    Cls
    'The LEDs are drawn from back to front
    For j = 0 To 179 Step 4
        
        'Fade LEDS toward back
        colr = RGB(100 + j, 0, 0)
              
        'Draw the column of LEDS (left half)
        chrPos = (Asc(Mid$(message, Int(j / 4 / 5) + 1, 1)) - 32) * 5 + (j / 4 Mod 5)
        Ang = interAng + 180 - j
        For i = 0 To 6 '7 LEDS = 7 bits in font...
            If Dfont(chrPos) And (2 ^ i) Then
                PlotOrbit Ang, Width / 2, 600 + (i * Gap), size, Ratio + i * 0.3, colr
            End If
        Next i
        
        'Draw the column of LEDS (right half)
        chrPos = (Asc(Mid$(message, 18 - Int(j / 4 / 5), 1)) - 32) * 5 + (4 - (j / 4 Mod 5))
        Ang = interAng + 184 + j
        For i = 0 To 6
            If Dfont(chrPos) And (2 ^ i) Then
                PlotOrbit Ang, Width / 2, 600 + (i * Gap), size, Ratio + i * 0.3, colr
            End If
        Next i
    Next j
End Sub

Private Sub PlotOrbit(Angle As Single, CntX As Single, CntY As Single, Radius As Single, Ratio As Single, Colour As Long)
Dim PntX1 As Single, PntY1 As Single
    PntX1 = CntX - (Sin(Angle * DEG) * Radius)
    PntY1 = CntY - (Cos(Angle * DEG) * (Radius / Ratio))
    PSet (PntX1, PntY1), Colour
End Sub

Private Sub Form_Load()
    'Change these variables to change the appearance
    Ratio = -7      ' The apparent viewing angle
    Gap = 100       ' The virtical distance between LEDS
    size = 1700     ' The Radius of circle
    Me.DrawWidth = 5 ' Dot (LED) size
    
    'Load T Jackson's font in a byte array
    Open "dotmatrix.font" For Binary Access Read As #1
         Get #1, , Dfont()
    Close #1
    
    txtMessage_Change
End Sub

Private Sub txtMessage_Change()
    message = Left$(txtMessage.Text & "                  ", 18)
End Sub

Private Sub Form_Click()
    TmrRotate.Enabled = Not TmrRotate.Enabled
End Sub

Private Sub TmrRotate_Timer()
    interAng = interAng + 4
    If interAng = 20 Then
        message = Right$(message, 17) & Left$(message, 1)
        interAng = 0
    End If
    DrawLED
End Sub
