VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colors Filter"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicGray 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   6480
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   22
      Top             =   600
      Width           =   1395
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "&Restor Original Picture"
      Height          =   705
      Left            =   240
      TabIndex        =   20
      Top             =   3030
      Width           =   1455
   End
   Begin VB.PictureBox PicRedGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1800
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   15
      Top             =   2280
      Width           =   1395
   End
   Begin VB.PictureBox PicRedBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   3360
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   12
      Top             =   2280
      Width           =   1395
   End
   Begin VB.PictureBox PicGreenBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   4920
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   11
      Top             =   2280
      Width           =   1395
   End
   Begin VB.PictureBox PicNeg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   6480
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   9
      Top             =   2280
      Width           =   1395
   End
   Begin VB.PictureBox PicBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   4920
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   8
      Top             =   600
      Width           =   1395
   End
   Begin VB.PictureBox PicGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   3360
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   7
      Top             =   600
      Width           =   1395
   End
   Begin VB.CommandButton CmdFilter 
      Caption         =   "&Filter Colors"
      Default         =   -1  'True
      Height          =   705
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox PicRed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1800
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   1
      Top             =   600
      Width           =   1395
   End
   Begin VB.PictureBox PicOrg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      Picture         =   "Colors.frx":0000
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   0
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Gray :"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6480
      TabIndex        =   23
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Colors.frx":66AE
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   360
      TabIndex        =   21
      Top             =   4320
      Width           =   7455
   End
   Begin VB.Image OriginalPic 
      Height          =   1455
      Left            =   4200
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5520
      TabIndex        =   19
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Red +"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Red +"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Green +"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Negative:"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Red:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Original Picture:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      X1              =   376
      X2              =   376
      Y1              =   152
      Y2              =   136
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      X1              =   264
      X2              =   376
      Y1              =   152
      Y2              =   136
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   160
      X2              =   264
      Y1              =   136
      Y2              =   152
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000C000&
      X1              =   264
      X2              =   160
      Y1              =   136
      Y2              =   152
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   160
      X2              =   160
      Y1              =   136
      Y2              =   152
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000C000&
      X1              =   264
      X2              =   368
      Y1              =   136
      Y2              =   152
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFE0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   240
      Top             =   4200
      Width           =   7695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   4815
      Left            =   120
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed By: Yehia Muhsen
'Date         : 8-10-2003
'Descriptin   : We know that the main three colors are red, green, and blue
'               and with different combinations of these three colors we can
'               make all colors. Making a specific color is easy because you can
'               use the RGB function, which requires a combination of the main
'               three colors by a number (0-255) for each color. However the
'               hardest part is to filter the main three colors from a specific
'               color. You have to make the RGB function work backward. This
'               program will show you how to do that by some mathematical
'               calculations. This program is going to filter all the main
'               three colors in a picture and redraw that picture with different
'               combinations. This program also shows you how to drag and drop
'               an object on another. After filter the colors, you can drop
'               any picture to the origianl picture, and refilter that picture
'               again. You'll see the presence of some colors and the absence
'               of otheres. This will give you some idea about how colors work.
'               In this program I used two API funtions GetPixel and SetPixel
'               which make the program much faster.
'               Not all ojbect can be dragged and dropped. You have to sent the
'               DragMode property to 1- Automatic. After that write the code for
'               the event of DragDrop for the target object, which is here the
'               PicOrg (The origial picture)


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Sub CmdFilter_Click()
Dim PColor As Long, Blue As Byte, Green As Byte, Red As Byte
Dim Y As Long, X As Long
Dim MaxColor As Integer, MinColor As Integer

'Scan the whole picture from right to left, the up to down
For Y = 0 To PicOrg.ScaleHeight Step 1
    For X = 0 To PicOrg.ScaleWidth Step 1
        
        'Get points' colors
        PColor = GetPixel(PicOrg.hdc, X, Y)

        If PColor >= 0 Then
            
            'Analyze colors
            'The RGB function convert three numbers to one according to this
            'Equation PColor=Red + Green*256 + Blue*256^2
            
            'This following calculation is the opposite of what the funcion RGB does
            Red = PColor Mod 256
            Green = (PColor Mod (256 ^ 2)) \ 256
            Blue = PColor \ (256 ^ 2)
            
            'Redraw picture in differnt color intensities
            
            ' You can use this one for the red color:
            'SetPixel PicRed.hdc, X, Y, Red
            SetPixel PicRed.hdc, X, Y, RGB(Red, 0, 0)
            
            ' You can use this one for the green color:
            'SetPixel PicRed.hdc, X, Y, Green*256
            SetPixel PicGreen.hdc, X, Y, RGB(0, Green, 0)
            
            ' You can use this one for the blue color:
            'SetPixel PicRed.hdc, X, Y, Blue*256^2
            SetPixel PicBlue.hdc, X, Y, RGB(0, 0, Blue)
            
            'You can use this for the red and green colors
            'SetPixel PicRedGreen.hdc, X, Y, Red + Green*256
            SetPixel PicRedGreen.hdc, X, Y, RGB(Red, Green, 0)
            
            'You can use this for the red and blue colors
            'SetPixel PicRedGreen.hdc, X, Y, Red + Blue*256^2
            SetPixel PicRedBlue.hdc, X, Y, RGB(Red, 0, Blue)
            
            'You can use this for the green and blue colors
            'SetPixel PicRedGreen.hdc, X, Y, Green*256 + Blue*256^2
            SetPixel PicGreenBlue.hdc, X, Y, RGB(0, Green, Blue)
            
            'For negative picture
            SetPixel PicNeg.hdc, X, Y, RGB(255 - Red, 255 - Green, 255 - Blue)

            
            'For gray, black, and white picture
            
            'Get the maximum color number
            MaxColor = Red
            If Green > MaxColor Then MaxColor = Green
            If Blue > MaxColor Then MaxColor = Blue
            
            'Get the minimum color number
            MinColor = Red
            If Green < MinColor Then MinColor = Green
            If Blue < MinColor Then MinColor = Blue
            
            'Calculate the average color
            PColor = (MaxColor + MinColor) / 2
            'Show picture
            SetPixel PicGray.hdc, X, Y, RGB(PColor, PColor, PColor)
            
        End If
    Next X
Next Y
End Sub

Private Sub CmdRestore_Click()
'Restore picture
PicOrg.Picture = OriginalPic.Picture

'Filter Colors
CmdFilter_Click
End Sub

Private Sub Form_Load()
'Save a copy of the picture in in a hidden objcet
'So that you can restore it later
OriginalPic.Picture = PicOrg.Picture
End Sub

Private Sub PicOrg_DragDrop(Source As Control, X As Single, Y As Single)

Dim Yy As Long, Xx As Long, PColor As Long

For Yy = 0 To PicOrg.ScaleHeight Step 1
    For Xx = 0 To PicOrg.ScaleWidth Step 1
        'Get point color form the dragged object
        PColor = GetPixel(Source.hdc, Xx, Yy)
        
        'Set a point in the origianl picture object
        SetPixel PicOrg.hdc, Xx, Yy, PColor
    Next Xx
Next Yy

'Filter Colors
CmdFilter_Click
End Sub
