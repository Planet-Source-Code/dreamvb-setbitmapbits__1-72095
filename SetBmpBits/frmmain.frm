VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SetBitmap Bits"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   4890
      TabIndex        =   6
      Top             =   2610
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoise 
      Caption         =   "Noise+30"
      Height          =   390
      Left            =   4890
      TabIndex        =   5
      Top             =   2130
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dark-10"
      Height          =   390
      Left            =   4890
      TabIndex        =   4
      Top             =   1650
      Width           =   1215
   End
   Begin VB.CommandButton cmdLight 
      Caption         =   "Light+10"
      Height          =   390
      Left            =   4890
      TabIndex        =   3
      Top             =   1155
      Width           =   1215
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert"
      Height          =   390
      Left            =   4890
      TabIndex        =   2
      Top             =   675
      Width           =   1215
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      Height          =   390
      Left            =   4890
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.PictureBox pDst 
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   165
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   210
      Width           =   4500
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Bitmap
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private bmpBits() As Byte
Private hBmp As Bitmap

Private Sub GetBits(pBox As PictureBox)
Dim iRet As Long
    'Get the bitmap header
    iRet = GetObject(pBox.Picture.Handle, Len(hBmp), hBmp)
    'Resize to hold image data
    ReDim bmpBits(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into bmpBits array
    iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, bmpBits(0, 0, 0))
End Sub

Private Sub SetBits(pBox As PictureBox)
Dim iRet As Long
    'Set the new image data back onto pBox
    iRet = SetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, bmpBits(0, 0, 0))
    'Erase bmpBits because we finished with it now
    Erase bmpBits
End Sub

Private Sub Noise(ByVal Value As Integer)
Dim X As Long
Dim Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim iRnd As Integer

    Call GetBits(pDst)
    'Now we can play with the image data
    For X = 0 To hBmp.bmWidth - 1
        For Y = 0 To hBmp.bmHeight - 1
            'Noise
            iRnd = Int(Rnd * Value)
            'Invert the colors
            R = bmpBits(0, X, Y) + iRnd 'Red
            G = bmpBits(1, X, Y) + iRnd 'Green
            B = bmpBits(2, X, Y) + iRnd 'Blue
            
            If (R < 0) Then R = 0
            If (G < 0) Then G = 0
            If (B < 0) Then B = 0
            If (R > 255) Then R = 255
            If (G > 255) Then G = 255
            If (B > 255) Then B = 255
            
            'Set colors
            bmpBits(0, X, Y) = R
            bmpBits(1, X, Y) = G
            bmpBits(2, X, Y) = B
        Next Y
    Next X
    'Here we set the new bits
    Call SetBits(pDst)
    'And Refresh the picturebox
    Call pDst.Refresh
End Sub

Private Sub Invert()
Dim X As Long
Dim Y As Long
    Call GetBits(pDst)
    'Now we can play with the image data
    For X = 0 To hBmp.bmWidth - 1
        For Y = 0 To hBmp.bmHeight - 1
            'Invert the colors
            bmpBits(0, X, Y) = 255 - bmpBits(0, X, Y) 'Red
            bmpBits(1, X, Y) = 255 - bmpBits(1, X, Y) 'Green
            bmpBits(2, X, Y) = 255 - bmpBits(2, X, Y) 'Blue
        Next Y
    Next X
    'Here we set the new bits
    Call SetBits(pDst)
    'And Refresh the picturebox
    Call pDst.Refresh
End Sub

Private Sub LightDark(ByVal Value As Integer)
Dim X As Long
Dim Y As Long
Dim R As Integer, G As Integer, B As Integer
    Call GetBits(pDst)
    'Now we can play with the image data
    For X = 0 To hBmp.bmWidth - 1
        For Y = 0 To hBmp.bmHeight - 1
            'Invert the colors
            R = bmpBits(0, X, Y) + Value 'Red
            G = bmpBits(1, X, Y) + Value 'Green
            B = bmpBits(2, X, Y) + Value 'Blue
            
            If (R < 0) Then R = 0
            If (G < 0) Then G = 0
            If (B < 0) Then B = 0
            If (R > 255) Then R = 255
            If (G > 255) Then G = 255
            If (B > 255) Then B = 255
            
            'Set colors
            bmpBits(0, X, Y) = R
            bmpBits(1, X, Y) = G
            bmpBits(2, X, Y) = B
        Next Y
    Next X
    'Here we set the new bits
    Call SetBits(pDst)
    'And Refresh the picturebox
    Call pDst.Refresh
End Sub

Private Sub cmdExit_Click()
    Unload frmmain
End Sub

Private Sub cmdInvert_Click()
    Call Invert
End Sub

Private Sub cmdLight_Click()
    Call LightDark(10)
End Sub

Private Sub cmdNoise_Click()
    Call Noise(30)
End Sub

Private Sub cmdRestore_Click()
    pDst.Picture = LoadPicture(App.Path & "\catty.bmp")
End Sub

Private Sub Command1_Click()
    Call LightDark(-10)
End Sub

Private Sub Form_Load()
    Call cmdRestore_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub
