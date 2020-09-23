VERSION 5.00
Begin VB.UserControl IndForm 
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   870
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   795
   ScaleWidth      =   870
   ToolboxBitmap   =   "frmIndividual.ctx":0000
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   0
      Picture         =   "frmIndividual.ctx":0312
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "IndForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "Gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "Gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "Gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "Gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteDC Lib "Gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "Gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function BeginPath Lib "Gdi32" (ByVal hDC As Long) As Long
Private Declare Function TextOut Lib "Gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function EndPath Lib "Gdi32" (ByVal hDC As Long) As Long
Private Declare Function PathToRegion Lib "Gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetRgnBox Lib "Gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CreateRectRgnIndirect Lib "Gdi32" (lpRect As RECT) As Long
Private Declare Function CombineRgn Lib "Gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "Gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const HTCAPTION = 2
    Private Const RGN_AND = 1

Private Function GetTextRgn(Font As String, Size As Integer, Text As String) As Long
    Dim moform As Form
    Set moform = UserControl.Parent
    moform.Font = Font
    moform.FontSize = Size
    Dim hRgn1 As Long, hRgn2 As Long
    Dim rct As RECT
    BeginPath moform.hDC
    TextOut moform.hDC, 10, 10, Text, Len(Text)
    EndPath moform.hDC
    hRgn1 = PathToRegion(moform.hDC)
    GetRgnBox hRgn1, rct
    hRgn2 = CreateRectRgnIndirect(rct)
    CombineRgn hRgn2, hRgn2, hRgn1, RGN_AND
    DeleteObject hRgn1
    GetTextRgn = hRgn2
    SetWindowRgn moform.hWnd, hRgn, 1
End Function

Private Sub UserControl_Initialize()
    UserControl.Width = Command1.Width
    UserControl.Height = Command1.Height
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Command1.Width
    UserControl.Height = Command1.Height
End Sub

Function FormText(Font As String, FontSize As Integer, Text As String)
    Dim moform As Form
    Set moform = UserControl.Parent
    Dim hRgn As Long
    hRgn = GetTextRgn(Font, FontSize, Text)
    SetWindowRgn moform.hWnd, hRgn, 1
    Dim Iter As Long
    Const Banding = 8

    For Iter = 0 To ScaleHeight Step Banding
        moform.Line (0, Iter)-(moform.ScaleWidth, Iter + Banding), , BF
    Next
End Function

Function Bitmap2Form(Picture As StdPicture, FilterColor As Long)
Dim moform As Form
Set moform = UserControl.Parent
Dim hRgn As Long, tRgn As Long
Dim x As Integer, y As Integer, X0 As Integer
Dim hDC As Long, BM As BITMAP
If hRgn Then DeleteObject hRgn
hDC = CreateCompatibleDC(0)
If hDC Then
    SelectObject hDC, Picture
    GetObject Picture, Len(BM), BM
    hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
    For y = 0 To BM.bmHeight
        For x = 0 To BM.bmWidth
            While x <= BM.bmWidth And GetPixel(hDC, x, y) <> FilterColor
                x = x + 1
            Wend
            X0 = x
            While x <= BM.bmWidth And GetPixel(hDC, x, y) = FilterColor
                x = x + 1
            Wend
            If X0 < x Then
                tRgn = CreateRectRgn(X0, y, x, y + 1)
                CombineRgn hRgn, hRgn, tRgn, 4
                DeleteObject tRgn
            End If
        Next x
    Next y
    Bitmap2Form = hRgn
    DeleteObject SelectObject(hDC, Picture)
End If
SetWindowRgn moform.hWnd, hRgn, True

DeleteDC hDC

End Function

