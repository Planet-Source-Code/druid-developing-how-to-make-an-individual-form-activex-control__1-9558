VERSION 5.00
Object = "{FAE1731F-430A-11D4-B183-D1B9690DF016}#22.0#0"; "FORMINDIVIDUAL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   5595
   ClientTop       =   4065
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   Picture         =   "DForm1.frx":0000
   ScaleHeight     =   2220
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin FormIndividual.IndForm IndForm1 
      Left            =   2760
      Top             =   840
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   0
      Picture         =   "DForm1.frx":9F2A
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Formen = IndForm1.Bitmap2Form(Me.Picture, vbBlack)
End Sub
