VERSION 5.00
Object = "{FAE1731F-430A-11D4-B183-D1B9690DF016}#22.0#0"; "FORMINDIVIDUAL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   3900
   ClientTop       =   3930
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin FormIndividual.IndForm IndForm1 
      Left            =   2280
      Top             =   1440
      _ExtentX        =   873
      _ExtentY        =   873
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Formen = IndForm1.FormText("Times new Roman", 60, "DEMO")
End Sub
