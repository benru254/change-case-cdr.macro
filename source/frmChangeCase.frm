VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangeCase 
   Caption         =   "ChangeCase"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5505
   OleObjectBlob   =   "frmChangeCase.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChangeCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSentence_Click()
    SentenceCase
End Sub
Private Sub cmdtitle_Click()
    TitleCase
End Sub

Private Sub cmdupper_Click()
    UpperCase
End Sub
Private Sub cmdlower_Click()
    LowerCase
End Sub
Private Sub cmdToggle_Click()
    ToggleCase
End Sub

Private Sub Label5_Click()
ShellExecute 0, vbNullString, "http://www.EngravingConcepts.com", vbNullString, vbNullString, 5
End Sub
