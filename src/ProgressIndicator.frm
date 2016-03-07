VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressIndicator 
   Caption         =   "Voortgang"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ProgressIndicator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    Dim out As Integer
    out = MsgBox("Het programma is vroegtijdig gestopt. Reeds gemaakte afbeeldingen zijn niet opgeruimd.", vbOKOnly Or vbExclamation, "Opgelet!")
    End
End Sub

Private Sub DoneButton_Click()
    End
End Sub

Private Sub UserForm_Activate()
    Convert_To_LaTeX
End Sub
