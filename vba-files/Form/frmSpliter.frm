VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpliter 
   Caption         =   "UserForm1"
   ClientHeight    =   8775.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "frmSpliter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpliter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SP              As clsSpliter

Private Sub CheckBox1_Click()
    Call SP.setVisibleAll(CheckBox1.Value)
End Sub

Private Sub CheckBox2_Click()
    Call SP.setEnabledAll(CheckBox2.Value)
End Sub

Private Sub CheckBox3_Click()
    Debug.Print SP.ByItemEnabled(1)
    Call SP.ByItemSetEnabled(1, CheckBox3.Value)
End Sub

Private Sub CheckBox4_Click()
    Debug.Print SP.ByItemVisible(1)
    Call SP.ByItemSetVisible(1, CheckBox4.Value)
End Sub

Private Sub CommandButton1_Click()
    Call SP.setForeColorAll(vbBlue)
End Sub

Private Sub CommandButton2_Click()
    Call SP.setCaptionAll(8)
End Sub

Private Sub CommandButton5_Click()
    Dim item As clsSpliterItem
    Set item = SP.item(1)
End Sub

Private Sub CommandButton6_Click()
    SP.RemoveAll
End Sub

Private Sub CommandButton7_Click()
    Debug.Print SP.ByItemRemove(1)
End Sub

Private Sub CommandButton8_Click()
    Debug.Print SP.ByItemForeColor(1)
    Call SP.ByItemSetForeColor(1, vbRed)
End Sub

Private Sub CommandButton9_Click()
    Debug.Print SP.ByItemCaption(1)
    Call SP.ByItemSetCaption(1, 5)
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    Set SP = New clsSpliter
    Call SP.Initialize(Me)
    With SP
        Call .AddItem(Frame1, TextBox2)
        Call .AddItem(TextBox2, TextBox3)

        Call .AddItem(TextBox4, TextBox5)
        Call .AddItem(TextBox5, TextBox6)

        Call .AddItem(TextBox7, TextBox8)
        Call .AddItem(TextBox8, TextBox9)

        Call .AddItem(Frame1, TextBox4)
        Call .AddItem(TextBox4, TextBox7)

        Call .AddItem(TextBox2, TextBox5)
        Call .AddItem(TextBox5, TextBox8)

        Call .AddItem(TextBox3, TextBox6)
        Call .AddItem(TextBox6, TextBox9, vbRed)
    End With
End Sub