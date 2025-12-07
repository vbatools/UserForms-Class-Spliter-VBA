VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpliterItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFrmParent  As Object
Private mControl_1  As MSForms.Control
Private mControl_2  As MSForms.Control
Private mNameCntrol_1 As String
Private mNameCntrol_2 As String
Private mNameLabel  As String
Private mIsHorizont As Boolean
Private WithEvents mSpliterFrame As MSForms.Frame
Attribute mSpliterFrame.VB_VarHelpID = -1
Private WithEvents mSpliterLabel As MSForms.Label
Attribute mSpliterLabel.VB_VarHelpID = -1

Private mItems      As Collection
Private mParent     As clsSpliter

Public Property Get Visible() As Boolean
    Visible = mSpliterLabel.Visible
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
    mSpliterLabel.Visible = bVisible
    mSpliterFrame.Visible = bVisible
End Property

Public Property Get Enabled() As Boolean
    Enabled = mSpliterLabel.Enabled
End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)
    mSpliterLabel.Enabled = bEnabled
    mSpliterFrame.Enabled = bEnabled
End Property

Public Property Get ForeColor() As Long
    ForeColor = mSpliterLabel.ForeColor
End Property

Public Property Let ForeColor(ByVal Color As Long)
    mSpliterLabel.ForeColor = Color
End Property

Public Property Get Caption() As String
    Caption = mSpliterLabel.Caption
End Property

Public Property Let Caption(ByVal sCaption As String)
    mSpliterLabel.Caption = sCaption
End Property

Public Property Get SpliterLabel() As MSForms.Label
    Set SpliterLabel = mSpliterLabel
End Property

Public Property Get SpliterFrame() As MSForms.Frame
    Set SpliterFrame = mSpliterFrame
End Property

Public Property Get Control_1() As MSForms.Control
    Set Control_1 = mControl_1
End Property

Public Property Get Control_2() As MSForms.Control
    Set Control_2 = mControl_2
End Property

Public Property Get IsHorizont() As Boolean
    IsHorizont = mIsHorizont
End Property

Public Property Get Name() As String
    Name = mNameLabel
End Property

Public Sub AddItem_(ByRef Parent As clsSpliter, ByRef objParent As Object, ByRef Items As Collection, ByRef Control_1 As MSForms.Control, _
        ByRef Control_2 As MSForms.Control, Optional lForeColor As Long = &HC0C0C0)
    Set mFrmParent = objParent
    Set mControl_1 = Control_1
    Set mControl_2 = Control_2
    mNameCntrol_1 = mControl_1.Name
    mNameCntrol_2 = mControl_2.Name
    mNameLabel = mNameCntrol_1 & "_" & mNameCntrol_2

    Set mSpliterFrame = objParent.Controls.Add("Forms.frame.1", mNameLabel & "_frm", True)
    With mSpliterFrame
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .Caption = vbNullString

        With .Font
            .Name = "Marlett"
            .Size = 8
            .Bold = True
        End With
        Set mSpliterLabel = .Controls.Add("Forms.label.1", mNameLabel & "_lbl", True)
        With mSpliterLabel
            .BackStyle = fmBackStyleTransparent
            .AutoSize = True
            .BorderStyle = fmBorderStyleNone
            .ForeColor = lForeColor
            .ZOrder
        End With
        'VERTICAL
        If mControl_1.Top = mControl_2.Top Then
            .MousePointer = fmMousePointerSizeWE
            mSpliterLabel.Caption = "34"

            If mControl_1.Left <= mControl_2.Left Then
                .Width = mControl_2.Left - mControl_1.Left - mControl_1.Width
                .Left = mControl_1.Left + mControl_1.Width
            Else
                .Width = mControl_1.Left - mControl_2.Left - mControl_2.Width
                .Left = mControl_2.Left + mControl_2.Width
            End If
        End If
        'HORIZON
        If mControl_1.Left = mControl_2.Left Then
            mIsHorizont = True
            .MousePointer = fmMousePointerSizeNS
            mSpliterLabel.Caption = "56"
            mSpliterLabel.Width = 19

            If mControl_1.Top <= mControl_2.Top Then
                .Top = mControl_1.Top + mControl_1.Height
                .Height = mControl_2.Top - mControl_1.Top - mControl_1.Height
            Else
                .Top = mControl_2.Top + mControl_2.Height
                .Height = mControl_1.Top - mControl_2.Top - mControl_2.Height
            End If
        End If
        mSpliterLabel.Name = mNameLabel
        Call resizeFrame(mSpliterLabel, mSpliterFrame, mControl_1, mControl_2, mIsHorizont)
    End With
    Set mParent = Parent
    Set mItems = Items
End Sub

Private Sub resizeFrame(ByRef SpliterLabel As MSForms.Label, ByRef SpliterFrame As MSForms.Frame, _
        ByRef Control_1 As MSForms.Control, ByRef Control_2 As MSForms.Control, ByRef IsHorizont As Boolean)
    With SpliterFrame
        If IsHorizont Then
            If Control_1.Width <= Control_2.Width Then
                .Width = Control_1.Width
                .Left = Control_1.Left
                SpliterLabel.Left = Control_1.Width / 2 - SpliterLabel.Width / 2
            Else
                .Width = Control_2.Width
                .Left = Control_2.Left
                SpliterLabel.Left = Control_2.Width / 2 - SpliterLabel.Width / 2
            End If
        Else
            If Control_1.Height <= Control_2.Height Then
                .Top = Control_1.Top
                .Height = Control_1.Height
                SpliterLabel.Top = Control_1.Height / 2 - SpliterLabel.Height / 2
            Else
                .Top = Control_2.Top
                .Height = Control_2.Height
                SpliterLabel.Top = Control_2.Height / 2 - SpliterLabel.Height / 2
            End If
        End If
    End With
End Sub

Private Sub mSpliterFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then Call resizeControls(mSpliterLabel, mSpliterFrame, mControl_1, mControl_2, X, Y, mIsHorizont)
End Sub

Private Sub mSpliterLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then Call resizeControls(mSpliterLabel, mSpliterFrame, mControl_1, mControl_2, X, Y, mIsHorizont)
End Sub

Private Sub resizeControls(ByRef SpliterLabel As MSForms.Label, ByRef SpliterFrame As MSForms.Frame, _
        ByRef Control_1 As MSForms.Control, ByRef Control_2 As MSForms.Control, ByRef X As Single, ByRef Y As Single, _
        ByRef IsHorizont As Boolean)

    If cantMove(Control_1, Control_2, X, Y, IsHorizont) Then Exit Sub
    If Not IsHorizont Then
        'VERTICAL
        If Control_1.Width + X >= 0 And Control_2.Width - X >= 0 Then
            SpliterFrame.Move SpliterFrame.Left + X, SpliterFrame.Top
            Control_1.Width = Control_1.Width + X
            Control_2.Left = Control_2.Left + X
            Control_2.Width = Control_2.Width - X
        End If
    Else
        'HORIZON
        If Control_1.Height + Y >= 0 And Control_2.Height - Y >= 0 Then
            SpliterFrame.Move SpliterFrame.Left, SpliterFrame.Top + Y
            Control_1.Height = Control_1.Height + Y
            Control_2.Top = Control_2.Top + Y
            Control_2.Height = Control_2.Height - Y
        End If
    End If
    Call resizeFrame(SpliterLabel, SpliterFrame, Control_1, Control_2, IsHorizont)
    Call updateAllSplitter(Control_1, Control_2)
End Sub

Private Sub updateAllSplitter(ByRef Control_1 As MSForms.Control, ByRef Control_2 As MSForms.Control)

    Dim item        As clsSpliterItem
    For Each item In mItems
        With item
            If mNameLabel <> .Name Then
                If mNameCntrol_1 = .Control_1.Name Or mNameCntrol_2 = .Control_1.Name Or _
                        mNameCntrol_1 = .Control_2.Name Or mNameCntrol_2 = .Control_2.Name Then
                    Call resizeFrame(.SpliterLabel, .SpliterFrame, .Control_1, .Control_2, .IsHorizont)
                End If
            End If
        End With
    Next item
    Set item = Nothing
End Sub

Private Function cantMove(ByRef Control_1 As MSForms.Control, ByRef Control_2 As MSForms.Control, ByRef X As Single, ByRef Y As Single, ByRef IsHorizont As Boolean) As Boolean

    Dim item        As clsSpliterItem
    For Each item In mItems
        With item
            If mNameLabel <> .Name Then
                If mNameCntrol_1 = .Control_1.Name Or mNameCntrol_2 = .Control_1.Name Or _
                        mNameCntrol_1 = .Control_2.Name Or mNameCntrol_2 = .Control_2.Name Then
                    If IsHorizont Then
                        'down-up
                        If intersectionTwoControls(Control_1, .Control_2, 0, Y * -1) Then cantMove = True
                        If intersectionTwoControls(Control_1, .Control_1, 0, Y * -1) Then cantMove = True
                        'up-down
                        If intersectionTwoControls(Control_2, .Control_2, 0, Y) Then cantMove = True
                        If intersectionTwoControls(Control_2, .Control_1, 0, Y) Then cantMove = True
                    Else
                        'left-right
                        If intersectionTwoControls(Control_1, .Control_2, X * -1, 0) Then cantMove = True
                        If intersectionTwoControls(Control_1, .Control_1, X * -1, 0) Then cantMove = True
                        'right-left
                        If intersectionTwoControls(Control_2, .Control_2, X, 0) Then cantMove = True
                        If intersectionTwoControls(Control_2, .Control_1, X, 0) Then cantMove = True
                    End If
                    If cantMove Then Exit Function
                End If
            End If
        End With
    Next item
    Set item = Nothing
End Function

Private Function intersectionTwoControls(ByRef cnt_1 As MSForms.Control, ByRef cnt_2 As MSForms.Control, ByRef X As Single, ByRef Y As Single) As Boolean
    If (cnt_1.Top + Y < cnt_2.Top + cnt_2.Height And cnt_2.Top + Y < cnt_1.Top + cnt_1.Height) And _
            (cnt_1.Left + X < cnt_2.Left + cnt_2.Width And cnt_2.Left + X < cnt_1.Left + cnt_1.Width) And _
            cnt_2.Name <> cnt_1.Name Then
        'Debug.Print cnt_1.Name; " "; cnt_2.Name
        intersectionTwoControls = True
    End If
End Function

Private Sub Class_Terminate()
    Set mControl_1 = Nothing
    Set mControl_2 = Nothing
    Set mSpliterFrame = Nothing
    Set mSpliterLabel = Nothing
    Set mFrmParent = Nothing
End Sub
