VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpliter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mParentFrame As Object
Private mItems      As Collection

Public Property Get Items() As Collection
    Set Items = mItems
End Property

Public Function ByItemForeColor(ByVal index As Byte) As Long
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    ByItemForeColor = item.ForeColor
    Set item = Nothing
End Function

Public Sub ByItemSetForeColor(ByVal index As Byte, ByVal Color As Long)
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Sub
    item.ForeColor = Color
    Set item = Nothing
End Sub

Public Function ByItemCaption(ByVal index As Byte) As String
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    ByItemCaption = item.Caption
    Set item = Nothing
End Function

Public Sub ByItemSetCaption(ByVal index As Byte, ByVal Caption As String)
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Sub
    item.Caption = Caption
    Set item = Nothing
End Sub

Public Function ByItemVisible(ByVal index As Byte) As Boolean
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    ByItemVisible = item.Visible
    Set item = Nothing
End Function

Public Sub ByItemSetVisible(ByVal index As Byte, ByVal Visible As Boolean)
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Sub
    item.Visible = Visible
    Set item = Nothing
End Sub

Public Function ByItemEnabled(ByVal index As Byte) As Boolean
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    ByItemEnabled = item.Enabled
    Set item = Nothing
End Function

Public Sub ByItemSetEnabled(ByVal index As Byte, ByVal Enabled As Boolean)
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Sub
    item.Enabled = Enabled
    Set item = Nothing
End Sub

Public Property Get item(ByVal index As Byte) As clsSpliterItem
    On Error GoTo endGetItem
    Set item = mItems(index)
    Exit Property
endGetItem:
    Err.Clear
End Property

Public Property Get Count() As Byte
    Count = mItems.Count
End Property

Public Sub Initialize(ByRef objParent As Object)
    Set mParentFrame = objParent
    Set mItems = New Collection
End Sub

Public Function AddItem(ByRef Control_1 As MSForms.Control, ByRef Control_2 As MSForms.Control, Optional lForeColor As Long = &HC0C0C0) As clsSpliterItem
    Dim spliterItem As clsSpliterItem
    Set spliterItem = New clsSpliterItem
    Call spliterItem.AddItem_(Me, mParentFrame, mItems, Control_1, Control_2, lForeColor)
    Call mItems.Add(spliterItem, spliterItem.Name)
    Set AddItem = spliterItem
End Function

Public Sub setForeColorAll(ByVal Color As Long)
    Dim item        As clsSpliterItem
    For Each item In mItems
        item.ForeColor = Color
    Next item
    Set item = Nothing
End Sub

Public Sub setCaptionAll(ByVal Caption As Long)
    Dim item        As clsSpliterItem
    For Each item In mItems
        item.Caption = Caption
    Next item
    Set item = Nothing
End Sub

Public Sub setEnabledAll(ByVal bEnabled As Boolean)
    Dim item        As clsSpliterItem
    For Each item In mItems
        item.Enabled = bEnabled
    Next item
    Set item = Nothing
End Sub

Public Sub setVisibleAll(ByVal bVisible As Boolean)
    Dim item        As clsSpliterItem
    For Each item In mItems
        With item
            .SpliterLabel.Visible = bVisible
            .SpliterFrame.Visible = bVisible
        End With
    Next item
    Set item = Nothing
End Sub

Public Sub RemoveAll()
    Dim item        As clsSpliterItem
    For Each item In mItems
        With item
            Call mParentFrame.Controls.Remove(.SpliterLabel.Name)
            Call mParentFrame.Controls.Remove(.SpliterFrame.Name)
        End With
    Next item
    Set item = Nothing
    Set mItems = Nothing
    Set mItems = New Collection
End Sub

Public Function ByItemRemove(ByVal index As Byte) As Boolean
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    With item
        Call mParentFrame.Controls.Remove(.SpliterLabel.Name)
        Call mParentFrame.Controls.Remove(.SpliterFrame.Name)
        Call mItems.Remove(index)
    End With
    Set item = Nothing
    ByItemRemove = True
End Function

Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mParentFrame = Nothing
End Sub


