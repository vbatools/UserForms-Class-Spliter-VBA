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

'========================================================================================
' Class: clsSpliter
' Author: VBATools
' Version: 1.0.2
' Creation Date: 05.12.2025 11:05
' Last Update Date: 07.12.2025 14:00
' License: Apache License
'========================================================================================

' Stores reference to parent frame
Private mParentFrame As Object

' Collection of splitter items
Private mItems      As Collection

Public Property Get Version() As String
    Version = "Version: 1.0.2" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Date of creation: 05.12.2025 11:05" & vbNewLine & _
            "Date of update: 07.12.2025 14:00"
End Property

' Returns collection of splitter items
Public Property Get Items() As Collection
    Set Items = mItems
End Property

Public Function ByItemForeColor(ByVal index As Long) As Long
    ' Returns text color of splitter item by index
    ' index: item index (starting from 1)
    ' Returns: text color of item
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    ByItemForeColor = item.ForeColor
    Set item = Nothing
End Function

Public Sub ByItemSetForeColor(ByVal index As Long, ByVal Color As Long)
    ' Sets text color of splitter item by index
    ' index: item index (starting from 1)
    ' Color: text color to set
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Sub
    item.ForeColor = Color
    Set item = Nothing
End Sub

Public Function ByItemCaption(ByVal index As Long) As String
    ' Returns caption of splitter item by index
    ' index: item index (starting from 1)
    ' Returns: caption of item
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    ByItemCaption = item.Caption
    Set item = Nothing
End Function

Public Sub ByItemSetCaption(ByVal index As Long, ByVal Caption As String)
    ' Sets caption of splitter item by index
    ' index: item index (starting from 1)
    ' Caption: caption to set
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Sub
    item.Caption = Caption
    Set item = Nothing
End Sub

Public Function ByItemVisible(ByVal index As Long) As Boolean
    ' Returns visibility of splitter item by index
    ' index: item index (starting from 1)
    ' Returns: True if item is visible, otherwise False
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    ByItemVisible = item.Visible
    Set item = Nothing
End Function

Public Sub ByItemSetVisible(ByVal index As Long, ByVal Visible As Boolean)
    ' Sets visibility of splitter item by index
    ' index: item index (starting from 1)
    ' Visible: visibility flag to set
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Sub
    item.Visible = Visible
    Set item = Nothing
End Sub

Public Function ByItemEnabled(ByVal index As Long) As Boolean
    ' Returns availability of splitter item by index
    ' index: item index (starting from 1)
    ' Returns: True if item is enabled, otherwise False
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Function
    ByItemEnabled = item.Enabled
    Set item = Nothing
End Function

Public Sub ByItemSetEnabled(ByVal index As Long, ByVal Enabled As Boolean)
    ' Sets availability of splitter item by index
    ' index: item index (starting from 1)
    ' Enabled: enabled flag to set
    Dim item        As clsSpliterItem
    Set item = Me.item(index)
    If item Is Nothing Then Exit Sub
    item.Enabled = Enabled
    Set item = Nothing
End Sub

Public Property Get item(ByVal index As Long) As clsSpliterItem
    ' Returns splitter item by index
    ' index: item index (starting from 1)
    ' Returns: splitter item or Nothing if index is out of range
    Set item = Nothing
    If index >= 1 And index <= mItems.Count Then
        Set item = mItems(index)
    End If
End Property

Public Property Get Count() As Long
    ' Returns number of splitter items
    Count = mItems.Count
End Property

Public Sub Initialize(ByRef objParent As Object)
    ' Initializes splitter with specified parent object
    ' objParent: parent object where splitters will be placed
    Set mParentFrame = objParent
    Set mItems = New Collection
End Sub

Public Function AddItem(ByRef Control_1 As MSForms.Control, ByRef Control_2 As MSForms.Control, Optional lForeColor As Long = &HC0C0C0) As clsSpliterItem
    ' Adds new splitter item between two controls
    ' Control_1: first control
    ' Control_2: second control
    ' lForeColor: text color of splitter (default &HC0C0C0)
    ' Returns: created splitter item
    Dim spliterItem As clsSpliterItem
    Set spliterItem = New clsSpliterItem
    Call spliterItem.AddItem_(Me, mParentFrame, mItems, Control_1, Control_2, lForeColor)
    Call mItems.Add(spliterItem, spliterItem.Name)
    Set AddItem = spliterItem
End Function

Public Sub setForeColorAll(ByVal Color As Long)
    ' Sets text color for all splitter items
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
    ' Removes all splitter items
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

Public Function ByItemRemove(ByVal index As Long) As Boolean
    ' Removes splitter item by index
    ' index: index of item to remove (starting from 1)
    ' Returns: True if item was successfully removed, otherwise False
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