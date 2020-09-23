VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Code from: http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=43859&lngWId=1
    Private Type NodeType
        item As String
        Left As Integer
        Right As Integer
        Parent As Integer
        Tag As String
    End Type
    
    Dim Nodes() As NodeType
    Dim NoNodes As Integer
    Dim Sorted() As String
    Dim NoSorted As Integer
    Dim ItemTags() As String

Public Sub Clear()
    NoNodes = 0
    ReDim Nodes(1 To 1)
End Sub

'sort items returns the number of items, results contains the results
Public Function SortItems(Reverse As Boolean, ByRef Results() As String, ByRef Tags() As String) As Integer
    Dim i As Integer
    Dim a As Integer
    
    ReDim Sorted(1 To NoNodes)
    ReDim ItemTags(1 To NoNodes)
    NoSorted = 0
    Sort (1)
    
    SortItems = NoSorted
    ReDim Results(1 To NoSorted)
    ReDim Tags(1 To NoSorted)
    If Reverse Then
        a = NoSorted
        For i = 1 To NoSorted
            Results(a) = Sorted(i)
            Tags(a) = ItemTags(i)
            a = a - 1
        Next i
    Else
        For i = 1 To NoSorted
            Results(i) = Sorted(i)
            Tags(i) = ItemTags(i)
        Next i
    End If
End Function
Public Sub AddItem(item As String, Tag As String)
    Dim Parent As Integer
    Dim Ext As Boolean
    Dim i As Integer
    
    NoNodes = NoNodes + 1
    ReDim Preserve Nodes(1 To NoNodes)
    i = 1
    Nodes(NoNodes).item = item
    Nodes(NoNodes).Tag = Tag
    
    If NoNodes = 1 Then
        Ext = True
    End If
    Do Until Ext = True
        If item > Nodes(i).item Then
            If Nodes(i).Right <> 0 Then
                i = Nodes(i).Right
            Else
                Parent = i
                Nodes(i).Right = NoNodes
                Nodes(NoNodes).Parent = i
                Ext = True
            End If
        ElseIf item < Nodes(i).item Then
            If Nodes(i).Left <> 0 Then
                i = Nodes(i).Left
            Else
                Parent = i
                Nodes(i).Left = NoNodes
                Nodes(NoNodes).Parent = i
                Ext = True
            End If
        Else
            NoNodes = NoNodes - 1
            ReDim Preserve Nodes(1 To NoNodes)
            Ext = True
        End If
    Loop
End Sub

Private Sub Sort(i As Integer)
    If Nodes(i).Left <> 0 Then
        Sort (Nodes(i).Left)
    End If
    NoSorted = NoSorted + 1
    ReDim Preserve Sorted(1 To NoSorted)
    Sorted(NoSorted) = Nodes(i).item
    ItemTags(NoSorted) = Nodes(i).Tag
    If Nodes(i).Right <> 0 Then
        Sort (Nodes(i).Right)
    End If
End Sub
