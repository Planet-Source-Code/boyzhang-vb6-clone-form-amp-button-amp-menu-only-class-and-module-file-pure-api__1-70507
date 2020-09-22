VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_MenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'²Ëµ¥ÌõÄ¿Êý×é
Private Menus() As Class_MenuItem
'ÉÏ±ê
Public Count As Long

'¹¹Ôìº¯Êý
Private Sub Class_Initialize()
        '³õÊ¼»¯Êý×é
        ReDim Menus(0&)
End Sub

'Îö¹¹º¯Êý
Private Sub Class_Terminate()
        Dim i As Long
        For i = 0& To Count
                'ÊÍ·ÅÀà
                Set Menus(i) = Nothing
        Next i
        'ÊÍ·ÅÊý×é
        Erase Menus
End Sub

'Ìí¼ÓÒ»¸ö²Ëµ¥Ïî
Public Function AddMenu() As Class_MenuItem
        'ÉÏ±êÀÛ¼Ó
        Count = Count + 1&
        'ÉêÇëÄÚ´æ
        ReDim Preserve Menus(Count)
        'ÀàÊµÀý»¯
        Set Menus(Count) = New Class_MenuItem
        '±£´æË÷Òý
        Menus(Count).Index = Count
        '´«³öË÷Òý
        Set AddMenu = Menus(Count)
End Function

'·ÃÎÊ½Ó¿Ú
Public Function GetItemByIndex(ByVal Index As Long) As Class_MenuItem
        Set GetItemByIndex = Menus(Index)
End Function

Public Function GetItemByhItem(ByVal hItem As Long) As Class_MenuItem
        Dim i As Long
        For i = 1& To Count
                Set GetItemByhItem = GetItemByIndex(i)
                If GetItemByhItem.hItem = hItem Then
                        Exit Function
                End If
        Next i
        Set GetItemByhItem = Nothing
End Function

Public Function GetItemByAlias(ByVal Alias As String) As Class_MenuItem
        Dim i As Long
        For i = 1& To Count
                Set GetItemByAlias = GetItemByIndex(i)
                If GetItemByAlias.Alias = Alias Then
                        Exit Function
                End If
        Next i
        Set GetItemByAlias = Nothing
End Function
