VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'¾ä±úÊÇ·ñÓÐÐ§
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
'¿É¼û
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
'±êÌâ
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'×ø±ê
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'ÑùÊ½
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'----------------------------------------------------------------------------------------------------------------
'´°ÌåÑùÊ½
Public Enum eWindowStyle
        None = &H6000000
        Fixed_Single = &H6C80000
        Sizable = &H6CF0000
        Fixed_Dialog = &H6C80080
        Fixed_ToolWindow = &H6C80000
        Sizable_ToolWindow = &H6CC0000
End Enum

Private mVisible        As Boolean      '¿É¼û
Private mCaption        As String       '±êÌâ
Private mHeight         As Long         '¸ß¶È
Private mWidth          As Long         '¿í¶È
Private mTop            As Long         '¶¥±ß
Private mLeft           As Long         '×ó±ß
Private mMostTop        As Boolean      '²ã´Î
Private mCenter         As Boolean      '¾ÓÖÐ
Private mWindowStyle    As eWindowStyle 'ÑùÊ½

Public hWnd             As Long         '´°Ìå¾ä±ú
Public hDC              As Long         'Éè±¸¾ä±ú

'ÊÂ¼þ
Public Event Create()
Public Event Resize()
Public Event MouseMove(ByVal X As Integer, ByVal Y As Integer)
Public Event Unload(ByRef Cancel As Boolean)
Public Event Command(ByVal wParam As Long, ByVal lParam As Long)



'¹¹Ôìº¯Êý
Private Sub Class_Initialize()
        'È±Ê¡Öµ
        mWidth = 200
        mHeight = 200
        mWindowStyle = Sizable
        mCaption = "Windows GUI App"
End Sub

'Îö¹¹º¯Êý
Private Sub Class_Terminate()
        'ÊÍ·ÅÀà
        Set CScreen = Nothing
End Sub



'=====ÊÂ¼þ½Ó¿Ú===================================================================================
Public Sub ICreate()
        RaiseEvent Create
End Sub

Public Sub IMouseMove(ByVal X As Integer, ByVal Y As Integer)
        RaiseEvent MouseMove(X, Y)
End Sub

Public Function IResize()
        RaiseEvent Resize
End Function

Public Function IUnload(ByRef Cancel As Boolean)
        RaiseEvent Unload(Cancel)
End Function

Public Function ICommand(ByVal wParam As Long, ByVal lParam As Long)
        RaiseEvent Command(wParam, lParam)
End Function


'=====ÊÇ·ñ¿É¼û===================================================================================
Public Property Get Visible() As Boolean
        Visible = mVisible
End Property
Public Property Let Visible(ByVal State As Boolean)
        mVisible = State
        If IsWindow(hWnd) <> 0 Then
                Const SW_NORMAL As Long = 1
                Const SW_HIDE = 0
                
                'ÏÔÊ¾/Òþ²Ø ´°Ìå
                If State = True Then
                        ShowWindow hWnd, SW_NORMAL
                Else
                        ShowWindow hWnd, SW_HIDE
                End If
                '¸üÐÂ´°Ìå
                UpdateWindow hWnd
        End If
End Property

'=====±êÌâÎÄ×Ö===================================================================================
Public Property Get Caption() As String
        Caption = mCaption
End Property

Public Property Let Caption(ByVal Text As String)
        mCaption = Text
        If IsWindow(hWnd) <> 0 Then SetWindowText hWnd, Text
End Property


'=====´°Ìå¸ß¶È===================================================================================
Public Property Get height() As Long
        height = mHeight
End Property

Public Property Let height(ByVal Value As Long)
        mHeight = Value
        Call SetWinPos
End Property

'=====´°Ìå¿í¶È===================================================================================
Public Property Get width() As Long
        width = mWidth
End Property

Public Property Let width(ByVal Value As Long)
        mWidth = Value
        Call SetWinPos
End Property

'=====´°Ìå¶¥±ß===================================================================================
Public Property Get Top() As Long
        Top = mTop
End Property

Public Property Let Top(ByVal Value As Long)
        mTop = Value
        Call SetWinPos
End Property

'=====´°Ìå×ó±ß===================================================================================
Public Property Get Left() As Long
        Left = mLeft
End Property

Public Property Let Left(ByVal Value As Long)
        mLeft = Value
        Call SetWinPos
End Property

'=====´°ÌåÖÃ¶¥===================================================================================
Public Property Get MostTop() As Boolean
        MostTop = mMostTop
End Property

Public Property Let MostTop(ByVal Value As Boolean)
        mMostTop = Value
        Call SetWinPos
End Property

'=====´°Ìå¾ÓÖÐ===================================================================================
Public Property Get Center() As Boolean
        Center = mCenter
End Property

Public Property Let Center(ByVal Value As Boolean)
        mCenter = Value
        mLeft = (CScreen.width - mWidth) \ 2
        mTop = (CScreen.height - mHeight) \ 2
        Call SetWinPos
End Property

'=====¹¤¾ßº¯Êý===================================================================================
Private Sub SetWinPos()
        Const HWND_TOPMOST As Long = -1
        Const SWP_SHOWWINDOW = &H40
        If IsWindow(hWnd) <> 0 Then
                SetWindowPos hWnd, IIf(mMostTop = True, HWND_TOPMOST, 0), mLeft, mTop, mWidth, mHeight, SWP_SHOWWINDOW
                '¸üÐÂ´°Ìå
                UpdateWindow hWnd
        End If
End Sub

'=====´°ÌåÑùÊ½===================================================================================
Public Property Get WindowStyle() As eWindowStyle
        WindowStyle = mWindowStyle
End Property

Public Property Let WindowStyle(ByVal Value As eWindowStyle)
        mWindowStyle = Value
        
        Const GWL_STYLE = (-16)
        If IsWindow(hWnd) <> 0 Then
                SetWindowLong hWnd, GWL_STYLE, mWindowStyle
                '¸üÐÂ´°Ìå
                UpdateWindow hWnd
        End If
End Property



