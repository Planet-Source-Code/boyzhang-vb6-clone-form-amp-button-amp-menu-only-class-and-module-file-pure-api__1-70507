VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Public mhWnd As Long
Private mCaption As String
Private mVisible As Boolean
Private mEnable As Boolean

'×ÓÀà»¯Àà
Private WithEvents CSubClass As Class_SubClass
Attribute CSubClass.VB_VarHelpID = -1
'ÊÂ¼þ
Public Event Click()


'¹¹Ôìº¯Êý
Private Sub Class_Initialize()
        'ÀàÊµÀý»¯
        Set CSubClass = New Class_SubClass
        'È±Ê¡Öµ
        mEnable = True
End Sub

Private Sub Class_Terminate()
        'ÍÑ¹³
        CSubClass.UnHook
        'ÊÍ·ÅÀà
        Set CSubClass = Nothing
End Sub

'×ÓÀà»¯¹ý³Ì
Private Sub CSubClass_NewMessage(ByVal hWnd As Long, FunctionPtr As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
        Const WM_COMMAND = &H111
        Const BN_CLICKED = 0
        
        Select Case uMsg
        Case WM_COMMAND
                If lParam = mhWnd And GetHiWord(wParam) = BN_CLICKED Then
                        '´¥·¢°´Å¥ÊÂ¼þ
                        RaiseEvent Click
                End If
        End Select
End Sub


'=====°´Å¥±êÌâ===================================================================================
Public Property Get Caption() As String
        Caption = mCaption
End Property

Public Property Let Caption(ByVal Text As String)
        mCaption = Text
        If IsWindow(mhWnd) <> 0 Then SetWindowText mhWnd, Text
End Property

'=====ÊÇ·ñ¿É¼û===================================================================================
Public Property Get Visible() As Boolean
        Visible = mVisible
End Property
Public Property Let Visible(ByVal State As Boolean)
        mVisible = State
        If IsWindow(mhWnd) <> 0 Then
                Const SW_NORMAL As Long = 1
                Const SW_HIDE = 0
                
                'ÏÔÊ¾/Òþ²Ø ´°Ìå
                If State = True Then
                        ShowWindow mhWnd, SW_NORMAL
                Else
                        ShowWindow mhWnd, SW_HIDE
                End If
                '¸üÐÂ´°Ìå
                UpdateWindow mhWnd
        End If
End Property

'=====ÊÇ·ñ¿ÉÓÃ===================================================================================
Public Property Get Enable() As Boolean
        Enable = mEnable
End Property
Public Property Let Enable(ByVal State As Boolean)
        mEnable = State
        If IsWindow(mhWnd) <> 0 Then
                Const SW_SHOW = 5
                Const SW_HIDE = 0
                
                'ÏÔÊ¾/Òþ²Ø ´°Ìå
                If State = True Then
                        EnableWindow mhWnd, SW_SHOW
                Else
                        EnableWindow mhWnd, SW_HIDE
                End If
                '¸üÐÂ´°Ìå
                UpdateWindow mhWnd
        End If
End Property

'Ïò´°ÌåÌí¼Ó°´Å¥
Public Function AppendButton(ByVal CFormParent As Class_Form, _
                             Optional ByVal X As Long = 0, _
                             Optional ByVal Y As Long = 0, _
                             Optional ByVal width As Long = 200, _
                             Optional ByVal height As Long = 60) As Boolean
        
        Const WS_CHILD = &H40000000
        Const WS_VISIBLE = &H10000000
        
        Dim Style As Long
        Style = WS_CHILD Or WS_VISIBLE
        mhWnd = CreateWindowEx(0&, "BUTTON", Caption, Style, X, Y, width, height, CFormParent.hWnd, vbNull, CApp.hInstance, ByVal 0&)
        Dim hFont As Long
        Const DEFAULT_GUI_FONT = 17
        Const WM_SETFONT = &H30
        If hFont = 0 Then hFont = GetStockObject(DEFAULT_GUI_FONT)
        SendMessage mhWnd, WM_SETFONT, hFont, 1
        Visible = True
        If mhWnd Then AppendButton = True
        
        '¹Ò¹³
        CSubClass.Hook CFormParent.hWnd
End Function


