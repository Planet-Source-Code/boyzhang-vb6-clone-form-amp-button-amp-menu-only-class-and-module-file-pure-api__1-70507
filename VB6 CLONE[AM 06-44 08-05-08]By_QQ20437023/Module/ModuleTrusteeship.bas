Attribute VB_Name = "ModuleTrusteeship"
'ÍÐ¹ÜÄ£¿é
Option Explicit
'½á¹¹Ìå
Private Type WNDCLASS   '´°Ìå½á¹¹
        Style As Long
        lpfnwndproc As Long
        cbClsextra As Long
        cbWndExtra2 As Long
        hInstance As Long
        hIcon As Long
        hCursor As Long
        hbrBackground As Long
        lpszMenuName As String
        lpszClassName As String
End Type
Private Type POINTAPI   '×ø±ê½á¹¹
        X As Long
        Y As Long
End Type
Private Type Msg        'ÏûÏ¢½á¹¹
        hWnd As Long
        Message As Long
        wParam As Long
        lParam As Long
        Time As Long
        pt As POINTAPI
End Type
'APIº¯Êý
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)


'App¶ÔÏó
Public CApp As Class_Application
'ÆÁÄ»¶ÔÏó
Public CScreen As Class_Screen

'ÊÂ¼þÍÐ¹Ü´°Ìå
Private IForm As Class_Form

'³õÊ¼»¯ÏµÍ³¶ÔÏó
Public Function sysInitialize()
        'ÏµÍ³¶ÔÏóÀàÊµÀý»¯
        Set CApp = New Class_Application
        Set CScreen = New Class_Screen
End Function

'¸ßµÍÎ»
Public Function GetHiWord(ByVal Value As Long) As Integer
        RtlMoveMemory GetHiWord, ByVal VarPtr(Value) + 2, 2
End Function
Public Function GetLoWord(ByVal Value As Long) As Integer
        RtlMoveMemory GetLoWord, Value, 2
End Function

'ÍÐ¹Üº¯Êý
Public Function Trusteeship(ByRef EventForm As Class_Form) As Boolean
        'ÀàÊµÀý»¯
        Set IForm = EventForm
        Const WinClassName = "MyWinClass"               '¶¨Òå´°¿ÚÀàÃû
        
        Dim WC As WNDCLASS 'ÉèÖÃ´°Ìå²ÎÊý
        With WC
                .hIcon = 0                                      '´°ÌåÍ¼±ê Ê¹ÓÃ LoadIcon(hInstance, ID)   ¼ÓÔØRESÍ¼±ê
                .hCursor = 0                                    '´°Ìå¹â±ê Ê¹ÓÃ LoadCursor(hInstance, ID) ¼ÓÔØRES¹â±ê
                .lpszMenuName = vbNullString                    '´°Ìå²Ëµ¥ Ê¹ÓÃ LoadMenu(hInstance,ID)    ¼ÓÔØRES²Ëµ¥
                .hInstance = CApp.hInstance                     'ÊµÀý
                .cbClsextra = 0
                .cbWndExtra2 = 0
                .Style = 0
                .hbrBackground = 16
                .lpszClassName = WinClassName                   'ÀàÃû
                .lpfnwndproc = GetAddress(AddressOf WinProc)    'ÏûÏ¢º¯ÊýµØÖ·
        End With
        '×¢²á´°ÌåÀà
        If RegisterClass(WC) = 0 Then CApp.ErrDescription = "RegisterClass Faild.": Exit Function
        '»ñÈ¡´°Ìå¾ä±ú
        With IForm
                .hWnd = CreateWindowEx(0&, WinClassName, .Caption, .WindowStyle, .Left, .Top, .width, .height, 0, 0, CApp.hInstance, ByVal 0&)
                If .hWnd = 0 Then CApp.ErrDescription = "CreateWindowEx Faild.": Exit Function
                .hDC = GetDC(.hWnd)     '»ñÈ¡´°ÌåGDI¾ä±ú
                .Visible = True         'ÏÔÊ¾´°Ìå
                
                '´°Ìå´´½¨
                Call .ICreate
                
                Dim WinMsg As Msg       'ÏûÏ¢½á¹¹
                'ÏûÏ¢Ñ­»·
                Do While GetMessage(WinMsg, 0, 0, 0) > 0
                        TranslateMessage WinMsg
                        DispatchMessage WinMsg
                Loop
        End With
        
        '·µ»ØÖµ
        Trusteeship = True
End Function

'´°Ìå¹ý³Ì
Private Function WinProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Const WM_CREATE = &H1
        Const WM_COMMAND = &H111
        Const WM_CLOSE = &H10
        Const WM_MOUSEMOVE = &H200
        Const WM_SIZE = &H5
        Const WM_DESTROY = &H2


        Dim bRet As Boolean 'È¡·µ»ØÖµ
        With IForm
                Select Case wMsg
                Case WM_COMMAND
                        Call .ICommand(wParam, lParam)
                Case WM_CLOSE
                        Call .IUnload(bRet)
                        If bRet = True Then Exit Function
                        DestroyWindow .hWnd 'Ïú»Ù´°Ìå
                Case WM_MOUSEMOVE
                        Call .IMouseMove(LoWord(lParam), HiWord(lParam))
                Case WM_SIZE
                        Call .IResize
                Case WM_DESTROY
                        PostQuitMessage 0
                Case Else
                        WinProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
                End Select
        End With
End Function

'È¡µØÖ·
Private Function GetAddress(Address) As Long
        GetAddress = Address
End Function

'µÍ×Ö
Private Function LoWord(ByVal DWord As Long) As Integer
        If DWord And &H8000& Then
                LoWord = DWord Or &HFFFF0000
        Else
                LoWord = DWord And &HFFFF&
        End If
End Function

'¸ß×Ö
Private Function HiWord(ByVal DWord As Long) As Integer
        HiWord = (DWord And &HFFFF0000) \ 65536
End Function
