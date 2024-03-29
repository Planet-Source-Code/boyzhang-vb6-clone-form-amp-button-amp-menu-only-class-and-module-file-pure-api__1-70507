VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_SubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Ä£¿éÃèÊö
'---------------------------------------------------------------------------------------
' Date-Time     : 2006-01-02 23:10:43
' Author        : Cadaver
' Purpose       : ¹Ò¹³Àà
'---------------------------------------------------------------------------------------
'½á¹¹Ìå
Private Type FunctionSPointerS
        FunctionPtr As Long
        FunctionAddress As Long
End Type
'APIº¯Êý
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'³£Êý
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
'±äÁ¿
Private mb_running As Boolean
Private mlng_memhandle As Long
Private mlng_proc As Long
Private mlng_handle As Long
Private mlng_hwnd As Long
'ÊÂ¼þ
Public Event NewMessage(ByVal hWnd As Long, ByRef FunctionPtr As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef Cancel As Boolean)

'×ÓÀà¹ý³Ì
Public Function SubClassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Dim Bool As Boolean
        '´¥·¢ÊÂ¼þ
        RaiseEvent NewMessage(hWnd, SubClassProc, uMsg, wParam, lParam, Bool)
        'ÏûÏ¢ÊÇ·ñ·µ»ØÔ­µØÖ·
        If Bool = True Then Exit Function
        '·µ»ØÔ­µØÖ·
        SubClassProc = CallWindowProc(mlng_handle, hWnd, uMsg, wParam, lParam)
End Function

'´«º¯ÊýÖ¸Õë
Public Function GetFunctionPtr()
        GetFunctionPtr = mlng_proc
End Function

'¹Ò¹³
Public Sub Hook(ByVal hWnd As Long, Optional Style As Long = -4)
        mlng_hwnd = hWnd
        If Not mb_running Then
                mb_running = True
                mlng_handle = SetWindowLong(mlng_hwnd, Style, mlng_proc)
        End If
End Sub

'ÍÑ¹³
Public Sub UnHook()
        If mb_running Then
                SetWindowLong mlng_hwnd, (-4), mlng_handle
                mb_running = False
        End If
End Sub

'¹¹Ôìº¯Êý
Private Sub Class_Initialize()
        'Éú³Éº¯ÊýÖ¸Õë
        Dim lng_fncptr As Long
        Dim lng_objptr As Long
        Dim lng_vtable As Long
        Dim lng_ptx As Long
        Dim lng_proc As Long
        Dim lng_varnum As Long
        Dim lng_objvarnum As Long
        Dim lng_funcnum As Long
        lng_objptr = ObjPtr(Me)
        CopyMemory lng_vtable, ByVal lng_objptr, 4
        lng_ptx = lng_vtable + 28 + (lng_varnum * 2 * 4) + (lng_objvarnum * 3 * 4) + lng_funcnum * 4
        CopyMemory lng_fncptr, ByVal lng_ptx, 4
        mlng_memhandle = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, 105)
        mlng_proc = GlobalLock(mlng_memhandle)
        DelegateFunction mlng_proc, Me, lng_fncptr, 4
End Sub

'Îö¹¹º¯Êý
Private Sub Class_Terminate()
        If mb_running Then UnHook
        Call GlobalUnlock(mlng_memhandle)
        Call GlobalFree(mlng_memhandle)
End Sub

'ÍÐ¹Üº¯Êý
Private Function DelegateFunction(ByVal CallingADR As Long, Obj As Object, ByVal MethodAddress As Long, ByVal NumberOfParameters As Byte)
        Dim TmpA As Long
        Dim u As Long
        TmpA = CallingADR
        CopyMemory ByVal CallingADR, &H68EC8B55, 4
        CallingADR = CallingADR + 4
        CopyMemory ByVal CallingADR, TmpA + 31 + (NumberOfParameters * 3) - 4, 4
        CallingADR = CallingADR + 4
        Dim StackP As Byte
        StackP = 4 + 4 * NumberOfParameters
        For u = 1 To NumberOfParameters
                CopyMemory ByVal CallingADR, CInt(&H75FF), 2
                CallingADR = CallingADR + 2
                CopyMemory ByVal CallingADR, StackP, 1
                CallingADR = CallingADR + 1
                StackP = StackP - 4
        Next u
        CopyMemory ByVal CallingADR, CByte(&H68), 1
        CallingADR = CallingADR + 1
        CopyMemory ByVal CallingADR, ObjPtr(Obj), 4
        CallingADR = CallingADR + 4
        CopyMemory ByVal CallingADR, CByte(&HE8), 1
        CallingADR = CallingADR + 1
        Dim PERFCALL As Long
        PERFCALL = CallingADR - TmpA - 1
        PERFCALL = MethodAddress - (TmpA + (CallingADR - TmpA - 1)) - 5
        CopyMemory ByVal CallingADR, PERFCALL, 4
        CallingADR = CallingADR + 4
        CopyMemory ByVal CallingADR, CByte(&HA1), 1
        CallingADR = CallingADR + 1
        CopyMemory ByVal CallingADR, TmpA + 31 + (NumberOfParameters * 3) - 4, 4
        CallingADR = CallingADR + 4
        CopyMemory ByVal CallingADR, CInt(&HC2C9), 2
        CallingADR = CallingADR + 2
        CopyMemory ByVal CallingADR, CInt(NumberOfParameters * 4), 2
End Function
