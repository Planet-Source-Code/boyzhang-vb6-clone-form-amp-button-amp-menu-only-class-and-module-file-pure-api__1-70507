VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'¶¨Òå´°ÌåÀà(ÊÂ¼þÊä³ö)
Private WithEvents CForm As Class_Form
Attribute CForm.VB_VarHelpID = -1
'¶¨Òå°´Å¥Àà
Private WithEvents CButton1 As Class_Button
Attribute CButton1.VB_VarHelpID = -1
Private WithEvents CButton2 As Class_Button
Attribute CButton2.VB_VarHelpID = -1
Private WithEvents CButton3 As Class_Button
Attribute CButton3.VB_VarHelpID = -1
'¶¨Òå²Ëµ¥Àà
Private WithEvents CMenus As Class_MenuManage
Attribute CMenus.VB_VarHelpID = -1

'µã»÷°´Å¥
Private Sub CButton1_Click()
        MsgBox "now is:" & Now()
End Sub

Private Sub CButton2_Click()
        If CButton1.Enable = True Then
                CButton1.Enable = False
        Else
                CButton1.Enable = True
        End If
End Sub

Private Sub CButton3_Click()
        MsgBox CApp.CommandLine, , "command line parameter"
End Sub

'Ö÷Àà¹¹Ôìº¯Êý
Private Sub Class_Initialize()
        '³õÊ¼»¯ÏµÍ³¶ÔÏó
        Call sysInitialize
        'ÀàÊµÀý»¯
        Set CForm = New Class_Form
        'ÉèÖÃ²ÎÊý
        With CForm
                .width = 400
                .height = 300
                .Center = True
                .Caption = "Hello!"
        End With
        'ÍÐ¹Ü´°ÌåÀà
        If Trusteeship(CForm) = False Then Debug.Print CApp.ErrDescription
End Sub

'Ö÷ÀàÎö¹¹º¯Êý
Private Sub Class_Terminate()
        'ÊÍ·ÅÀà
        Set CForm = Nothing
End Sub

'---------------------------------------------------------------------------------------------------------------
'´°ÌåÀàÊÂ¼þ
'---------------------------------------------------------------------------------------------------------------

Private Sub CForm_Create()
        '´´½¨´°ÌåÊ±Ìí¼Ó¿Ø¼þ
        
        '°´Å¥
        Set CButton1 = New Class_Button
        CButton1.AppendButton CForm, 10, 10, 50, 30
        CButton1.Caption = "Button1"
        CButton1.Visible = True
        
        Set CButton2 = New Class_Button
        CButton2.AppendButton CForm, 80, 10, 50, 30
        CButton2.Caption = "Button2"
        CButton2.Visible = True
        
        Set CButton3 = New Class_Button
        CButton3.AppendButton CForm, 150, 10, 50, 30
        CButton3.Caption = "Button3"
        CButton3.Visible = True
        
        '²Ëµ¥
        Set CMenus = New Class_MenuManage
        With CMenus
                If .Initial(CForm.hWnd) = False Then End
                
                .AddForm("File(&S)", "File").HaveSub = True
                        .AddStd "File", "New(&N)         Ctrl+N", "New"
                        .AddStd "File", "Open(&O)...   Ctrl+O"
                        .AddStd "File", "Save(&S)        Ctrl+S"
                        .AddStd("File", "Save As(&A)...", "SaveAs").HaveSub = True
                                .AddStd "SaveAs", "HTML(&H)"
                                .AddStd "SaveAs", "Txt(&T)"
                        .AddStd("File").IsLine = True
                        .AddStd "File", "Page Setting(&U)..."
                        .AddStd "File", "Print(&P)...     Ctrl+P"
                        .AddStd("File").IsLine = True
                        .AddStd "File", "Exit(&X)", "Exit"
                
                .AddForm("Edit(&E)", "Edit").HaveSub = True
                        .AddStd("Edit", "UnDo(&U)      Ctrl+Z").Enabled = False
                        
                        .AddStd("Edit").IsLine = True
                        
                        .AddStd "Edit", "Cut(&T)       Ctrl+X"
                        .AddStd "Edit", "Copy(&C)      Ctrl+C"
                        .AddStd "Edit", "Paste(&P)     Ctrl+V"
                        .AddStd "Edit", "Delete(&D)"
                        
                        .AddStd("Edit").IsLine = True
                        
                        .AddStd "Edit", "Find(&F)Ctrl+F"
                        .AddStd "Edit", "Find Nex(&N)F3"
                        .AddStd "Edit", "Replace(&R)Ctrl+H"
                        .AddStd "Edit", "To Row(&G)Ctrl+G"
                        
                        .AddStd("Edit").IsLine = True
                        
                        .AddStd "Edit", "Select All(&A)Ctrl+A"
                        .AddStd "Edit", "Time/Date(&D) F5"
                
                .AddForm("Format(&F)", "Format").HaveSub = True
                        .AddStd("Format", "Auto Wrap(&W)", "Wrap").Checked = True
                        .AddStd "Format", "Font(&F)..."
                        
                .AddForm("Window(&W)", "Window").HaveSub = True
                        .AddStd("Window", "Split(&S)", "Split").Checked = True
                        
                        .AddStd("Window").IsLine = True
                        With .AddStd("Window", "Horizontal Tile(&H)", "HTile")
                                .IsRadio = True
                                .Checked = True
                        End With
                        .AddStd("Window", "Verticality Tile(&V)", "VTile").IsRadio = True
                        .AddStd("Window", "Cascade(&C)", "Cascade").IsRadio = True
                        .AddStd("Window", "Icon(&I)", "Icon").IsRadio = True
                        
                        .AddStd("Window").IsLine = True
                        
                        .AddStd("Window", "Window Lists(&L)", "WindowList").HaveSub = True
                        .AddStd "WindowList", "[" & CForm.Caption & "]"
                        
           With .AddForm("Help(&H)", "Help")
                .HaveSub = True
                .IsHelp = True
           End With
                        .AddStd "Help", "Theme(&T)", "Theme"
                        .AddStd("Help").IsLine = True
                        .AddStd "Help", "About(&A)", "About"
                        
                .AddPop "UnDo(&U)"
                .AddPop.IsLine = True
                .AddPop "Cut(&T)"
                .AddPop "Copy(&C)"
                .AddPop "Paste(&P)"
                .AddPop "Delete(&D)"
                .AddPop.IsLine = True
                .AddPop "Select All(&A)"
                .AddPop.IsLine = True
                
                .AddPop("Language(&L)", "Language").HaveSub = True
                        .AddStd "Language", "Chinese(GB2312)(&C)"
                        .AddStd "Language", "English(UTF-8)(&E)"
                        
                .AddSys.IsLine = True
                .AddSys "About(&A)...", "About"
                .AddSys.IsLine = True
                With .GetItemByAlias("Window")
                        '½«Window²Ëµ¥°ó¶¨µ½¸ÃÏî
                        CMenus.AddSys(.Caption, , .hItem).HaveSub = True
                End With
                .AddSys("TopMost(&T)", "TopMost").Checked = True
                
                .Enabled = True
                
                '°´ÏÂÐÂ½¨²Ëµ¥
                .EventMethod .GetItemByAlias("New")
        End With

        
        
        '´°ÌåÖÃ¶¥
        CForm.MostTop = True
End Sub

'²Ëµ¥ÊÂ¼þ
Private Sub CMenus_Click(MenuItem As Class_MenuItem)
        With MenuItem
                Select Case .Alias
                Case "TopMost"
                        If CForm.MostTop = True Then
                                CForm.MostTop = False
                        Else
                                CForm.MostTop = True
                        End If
                Case "New"
                        CForm.Caption = "ÐÂ½¨"
                Case "Exit"
                        'Ð¶ÔØ´°Ìå
                        CApp.Unload CForm
                Case "HTile", "VTile", "Cascade", "Icon"
                        With CMenus
                                .GetItemByAlias("HTile").Checked = False
                                .GetItemByAlias("VTile").Checked = False
                                .GetItemByAlias("Cascade").Checked = False
                                .GetItemByAlias("Icon").Checked = False
                        End With
                        .Checked = True
                Case "About"
                        MsgBox "By QQ:20437023 China" & vbCrLf & _
                               "MSN:Liaisonme@hotmail.com"
                Case "Wrap", "Split", "TopMost"
                        .Checked = Not .Checked
                End Select
        End With
End Sub

Private Sub CForm_MouseMove(ByVal X As Integer, ByVal Y As Integer)
        CForm.Caption = CStr(X) & "/" & CStr(Y)
End Sub

Private Sub CForm_Unload(Cancel As Boolean)
        If MsgBox("are you sure?", vbYesNo, "exit program:") = vbNo Then
                Cancel = True
        Else
                MsgBox "Goodbye!"
        End If
End Sub


