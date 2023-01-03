# Driver-COM-TaskPane
Driver COM TaskPane For Ms Office

Sử dụng  cho  Ms Office :  Excel - Word - Access - PowerPoint - Outlook

1/ tạo 1 UseerControl trên bất cứ ngôn ngữ nào có hổ trợ COM

2/ sử dụng code như sau

Rem ------------------------------------------------------------------
Dim TP As Object, CTP As Object
Rem ------------------------------------------------------------------
Sub ShowTaskPane()
    Set CTP = CreateObject("MyTaskPane.cTaskPane")
    Set TP = CTP.CreateCTP("MyTaskPane.TaskPane", "My Caption")
    Rem =========
    TP.DockPosition = msoCTPDockPositionRight       ''ben phai
    Rem TP.DockPosition = msoCTPDockPositionLeft    ''ben trai
    Rem TP.DockPosition = msoBarRight               ''ben phai
    Rem TP.DockPosition = msoBarLeft                ''ben trai
    Rem =========
    TP.DockPositionRestrict = msoCTPDockPositionRestrictNone
    Rem TP.DockPositionRestrict = msoCTPDockPositionRestrictNoChange     ''Khong cho thay doi keo Tha
    TP.Visible = True
    Rem =========
    Set CTP = Nothing
    Set TP = Nothing
End Sub
Rem ------------------------------------------------------------------
Sub ShowTaskPane2()
    Set CTP = CreateObject("MyTaskPane.cTaskPane")
    With CTP.CreateCTP("MyTaskPane.TaskPane", "My Caption")
        .DockPosition = msoCTPDockPositionRight
        .Visible = True
   End With
   Set CTP = Nothing
End Sub
Rem ------------------------------------------------------------------

