Attribute VB_Name = "modSetting"
Option Explicit

Private Type ExternalApplicationType
    Viewer As String
    Editor As String
    Printer As String
End Type

Public Type SettingType
    ExternalApplication As ExternalApplicationType
End Type

Public Function SettingLoad() As SettingType
    SettingLoad.ExternalApplication.Viewer = INIGetString("Setting", "External application- Viewer")
    SettingLoad.ExternalApplication.Editor = INIGetString("Setting", "External application- Editor")
    SettingLoad.ExternalApplication.Printer = INIGetString("Setting", "External application- Printer")
End Function

Public Sub SettingSave(New_Setting As SettingType)
    With New_Setting
        INISet "Setting", "External application- Viewer", .ExternalApplication.Viewer
        INISet "Setting", "External application- Editor", .ExternalApplication.Editor
        INISet "Setting", "External application- Printer", .ExternalApplication.Printer
    End With
    
    With ApplicationSetting
        .ExternalApplication.Viewer = New_Setting.ExternalApplication.Viewer
        .ExternalApplication.Editor = New_Setting.ExternalApplication.Editor
        .ExternalApplication.Printer = New_Setting.ExternalApplication.Printer
        
        frmBrowser.ctlThumbNailList.ExternalViewer = .ExternalApplication.Viewer
        frmBrowser.ctlThumbNailList.ExternalEditor = .ExternalApplication.Editor
        frmBrowser.ctlThumbNailList.ExternalPrinter = .ExternalApplication.Printer
    End With
End Sub
