Attribute VB_Name = "modImageBrowser"
Option Explicit

Public ApplicationSetting As SettingType

Sub Main()
    ApplicationSetting = SettingLoad
    
    frmBrowser.Show
End Sub

