Option Explicit

' MISCsellaneous support functions:
'   gsMISCIniPutInt      -   Save a Int in your Ini
'   gfMISCIniGetInt      -   Get a Int from your Ini
'   gsMISCAboutLoad      -   Funtion you call to setup About.FRM
'   gsMISCAboutUnload    -   Function called by About.FRM

' NOTE: This module DEPENDS ON an Ini control called
'       "Ini" on your main form called "Main"

Function gfMISCIniGetInt (sEntry As String, lValue As Long) As Long
Dim lTemp As Long

    Main!Ini.Entry = sEntry
    lTemp = Val(Main!Ini.Value)
    If lTemp = 0 Then
        Main!Ini.Value = lValue
        lTemp = lValue
    End If
    gfMISCIniGetInt = lTemp

End Function

Sub gsMISCAboutLoad (sAppCaption As String, sAppTitle As String, sAppDesc As String)

    Load About

    About.Caption = sAppCaption
    About!PanelAppTitle = sAppTitle
    About!LabelAppDescription = sAppDesc

    About.Top = Main.Top + 15
    About.Left = Main.Left + 15

    About.Top = gfMISCIniGetInt("About.Top", (About.Top))
    About.Left = gfMISCIniGetInt("About.Left", (About.Left))

    About.Show

End Sub

Sub gsMISCAboutUnload ()
    
    gsMISCIniPutInt "About.Top", (About.Top)
    gsMISCIniPutInt "About.Left", (About.Left)

    Unload About

End Sub

Sub gsMISCIniPutInt (sEntry As String, lValue As Long)

    Main!Ini.Entry = sEntry
    Main!Ini.Value = CStr(lValue)

End Sub

