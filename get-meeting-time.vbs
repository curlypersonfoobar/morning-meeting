On Error Resume Next
Set objOutlook = CreateObject("Outlook.Application")
If Err.Number <> 0 Then
    WScript.Echo "COM_ERROR: " & Err.Description
    WScript.Quit 1
End If

Set objNamespace = objOutlook.GetNameSpace("MAPI")
Set objCalendar = objNamespace.GetDefaultFolder(9)

dtmTomorrow = DateAdd("d", 1, Date)
dtmStart = DateSerial(Year(dtmTomorrow), Month(dtmTomorrow), Day(dtmTomorrow))
dtmEnd = DateAdd("s", -1, DateAdd("d", 1, dtmStart))

Set objItems = objCalendar.Items
objItems.IncludeRecurrences = True
objItems.Sort "[Start]"

strFilter = "[Start] >= '" & dtmStart & "' AND [Start] <= '" & dtmEnd & "'"
Set objFiltered = objItems.Restrict(strFilter)

If objFiltered.Count > 0 Then
    WScript.Echo FormatDateTime(objFiltered.Item(1).Start, vbShortTime)
Else
    WScript.Echo "no meeting"
End If