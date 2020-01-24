' Create Shortcut for Total Recall Test Pilot
' Written by: Mario Steele <mario@ruby-im.net>
' Written: 01/24/2020

Set Shell = CreateObject("WScript.Shell")
DesktopPath = Shell.SpecialFolders("Desktop")
Set link = Shell.CreateShortcut(DesktopPath & "\TotalRecall Test Pilot.lnk")
link.Arguments = """https://remote.storage-mart.com"""
link.Description = "Test Pilot for new Total Recall Login"
link.TargetPath = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
link.WindowStyle = 3
link.Save
MsgBox("Total Recall Test Pilot Link created.")
