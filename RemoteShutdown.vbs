	Dim oFSO, oTS, sClient, oWindows, oLocator, oConnection, oSys
	Dim sUser, sPassword

	'set remote credentials
	'Need administrative account for the credentials below
	sUser	    = "username"
	sPassword = "password"

	'open list of client names
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oTS = oFSO.OpenTextFile("C:\pcs.txt")
	Set CMDShell = wscript.createobject("wscript.shell")

	Do Until oTS.AtEndOfStream
		'get next client name
		sClient = oTS.ReadLine

		CMDShell.run("%comspec% /k ping " & sClient & " -t")

		'get WMI locator
		Set oLocator = CreateObject("WbemScripting.SWbemLocator")

		'Connect to remote WMI
		Set oConnection = oLocator.ConnectServer(sClient, "root\cimv2", sUser, sPassword)

		'issue shutdown to OS
		' 4 = force logoff
		' 5 = force shutdown
		' 6 = force reboot
		' 12 = force power off
		Set oWindows = oConnection.ExecQuery("Select " & "Name From Win32_OperatingSystem")

		For Each oSys In oWindows
			oSys.Win32ShutDown(6)
		Next

		wscript.sleep 10000	' 1min * 10secs * 1000ms = 10,000ms
	Loop

	'close the text file
	oTS.Close
	msgbox "Done."

