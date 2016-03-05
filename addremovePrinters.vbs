Option Explicit
On Error Resume Next
 
Dim objNetwork
Dim oPrinters
Dim intCounter
Dim strAllPrinters

Set objNetwork = CreateObject("WScript.Network")
 
Dim printServer 
printServer = "\\print"
Dim lawprintServer 
lawprintServer = "\\lawprint"
Dim lawprint2Server 
lawprint2Server = "\\lawprint.law.berkeley.edu"

 
'Create list of all printers
Set oPrinters = objNetwork.EnumPrinterConnections
Dim printerName
Dim shortPrinterName
Dim newPrinterName
For intCounter = 0 To oPrinters.Count - 1 Step 2
	printerName = oPrinters.Item(intCounter + 1)
	'bad printers, remove them'
	If InStr(printerName, lawprintServer) Or InStr(printerName, lawprint2Server) > 0 Then
		If InStr(printerName, lawprint2Server) > 0 Then
			shortPrinterName = Mid(printerName,Len(lawprint2Server)+1)
		ElseIf InStr(printerName, lawprintServer) > 0 Then
			shortPrinterName = Mid(printerName, Len(lawprintServer)+1)
		End If 
		newPrinterName = printServer & shortPrinterName
		'Wscript.Echo printerName
		'Wscript.Echo shortPrinterName
		'Wscript.Echo newPrinterName
		
		objNetwork.RemovePrinterConnection printerName,true,true
		objNetwork.AddWindowsPrinterConnection  newPrinterName
	End If
	'strAllPrinters = strAllPrinters & objPrinters.Item(LOOP_COUNTER + 1) & VbCrLf   
Next
 
 