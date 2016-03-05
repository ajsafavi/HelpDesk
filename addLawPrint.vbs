Option Explicit
On Error Resume Next
Dim objPrinter
Set objPrinter = WScript.CreateObject("WScript.Network")

objPrinter.AddWindowsPrinterConnection "\\lawprint.law.berkeley.edu\325C-Xerox-7556C"
objPrinter.AddWindowsPrinterConnection "\\lawprint.law.berkeley.edu\325A-M601dn-GoldenGate"
objPrinter.AddWindowsPrinterConnection "\\lawprint.law.berkeley.edu\325A-M601dn-Fog"
objPrinter.AddWindowsPrinterConnection "\\lawprint.law.berkeley.edu\325-M601dn-Jake"
objPrinter.AddWindowsPrinterConnection "\\lawprint.law.berkeley.edu\325-M601dn-Reception"