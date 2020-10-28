' PDFCreator COM Interface test for VBScript
' Part of the PDFCreator application
' License: GPL
' Homepage: http://www.pdfforge.org/pdfcreator
' Version: 1.1.0.0
' Created: June, 16. 2015
' Modified: May, 20. 2020 
' Author: pdfforge GmbH
' Comments: This project demonstrates the use of the COM Interface of PDFCreator.
'           This script converts a windows testpage to a .pdf file.
' Note: More usage examples then in the VBScript directory can be found in the JavaScript directory only.

Dim ShellObj, PDFCreatorQueue, scriptName, strInputFileName, strDestFileName, printJob, objFSO, tmp
Const TemporaryFolder = 2
 
if (WScript.Version < 5.6) then
    MsgBox "You need the Windows Scripting Host version 5.6 or greater!"
    WScript.Quit
end if

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set ShellObj = CreateObject("Shell.Application")
Set PDFCreatorQueue = CreateObject("PDFCreator.JobQueue")
Set PDFCreator = CreateObject("PDFCreator.PdfCreatorObj")

strInputFileName = objFSO.GetAbsolutePathName("sample1.txt") 
strDestFileName = objFSO.GetAbsolutePathName("printer_details_pdf.pdf")

MsgBox "Initializing PDFCreator queue..."
PDFCreatorQueue.Initialize

' ' Task: Creating the content of PDF using Printer Tester Details
' MsgBox "Printing a windows testpage"
' ShellObj.ShellExecute "RUNDLL32.exe", "PRINTUI.DLL,PrintUIEntry /k /n ""PDFCreator""", "", "open", 1

' ' Task: Creating the content of PDF using an existing TXT file
If Not objFSO.FileExists(strInputFileName) Then
    MsgBox "PDFCreator: Can't find the file: " & strInputFileName
Else 
    MsgBox "Printing Page: " & strInputFileName 
    PDFCreator.PrintFile strInputFileName
    WScript.Sleep 1000
End If

MsgBox "Waiting for the job to arrive at the queue..."
if not PDFCreatorQueue.WaitForJob(10) then
    MsgBox "The print job did not reach the queue within " & " 10 seconds"
else 
    MsgBox "Currently there are " & PDFCreatorQueue.Count & " job(s) in the queue"
    MsgBox "Getting job instance"
    Set printJob = PDFCreatorQueue.NextJob
    
    ' printJob.SetProfileByGuid("DefaultGuid")
    printJob.SetProfileSetting "PdfSettings.PageOrientation", "Landscape"
    printJob.ConvertTo(strDestFileName)
    
    if (not printJob.IsFinished or not printJob.IsSuccessful) then
		MsgBox "Could not convert the file: " & strDestFileName
	else
        MsgBox "File created: " & strDestFileName
		MsgBox "Job finished successfully"
    end if
end if

MsgBox "Releasing the object"
PDFCreatorQueue.ReleaseCom