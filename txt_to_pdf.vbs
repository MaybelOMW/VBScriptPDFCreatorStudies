Dim strExt, intStatus, strDestFileName, strInputFileName, strReason 

Const ForReading = 1 
Const ForWriting = 2
Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")
Set PDFCreatorQueue = CreateObject("PDFCreator.JobQueue")
Set PDFCreator = CreateObject("PDFCreator.PdfCreatorObj")

strInputFileName = objFSO.GetAbsolutePathName("sample1.txt") 
strDestFileName = objFSO.GetAbsolutePathName("txt_to_pdf.pdf")
strLogFileName = objFSO.GetAbsolutePathName("txt_to_pdf.log")

PDFProcess 

' ** Sub Routine to render file as PDF
Sub PDFProcess  
	Dim objFolder, job, intStatPDFCreator, intPageCount
	
	intPageCount = 1  
	
	Loggit "PDF Destination Name: " & strDestFileName 
    Loggit "Initializing PDFCreator queue..."
	intStatPDFCreator = PDFCreatorQueue.Initialize()
	Loggit "PDFCreator Object Status: " & intStatPDFCreator 
	
	If intStatPDFCreator = 0 Then 
			If Not objFSO.FileExists(strInputFileName) Then
				Loggit "PDFCreator: Can't find the file: " & strInputFileName
			Else 
				Loggit "Printing Page: " & strInputFileName 
				
				PDFCreator.PrintFile strInputFileName
				
				WScript.Sleep 1000
				Loggit "Currently there are " & PDFCreatorQueue.Count & " job(s) in the queue"
			End If
		
		Loggit "Waiting for the job to arrive at the queue..."
		if Not(PDFCreatorQueue.WaitForJobs(intPageCount, 100)) Then 
			strReason = "The print job did not reach the queue within " & 10 & " seconds" 
			Loggit strReason 
			intStatus = 0
		Else
			Loggit "Currently there are " & PDFCreatorQueue.Count & " job(s) in the queue" 
			Loggit "Getting job instance and merging"
			
			PDFCreatorQueue.MergeAllJobs
		
			while(PDFCreatorQueue.Count > 0)
				Set job = PDFCreatorQueue.NextJob
					Loggit "Staging PDF File: " & strDestFileName 
				job.SetProfileSetting "PdfSettings.PageOrientation", "Landscape"
				job.ConvertTo(strDestFileName)
					WScript.sleep 5000
				
				If Not(job.IsFinished Or job.IsSuccessful) Then
					strReason = "Could not convert the file: " & strDestFileName
						Loggit strReason 
					intStatus = 0
				Else
					Loggit "Job finished successfully" 
				End If 
			Wend 
		End If 
			Loggit "Releasing the object"
		PDFCreatorQueue.ReleaseCom()
	Else
		strReason =  "Failed to create PDFCreator COM instance."
			Loggit strReason 
		intStatus = 0
	End If


End Sub 

' ** Sub Routine for Logging to Text File
Sub Loggit(msg) 
	Dim stream 
    Set stream = objFSO.OpenTextFile(strLogFileName , 8, True)
		stream.writeline Date & " " & Time & ": " & msg
		WScript.Echo Date & " " & Time & ": " & msg
		stream.close
End Sub 