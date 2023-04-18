'Author: Rob Lawton
'Date of release: 28/03/2017
'Update: v2.8 - 15/04/2023
'Rel. candidate: Yr_Wyddfa
'Purpose: Take a large file and split into multiple smaller differing sized files to both circumvent mail attachment file type policy blocking and email attachment size policy.
'Usage: Pass no params for help text.  
'Note: Execute VBScript using console(cscript) or Win(wscript).
'Addition: If you do find this script useful or any sections of code within it, -
'		   - feel free to distribute however all I ask is that you keep the - 
'          - author details present.
'
'Version: v1.0 - Release
'Version: v1.1 - Fixxed issue, wrapped OS hosted WScript.Shell in quotes to deal with filenames containing spaces
'Version: v1.2 - Fixed issue with explicit close of class after consumption.
'Version: v2.0 - Integrated split/join functions into endecoder. 
'Version: v2.1 - Function fcnUpdatedParFiles rewritten, originally used Scripting.FileSystemObject class - 
'		  	   - and a bubble sort for par file access in date created order. Now leverages host OS efficiencies using WScript.Shell.
'Version: v2.2 - Moved mandatory file/folder checks to start
'Version: v2.3 - Created a temp file based on original filename and added check for deletion.
'Version: v2.4 - Moved delete and check to a function to remove duplication.
'Version: v2.5 - Added original encoded source filename to .par file name as hint.
'Version: v2.6 - Added more error detection and verbose output to assist in bad parameter values being passed.
'Version: v2.7 - Added removal of tmp b64 file created when encoding after split
'Version: v2.8 - Added corruption feature to enable support for emailing banned file extension types (for Windows mainly), where LineFeed is removed but magically replaced when CarriageReturn is detected without LineFeed during decode.
'--------------------------------

Option Explicit
'main
dim arguments,vbTab, fileListVar,oShell, strHelpInfo, strWarn, x, lineVar
Set arguments = WScript.Arguments
if wscript.arguments.count < 3 then
	call fcnHelper 
else
	select case lcase(arguments(0))
	case "d"
		if wscript.arguments.count < 3 then
			call fcnHelper
		else
			if fcnCheckFolder(arguments(1)) then  'look for .par files in this folder
				wscript.echo "Verified that [" & arguments(1) & "] source par folder exists."
				dim dfileVar, dfolderVar
				dfolderVar=arguments(1)
				dfileVar=arguments(2)
				if right(dfolderVar,1) <> "\" then dfolderVar=dfolderVar & "\"
					if not fcnCheckFile(dfileVar & "-tmp.b64") then
						if not fcnCheckFile(dfileVar)then 
							if fcnUpdatedParFiles(dfolderVar,dfileVar) then 
								wscript.echo "Decode: [" & dfileVar & "-tmp.b64] to: [" & dfileVar & "]."
								call fDecodeB64(dfileVar & "-tmp.b64",dfileVar)
								if fcnCheckFile(dfileVar) then
									if fcnDelfileForSure(dfileVar & "-tmp.b64") then
										wscript.echo "Removed the temp file: [" & dfileVar & "-tmp.b64]." 
									else
										wscript.echo "Destination file created, however temporary file could not be removed: [" & dfileVar & "-tmp.b64]."
									end if
									wscript.echo "Your original par files have been left in: [" & dfolderVar & "] perhaps clean-up if you're done." &_
									vbcrlf & vbcrlf & "Done! See file created:--> [" & dfileVar & "] <--"
								else
									wscript.echo "File creation failed."
								end if
							else
								wscript.echo "Problems detected, halting!"
							end if
					else
						wscript.echo "[" & dfileVar & "] already exists - halting!  Have you ran this before? Remove it or it's already done."
					end if
				else
					wscript.echo "Temporary file [" & dfileVar & "-tmp.b64] already exists - halting! Have you ran this before? Clean your directories. Try another destination folder or do some spring cleaning."
				end if
			else
				wscript.echo ".par files cannot be located in supplied location: [" & arguments(1) & "]." & vbcrlf &_
				"If you are attempting to decode a single file, just rename to a filename with the .par extension and use a name that matches no other in that folder, example: " & chr(34) & "afile1.par" & chr(34) & "."
			end if 
		end if
	case "e" 
		if wscript.arguments.count < 5 then  
			call fcnHelper
		else
			dim minsizeVar,maxsizeVar,filetosplitVar,folderoutVar
			filetosplitVar=arguments(1)
			folderoutVar=arguments(2)
			minsizeVar=clng(arguments(3)*1000)
			maxsizeVar=clng(arguments(4)*1000)
			if (isnumeric(minsizeVar) and isnumeric(maxsizeVar) and (maxsizeVar > minsizeVar)) then
				if fcnCheckFile(filetosplitVar) then
					wscript.echo "Verified: [" & filetosplitVar & "] exists."
					if fcnCheckFolder(folderoutVar) then
						wscript.echo "Verified: [" & folderoutVar & "] exists."
						if right(folderoutVar,1) <> "\" then folderoutVar=folderoutVar & "\"						
						dim sourcepathVar,filenameVar,filenameVar2,fileextVar
						call fcnFileFolderSplit("\",filetosplitVar,sourcepathVar,filenameVar)
						if right(sourcepathVar,1) <> "\" then sourcepathVar=sourcepathVar & "\"
						call fcnFileFolderSplit(".",filenameVar,filenameVar2,fileextVar)
						wscript.echo "Decode: [" & filetosplitVar & "]." & vbcrlf &_
						"Out to path: [" & folderoutVar & "]." & vbcrlf &_
						"Out to file: [" & filenameVar2 & "-" & fileextVar & ".b64]."
						if fcnCheckFile(folderoutVar & filenameVar2 & "-" & fileextVar & ".b64") then
							wscript.echo "File already exists: [" & folderoutVar & filenameVar2 & "-" & fileextVar & ".b64].  Have you already executed me? Try a tidy-up."
						else
							call fEncodeB64(filetosplitVar,folderoutVar & filenameVar2 & "-" & fileextVar & ".b64")
							if fcnCheckFile(folderoutVar & filenameVar2 & "-" & fileextVar & ".b64") then
								wscript.echo "Success, now to split the file."
								dim returnMsg
								if fcnCheapAreTherePacFiles(folderoutVar, folderoutVar & filenameVar2 & "-" & fileextVar & ".b64", returnMsg) then
									wscript.echo returnMsg & " - There are already par files in the folder: [" & folderoutVar & "], have a tidy-up or use a different location."
								else
									dim ThisFilelistArray()
									if fcnSplitUp(folderoutVar & filenameVar2 & "-" & fileextVar & ".b64",minsizeVar,maxsizeVar,ThisFilelistArray)	then
										wscript.echo vbcrlf & "Par files generated and ready for you to send:"
										for x=0 to ubound(ThisFilelistArray)
											wscript.echo "File " & x+1 & ": --> [" & ThisFilelistArray(x) & "]."
										next
									wscript.echo "Removing tmp b64 file: [" & folderoutVar & filenameVar2 & "-" & fileextVar & ".b64]."
									if fcnDelfileForSure(folderoutVar & filenameVar2 & "-" & fileextVar & ".b64") Then
										wscript.echo "Removed tmp b64 file: [" & folderoutVar & filenameVar2 & "-" & fileextVar & ".b64]."
									Else
										wscript.echo "Unable to remove tmp b64 file: [" & folderoutVar & filenameVar2 & "-" & fileextVar & ".b64]."
									end if
									else
										wscript.echo "There was an error detected in par file generated."
									end if
								end if
							else
								wscript.echo "File not created: [" & folderoutVar & filenameVar2 & ".b64]."
							end if
						end if
					else
						wscript.echo "Cannot access the folder: [" & folderoutVar & "]."
					end if
				else
					wscript.echo "Filename supplied cannot be accessed: [" & filetosplitVar & "]."
				end if
			else
				wscript.echo "Maximum and minimum file size parameters supplied need to be numerical and the maximum value larger than the minimum value." & vbcrlf &_
				"Try executing without any parameters supplied to display the help text."
			end if
		end if
	case else
		call fcnHelper
	end select
end if
'--------------------------------
Function fDecodeB64(byref B64FileIn,BinFileOut)
	Dim FSys
	Dim InFile
	Dim DataStreamIn
	Dim XMLRefob
	Dim Elem
	Dim StreamX
	Set FSys = CreateObject("Scripting.FileSystemObject")
	Set InFile  = FSys.GetFile(B64FileIn)
	Set DataStreamIn = InFile.OpenAsTextStream(1, 0)
	Set XMLRefob = CreateObject("MSXml2.DOMDocument")
	Set Elem = XMLRefob.createElement("Base64Data")
	Elem.DataType = "bin.base64"
	Elem.text = DataStreamIn.ReadAll()
	Set StreamX = CreateObject("ADODB.Stream")
	StreamX.Type = 1
	StreamX.Open()
	StreamX.Write Elem.NodeTypedValue
	StreamX.SaveToFile BinFileOut, 2
	if not FSys is nothing then Set FSys = Nothing
	if not InFile is nothing then Set InFile = Nothing
	if not DataStreamIn is nothing then Set DataStreamIn = Nothing
	if not XMLRefob is nothing then Set XMLRefob = Nothing
	if not Elem is nothing then Set Elem = Nothing
	if not StreamX is nothing then Set StreamX = Nothing
End Function
'--------------------------------
Function fEncodeB64(byref BinFileIn,B64FileOut)
	Dim inputStream
	Set inputStream = CreateObject("ADODB.Stream")
	inputStream.Open
	inputStream.Type = 1  
	inputStream.LoadFromFile BinFileIn
	Dim bytes: bytes = inputStream.Read
	Dim dom: Set dom = CreateObject("Microsoft.XMLDOM")
	Dim elem: Set elem = dom.createElement("tmp")
	elem.dataType = "bin.base64"
	elem.nodeTypedValue = bytes
	dim FSys
	Set FSys = CreateObject("Scripting.FileSystemObject")
	dim OFile
	set OFile=FSys.CreateTextFile(B64FileOut,True)
	OFile.writeline Replace(elem.text, vbLf, "") 'cheeky 
	OFile.close
	if not inputStream is Nothing then set inputStream=Nothing
	if not dom is Nothing then set dom=Nothing
	if not elem is Nothing then set elem=Nothing
	if not FSys is Nothing then set FSys=Nothing
	if not OFile is Nothing then set OFile=Nothing
End Function
'--------------------------------
Function fcnUpdatedParFiles (byref FolderIn,FileOut)
	fcnUpdatedParFiles=False
	if right(FolderIn,1) <> "\" then FolderIn=FolderIn & "\"
	wscript.echo "Will look in: [" & FolderIn & "] for par files. Will decode out to supplied file: [" & FileOut & "]."
	wscript.echo "Great, temp file: [" & FileOut & "-tmp.b64] doesn't exist, which is a good thing, creating file."
	dim returnMsg
	if fcnCheapAreTherePacFiles(FolderIn, FileOut, returnMsg) then
		wscript.echo "Have located par files to process, generating: [" & FileOut & "-tmp.b64]."
		dim objShell, shellCmd
		set objShell = WScript.CreateObject ("WScript.Shell")
		shellCmd="cmd /c for /f " & chr(34) & "tokens=*" & chr(34) & " %A in ('dir /on /b " & chr(34) & FolderIn & "*.par" & chr(34) & "') do type " & chr(34) & FolderIn & "%A" & chr(34) & " >> " & chr(34) & FileOut & "-tmp.b64" & chr(34)
		objShell.run shellCmd,0,true
		if not objShell is Nothing then set objShell=Nothing
		if fcnCheckFile(FileOut & "-tmp.b64") then
			wscript.echo "Output file created - [" & FileOut & "-tmp.b64]."
			fcnUpdatedParFiles=True
		else
			wscript.echo returnMsg & " - Output file was not created - " & FileOut & "-tmp.b64."
		end if
	else
		wscript.echo returnMsg & " - No par files located or error detected in par file check function."
	end if
End Function
'--------------------------------
Function fcnCheckFile(byref FileIn)
	Dim FSys
	Set FSys = CreateObject("Scripting.FileSystemObject")
	if instr(1,FileIn,"\",1)=0 then wscript.echo "No folder name was supplied, so all I can do is assume current folder and check: " & FSys.GetAbsolutePathName(".")
	fcnCheckFile=false
	if FSys.FileExists(FileIn) then	fcnCheckFile=True
	if not FSys is Nothing then set FSys=Nothing
End Function
'--------------------------------
Function fcnCheckFolder(byref FolderIn)
	Dim FSys
	Set FSys = CreateObject("Scripting.FileSystemObject")
	fcnCheckFolder=False
	if FSys.FolderExists(FolderIn) then	fcnCheckFolder=True
	if not FSys is Nothing then set FSys=Nothing
End Function
'--------------------------------
Function fcnCheapAreTherePacFiles(byref FolderIn, FileOut, returnMsg)
	dim objShell,shellCmd
	fcnCheapAreTherePacFiles=False
	set objShell = WScript.CreateObject ("WScript.Shell")
	shellCmd="cmd /c set count=0 & for /f " & chr(34) & "tokens=*" & chr(34) & " %A in ('dir /on /b " & chr(34) & FolderIn & "*.par" & chr(34) & "') do @if %errorlevel% EQU 0 @echo %A >> " & chr(34) & FileOut & "-tmp.b64" & chr(34) 
	returnMsg="Executing shell command"
	objShell.run shellCmd,0,true
	if fcnCheckFile(FileOut & "-tmp.b64") then
		returnMsg="Found par files, removing tmp file used."
		if fcnDelfileForSure(FileOut & "-tmp.b64") then
			fcnCheapAreTherePacFiles=True
			returnMsg="Par files exist and tmp file used removed."
		else
			returnMsg="Par files exist, however tmp file could not be removed - halting."
		end if
	else
		returnMsg="No par files located."
	end if
	if not objShell is Nothing then set objShell=Nothing
End Function
'--------------------------------
Function fcnDelfileForSure(fileIn)
	fcnDelfileForSure=False
	wscript.echo "Deleting File: [" & fileIn & "]."
	dim objShell,shellCmd
	set objShell = WScript.CreateObject ("WScript.Shell")
	shellCmd="cmd /c @del /q /f " & chr(34) & fileIn & chr(34)
	objShell.run shellCmd,0,true
	if not objShell is Nothing then set objShell=Nothing
	wscript.echo "Checking File: [" & fileIn & "] has been removed."
	if not fcnCheckFile(fileIn) then fcnDelfileForSure=True
End Function
'--------------------------------
Function fcnFileFolderSplit(byref SplitVal,DataIn,FileFolderOut,FileExtOut)
	dim splitArray,x
	splitArray=split(DataIn,SplitVal)
	if ubound(splitArray)=0 then
		FileFolderOut=FileFolderOut
		FileExtOut=vbnullstring
	else
		for x = 0 to ubound(splitArray)-1
			FileFolderOut=FileFolderOut & SplitVal & splitArray(x)
		next
	FileFolderOut=mid(FileFolderOut,2,len(FileFolderOut))
	FileExtOut=splitArray(ubound(splitArray))
	end if
	
End Function
'--------------------------------
Function fcnSplitUp(byref FileSupplied,MinFileSize,MaxFileSize,OutputfileList)
	fcnSplitUp=False
	wscript.echo "FileSupplied: [" & FileSupplied & "]." & vbcrlf &_
	"Minimum file size: [" & MinFileSize & " bytes]." & vbcrlf &_
	"Maximum file size: [" & MaxFileSize & " bytes]."
	dim OutPutfileListCount
	dim objFso,objFile,FileSize,DataStreamIn,RndListUsed,ArrayM,RandVal,TotalCharCount,DataStreamOut,TotalFileCount,DataRemain
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.getFile(FileSupplied)
	FileSize=objFile.size
	Set DataStreamIn = objFile.OpenAsTextStream()
	do while not DataStreamIn.atendofstream
		RndListUsed=fcnRndFileSizer(RndListUsed,MinFileSize,MaxFileSize)
		if instr(RndListUsed,"ERROR!") > 0 then exit do
		ArrayM=split(RndListUsed, ",")
		RandVal=ArrayM(ubound(ArrayM))
		if clng(RandVal) > clng(clng(FileSize)-clng(TotalCharCount)) then	
			wscript.echo "The rand size generated is larger than the amount of bytes remaining to process." & vbcrlf & "Rand value will use bytes remaining as file size."
			RandVal=clng(clng(FileSize)-clng(TotalCharCount))
		end if
		TotalCharCount=clng(TotalCharCount+RandVal)
		DataStreamOut=DataStreamIn.read(RandVal)
		TotalFileCount=TotalFileCount+1
		DataRemain=clng(clng(FileSize)-clng(TotalCharCount))
		redim preserve OutputfileList(OutPutfileListCount)
		OutputfileList(OutPutfileListCount)=fcnFileOutIt(FileSupplied,DataStreamOut,TotalFileCount,DataRemain,RandVal)
		OutPutfileListCount=OutPutfileListCount+1
		DataStreamOut=vbNullString	
	loop
	if not objFso is nothing then set objFso=nothing
	if not objFile is nothing then set objFile=nothing
	DataStreamIn.close
	wscript.echo "Checking files generated exist..." 
	dim CheckMsg,CheckResult,PassVar
	for x=0 to ubound(OutputfileList)
		CheckResult="Failed"
		CheckMsg="Checking: [" & OutputfileList(x) & "]"
		if fcnCheckFile(OutputfileList(x)) then 
			CheckResult="Exists"
			PassVar=PassVar+1
		end if
		wscript.echo CheckMsg & " - " & CheckResult
	next
	if ubound(OutputfileList)+1 = PassVar then 
		wscript.echo PassVar & " files passed check"
		fcnSplitUp=True
	end if
End Function
'--------------------------------
Function fcnFileOutIt(byref Filename,DataOut,FileOutCount,LeftToDo,RndSize)
	dim ParFilenameout,ParFolderOut,extVar
	call fcnFileFolderSplit(".",Filename,ParFilenameout,extVar)
	call fcnFileFolderSplit("/",Filename,ParFolderOut,extVar)
	dim OutFile,outfilename,objFso
	Set objFso = CreateObject("Scripting.FileSystemObject")
	outfilename=vbnullstring
	outfilename=ParFolderOut & ParFilenameout & FileOutCount & ".par"
	Set OutFile = objFso.CreateTextFile(outfilename,True)
	OutFile.write DataOut
	OutFile.close
	if not OutFile is nothing then set OutFile=nothing
	if not objFso is nothing then set objFso=nothing
	wscript.echo "File No." & FileOutCount & " = " & outfilename & ". Size:" & RndSize & " Bytes. " & LeftToDo & " Bytes remaining."
	fcnFileOutIt=outfilename
	End Function
'--------------------------------
Function fcnRndFileSizer(ListUsedin,minsizein,maxsizein)
	dim arrayn
	dim GenVal
	dim UniqueVal
	dim tries
	dim oar
	arrayn = split(ListUsedin, ",")
	if ubound(arrayn) < 0 then	
		fcnRndFileSizer=sSubfRndFileSizer(minsizein,maxsizein) 
	else
		UniqueVal=0
		do until UniqueVal=1 or tries=100
			tries=tries+1
			GenVal=sSubfRndFileSizer(minsizein,maxsizein)
			for oar=lbound(arrayn) to ubound(arrayn)
				if int(GenVal)=int(arrayn(oar))then 
					UniqueVal=0
					exit for	
				end if
				if not int(GenVal)=int(arrayn(oar)) and int(oar)=int(ubound(arrayn)) then 
					UniqueVal=1	
				end if
			next
		Loop
		if tries=100 then 
			wscript.echo "You have not generated a wide enough range to cover the number of random numbers needed!"
			fRndFileSizer="ERROR!"
		else
			fcnRndFileSizer=ListUsedIn & "," & GenVal
		end if
	end if
End Function
'--------------------------------
Function sSubfRndFileSizer(minin,maxin)
		randomize
		sSubfRndFileSizer=Int((maxin-minin+1)*Rnd+minin)
End Function
'--------------------------------
Function fcnHelper
	vbTab="  "
	for x=1 to 50
		lineVar=LineVar & chr(45)
	next
	wscript.echo vbcrlf & lineVar & ">" & vbcrlf & "Supply parameters (not case sensitive) as follows:  " & vbcrlf & vbTab & chr(34) & "D" &  chr(34) & " to Decode and join  ~or~ " & chr(34) & "E" & chr(34) & " to Encode and split." & vbcrlf &_
	vbcrlf & vbcrlf & "For the selected option, supply the  additional parameters as detailed below:" & vbcrlf & vbcrlf & vbcrlf &_
	vbTab & "(D) Decode and split - A bunch of files you've received that need to be assembled." & vbcrlf &_
	vbTab & vbTab & "You need to supply:" & vbcrlf &_
	vbTab & vbTab & vbTab & "(1) - The location of the filenames received." & vbcrlf &_
	vbTab & vbTab & vbTab & "(2) - The path and filename to be assembled upon rebuilding the files received." &_
	vbcrlf & vbtab & vbtab & vbTab & vbTab & vbTab & "Hint, if the originally created part ('.par') files have not been renamed, the original [filename-ext] should proceed the '.b64'." & vbcrlf &_
	vbTab & vbTab & vbTab & vbTab & vbTab & "For example: " & chr(34) & "image-jpg1.par" & chr(34) & " would be the first file in a set for the file " & chr(34) & "image.jpg" & chr(34) & vbcrlf &_
	vbTab & vbTab & vbTab & vbTab & vbTab & "Surround files / folder names containing spaces in " & chr(34) & "quotes" & chr(34) & "." & vbcrlf &_
	vbTab & vbTab & vbTab & vbTab & vbTab & "Note, any part ('.par') files found will be consumed, so keep your folders tidy." & vbcrlf &_
	vbcrlf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "--> Example: " & Wscript.ScriptName & " D " & chr(34) & "c:\users\bob\documents\email attachments" & chr(34) & " " & chr(34) & "c:\temp\That Really Big File.pdf" & chr(34) & vbcrlf & vbcrlf &_
	vbTab & vbcrlf & vbtab & "(E) Encode and join - A large file that you need to split-up to send via email." & vbcrlf &_
	vbTab & vbTab & "You need to supply:" & vbcrlf &_
	vbTab & vbTab & vbTab & "(1) - The location of the file to be split," & vbcrlf &_
	vbTab & vbTab & vbTab & "(2) - The location where you wish to place the smaller files," & vbcrlf &_
	vbTab & vbTab & vbTab & "(3) - The minimum file size to generate (in bytes)," & vbcrlf &_
	vbTab & vbTab & vbTab & "(4) - The maximum file size to generate (in bytes)." & vbcrlf &_
	vbTab & vbTab & vbTab & vbTab & vbTab & "Be sure to surround files / folder names containing spaces in " & chr(34) & "quotes" & chr(34) & "." & vbcrlf &_
	vbcrlf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "--> Example: " & WScript.ScriptName & " E " & chr(34) & "c:\bobs folder\mail files\Big File.pdf" & chr(34) & " " & chr(34)& "c:\temp\my destination folder\Email Attachments"  & chr(34) & " 9000 10000" & vbcrlf &_
	vbcrlf & vbTab & vbTab & vbTab & "*Note, please consider that executing this utility against a file that is situated on a NAS/SAN or -" &_
	vbcrlf & vbTab & vbTab & vbTab & "- equivalent network file share will result in a longer duration of execution, but it will still do it!" & vbcrlf & "<" & lineVar
End Function
