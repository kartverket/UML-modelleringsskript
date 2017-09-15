option explicit

!INC Local Scripts.EAConstants-VBScript

' script:			listSKOSfraKodeliste
' description:		Skriver en kodeliste til egne SKOS-filer under samme sti som .eap-fila ligger.
' author:			Kent
' date:				2017-06-29,07-07,09-08,09-14
	DIM objFSO
	DIM outFile
	DIM objFile

	DIM codeFSO
	DIM outCodeFile
	DIM objCodeFile

sub listKoderForEnValgtKodeliste()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
	Repository.CreateOutputTab "Error"
	Repository.ClearOutput "Error"

	'Get the currently selected CodeList in the tree to work on

	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()

	if not theElement is nothing  then 
		if theElement.Type="Class" and ( LCASE(theElement.Stereotype) = "codelist" or LCASE(theElement.Stereotype) = "enumeration") or theElement.Type="Enumeration"then
			'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
			dim message
			dim box
			box = Msgbox ("List class : [«" & theElement.Stereotype &"» "& theElement.Name & "]. to SKOS/RDF/xml format."& vbCrLf & "Creates one file with all codes in the same folder as the .eap-file,"& vbCrLf & " and a subfolder with one file for each code in the list.",1)
			select case box
			case vbOK
		 		'Session.Output("Debug: ------------ Start class: [«" &theElement.Stereotype& "» " &theElement.Name& "] of type. [" &theElement.Type& "]. ")
				'inputBoxGUI to receive user input regarding the namespace
				dim namespace, nsp
				'namespace = "http://skjema.geonorge.no/SOSI/produktspesifikasjon/Stedsnavn/5.0/"
				namespace = getTaggedValue(theElement, "codeList")
				if namespace <> "" then
					nsp = Mid(namespace,Len(namespace)-Len(theElement.Name)+1,Len(theElement.Name))
					'Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
					if nsp = theElement.Name then
						'Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
						namespace = Mid(namespace,1,Len(namespace)-Len(nsp)-1)
						'Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
					end if
				end if
				if namespace = "" then
					namespace = getPackageTaggedValue(getAppSchPackage(theElement),"targetNamespace")
				end if

				namespace = InputBox("Please select the namespace name for the codelist.", "namespace", namespace)
				call listCodelistCodes(theElement,namespace)
			case VBcancel

			end select
		else
		  'Other than CodeList selected in the tree
		  MsgBox( "This script requires a CodeList class to be selected in the Project Browser." & vbCrLf & _
			"Please select a CodeList class in the Project Browser and try once more." )
		end if
		'Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
		Repository.EnsureOutputVisible "Script"
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires a CodeList class to be selected in the Project Browser." & vbCrLf & _
	  "Please select a CodeList class in the Project Browser and try again." )
	end if
end sub

sub listCodelistCodes(el,namespace)
	'Repository.WriteOutput "Script", Now & " CodeList: " & el.Name, 0
	'Repository.WriteOutput "Script", Now & " " & el.Stereotype & " " & el.Name, 0
	dim presentasjonsnavn
	'TODO: endre linjeskift i noter til blanke?
	' må vi legge på / på slutten der angitt namespace ikke ender på / ?
	' pakke inn noter som inneholder <>?
	'Repository.WriteOutput "Script", "Codelist Name: " & el.Name,0
	
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	outFile = getNCNameX(el.Name)&".rdf"
	Repository.WriteOutput "Script", Now & " outFile: " & outFile, 0
	Set objFile = objFSO.CreateTextFile(outFile,True,False)
	'  får ut 16-bits unicode ved å sette True som siste flagg i kallet over.
	if not objFSO.FolderExists(el.Name) then
		objFSO.CreateFolder el.Name
	end if

	Repository.WriteOutput "Script", "Writes Codelist Name: " & el.Name & " to file " & outfile& " and subfolder " & el.Name,0
	Repository.WriteOutput "Script", "With namespace: " & namespace,0

	objFile.Write"<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
	objFile.Write"<rdf:RDF" & vbCrLf
    objFile.Write"  xmlns:skos=""http://www.w3.org/2004/02/skos/core#""" & vbCrLf
	objFile.Write"  xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""" & vbCrLf
	objFile.Write"  xml:base="""&utf8(namespace)&"/"">" & vbCrLf


    objFile.Write"  <skos:ConceptScheme rdf:about="""&utf8(el.Name)&""">" & vbCrLf
	presentasjonsnavn = getTaggedValue(el,"SOSI_presentasjonsnavn") 
	if presentasjonsnavn = "" then presentasjonsnavn = el.Name
    objFile.Write"    <skos:prefLabel xml:lang=""no"">"&utf8(presentasjonsnavn)&"</skos:prefLabel>" & vbCrLf
    'objFile.Write"    <skos:prefLabel xml:lang=""en"">"&getTaggedValue(el,"definition")&"</skos:prefLabel>" & vbCrLf
    objFile.Write"    <skos:definition xml:lang=""no"">"&utf8(getCleanDefinitionText(el))&"</skos:definition>" & vbCrLf
    'objFile.Write"    <skos:definition xml:lang=""en"">"&getTaggedValue(el,"definition")&"</skos:definition>" & vbCrLf
    objFile.Write"  </skos:ConceptScheme>" & vbCrLf


	dim attr as EA.Attribute
	for each attr in el.Attributes
		'Repository.WriteOutput "Script", "Debug: attr.Name ["&attr.Name&"]",0
		call listSKOSfraKode(attr,el.Name,namespace)
	next
	'Repository.WriteOutput "Script", "</rdf:RDF>",0
	objFile.Write"</rdf:RDF>" & vbCrLf
	objFile.Close


	' Release the file system object
    Set objFSO= Nothing
	Repository.WriteOutput "Script", "SKOS/RDF/xml-file: "&outFile&" written",0
	
end sub

Sub listSKOSfraKode(attr, codelist, namespace)

	dim presentasjonsnavn, uricode
	if attr.Default <> "" then
		uricode = underscore(attr.Default)
		if attr.Default <> getNCNameX(attr.Default) then
			Repository.WriteOutput "Script", "Trying to make legal http-URI out of initial value for this code: [" & attr.Name & " = " & attr.Default & "] -> [" & uricode & "]",0
		end if
	else
		uricode = underscore(attr.Name)
		if attr.Name <> getNCNameX(attr.Name) then
			Repository.WriteOutput "Script", "Trying to make legal http-URI out of this code: [" & attr.Name & "] -> ["& uricode &"]",0
		end if
	end if

	'objFile.Write"  <skos:Concept rdf:about="""&utf8(codelist)&"/"&utf8(attr.Name)&""">" & vbCrLf
	objFile.Write"  <skos:Concept rdf:about="""&utf8(codelist)&"/"&utf8(uricode)&""">" & vbCrLf
	objFile.Write"    <skos:inScheme rdf:resource="""&utf8(codelist)&"""/>" & vbCrLf
	presentasjonsnavn = getTaggedValue(attr,"SOSI_presentasjonsnavn") 
	if presentasjonsnavn = "" then presentasjonsnavn = attr.Name
	objFile.Write"    <skos:prefLabel xml:lang=""no"">"&utf8(presentasjonsnavn)&"</skos:prefLabel>" & vbCrLf
        '<skos:prefLabel xml:lang=""en""">"&getTaggedValue(el,"SOSI_presentasjonsnavn")&"</skos:prefLabel>
    objFile.Write"    <skos:definition xml:lang=""no"">"&utf8(getCleanDefinitionText(attr))&"</skos:definition>" & vbCrLf
        '<skos:definition xml:lang="en">Measured in terrain</skos:definition>
    objFile.Write"  </skos:Concept>" & vbCrLf

		'<skos:broader rdf:resource="Målemetode/terrengmåltUspesifisertMåleinstrument"/>
		
		
	' write each code to a to separate filer in a subfolder
	Set codeFSO=CreateObject("Scripting.FileSystemObject")
	outCodeFile = codeList & "\" & uricode & ".rdf"
	'Repository.WriteOutput "Script", "Debug: outCodeFile ["&outCodeFile&"]",0
	Set objCodeFile = codeFSO.CreateTextFile(outCodeFile,True,False)
	'  får ut 16-bits unicode ved å sette True som siste flagg i kallet over.
	objCodeFile.Write"<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
	objCodeFile.Write"<rdf:RDF" & vbCrLf
    objCodeFile.Write"  xmlns:skos=""http://www.w3.org/2004/02/skos/core#""" & vbCrLf
	objCodeFile.Write"  xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""" & vbCrLf
	objCodeFile.Write"  xml:base="""&utf8(namespace)&"/"&utf8(codelist)&"/"&""">" & vbCrLf

 	objCodeFile.Write"  <skos:Concept rdf:about="""&utf8(uricode)&""">" & vbCrLf
 	objCodeFile.Write"    <skos:inScheme rdf:resource="""&utf8(namespace)&"/"&utf8(codelist)&"""/>" & vbCrLf

	objCodeFile.Write"    <skos:prefLabel xml:lang=""no"">"&utf8(presentasjonsnavn)&"</skos:prefLabel>" & vbCrLf
    objCodeFile.Write"    <skos:definition xml:lang=""no"">"&utf8(getCleanDefinitionText(attr))&"</skos:definition>" & vbCrLf
	if codelist = "Kommunenummer" then
		dim fy
		fy = Mid(attr.Name,1,2)
		objCodeFile.Write"    <skos:broader rdf:resource="""&utf8(namespace)&"/Fylkesnummer/"&fy&"""/>" & vbCrLf
	end if
	objCodeFile.Write"  </skos:Concept>" & vbCrLf
	objCodeFile.Write"</rdf:RDF>" & vbCrLf

	objCodeFile.Close

    Set codeFSO= Nothing
		
	
End Sub

function getTaggedValue(element,taggedValueName)
		dim i, existingTaggedValue
		getTaggedValue = ""
		for i = 0 to element.TaggedValues.Count - 1
			set existingTaggedValue = element.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = taggedValueName then
				getTaggedValue = existingTaggedValue.Value
			end if
		next
end function

function getPackageTaggedValue(package,taggedValueName)
		dim i, existingTaggedValue
		getPackageTaggedValue = ""
		for i = 0 to package.element.TaggedValues.Count - 1
			set existingTaggedValue = package.element.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = taggedValueName then
				getPackageTaggedValue = existingTaggedValue.Value
			end if
		next
end function

function getCleanDefinitionText(currentElement)
	'removes all formatting in notes fields
    Dim txt, res, tegn, i, u
    u=0
	getCleanDefinitionText = ""
		txt = Trim(currentElement.Notes)
		res = ""
		' loop gjennom alle tegn
		For i = 1 To Len(txt)
		  tegn = Mid(txt,i,1)
		  If tegn = "<" Then
				u = 1
			   'res = res + " "
		  Else 
			If tegn = ">" Then
				u = 0
			   'res = res + " "
				'If tegn = """" Then
				'  res = res + "'"
			Else
				  If tegn < " " Then
					res = res + " "
				  Else
					if u = 0 then
						res = res + Mid(txt,i,1)
					end if
				  End If
				'End If
			End If
		  End If
		  
		Next
		
	getCleanDefinitionText = res

end function

function getAppSchPackage(element)
	dim package as EA.Package
	dim package2 as EA.Package
		
	set package = Repository.GetPackageByID(element.PackageID)
	if LCASE(package.element.Stereotype) = "applicationschema" or package.ParentID = 0 then
		set getAppSchPackage = package
	else
		set package2 = getAppSchParentPackage(package)
		set getAppSchPackage = package2
	end if
	
		
end function

function getAppSchParentPackage(pkg)
	dim package as EA.Package
	dim package2 as EA.Package
	set package = Repository.GetPackageByID(pkg.ParentID)
	if package.ParentID <> 0 then
		if LCASE(package.element.Stereotype) = "applicationschema" then
			set getAppSchParentPackage = package
		else
			set package2 = getAppSchParentPackage(package)
			set getAppSchParentPackage = package2
		end if
		
	end if
	
end function


function getNCNameX(str)
	' make name legal NCName
	Dim txt, res, tegn, i, u
    u=0
		txt = Trim(str)
		'res = LCase( Mid(txt,1,1) )
		res = Mid(txt,1,1)
			'Repository.WriteOutput "Script", "New NCName: " & txt & " " & res,0

		' loop gjennom alle tegn
		For i = 2 To Len(txt)
		  ' blank, komma, !, ", #, $, %, &, ', (, ), *, +, /, :, ;, <, =, >, ?, @, [, \, ], ^, `, {, |, }, ~
		  ' (tatt med flere fnuttetyper, men hva med "."?)
		  tegn = Mid(txt,i,1)
		  if tegn = " " or tegn = "," or tegn = """" or tegn = "#" or tegn = "$" or tegn = "%" or tegn = "&" or tegn = "(" or tegn = ")" or tegn = "*" Then
			  'Repository.WriteOutput "Script", "Bad1: " & tegn,0
			  u=1
		  Else
		    if tegn = "+" or tegn = "/" or tegn = ":" or tegn = ";" or tegn = "<" or tegn = ">" or tegn = "?" or tegn = "@" or tegn = "[" or tegn = "\" Then
			    'Repository.WriteOutput "Script", "Bad2: " & tegn,0
			    u=1
		    Else
		      If tegn = "]" or tegn = "^" or tegn = "`" or tegn = "{" or tegn = "|" or tegn = "}" or tegn = "~" or tegn = "'" or tegn = "´" or tegn = "¨" Then
			      'Repository.WriteOutput "Script", "Bad3: " & tegn,0
			      u=1
		      else
			      'Repository.WriteOutput "Script", "Good: " & tegn,0
			      If u = 1 Then
		          res = res + UCase(tegn)
		          u=0
			      else
		          res = res + tegn
		        End If
		      End If
		    End If
		  End If
		Next
		' return res
		getNCNameX = res

End function


function utf8(str)
	' make string utf-8
	Dim txt, res, tegn, utegn, vtegn, wtegn, i
	
    res = ""
	txt = Trim(str)
	' loop gjennom alle tegn
	For i = 1 To Len(txt)
		tegn = Mid(txt,i,1)

		'if      (c <    0x80) {  *out++=  c;                bits= -6; }
        'else if (c <   0x800) {  *out++= ((c >>  6) & 0x1F) | 0xC0;  bits=  0; }
        'else if (c < 0x10000) {  *out++= ((c >> 12) & 0x0F) | 0xE0;  bits=  6; }
        'else                  {  *out++= ((c >> 18) & 0x07) | 0xF0;  bits= 12; }

		if AscW(tegn) < 128 then
			res = res + tegn
		else if AscW(tegn) < 2048 then
			'u = AscW(tegn)
			'Repository.WriteOutput "Script", "tegn: " & AscW(tegn) & " " & Chr(AscW(tegn) / 64) & " " & int(u / 64),0
			'            c   229=E5/1110 0101
			'            c   192=C0/1100 0000  64=40/0100 0000
			utegn = Chr((int(AscW(tegn) / 64) or 192) )
			res = res + utegn
			'               c          63=3F/0011 1111
			vtegn = Chr((AscW(tegn) and 63) or 128)
			res = res + vtegn
			'            C3A5=å   195/1100 0011   165/1010 0101
			'Repository.WriteOutput "Script", "utf8: " & tegn & " -> " & utegn & " + " & vtegn,0
			'Repository.WriteOutput "Script", "int : " & AscW(tegn) & " -> " & Asc(utegn) & " + " & Asc(vtegn),0
		else if AscW(tegn) < 65536 then
			utegn = Chr((int(AscW(tegn) / 4096) or 224) )
			res = res + utegn
			vtegn = Chr((int(AscW(tegn) / 64) or 128) )
			res = res + vtegn
			wtegn = Chr((AscW(tegn) and 63) or 128)
			res = res + wtegn
			'putchar (0xE0 | c>>12);
			'putchar (0x80 | c>>6 & 0x3F);
			'putchar (0x80 | c & 0x3F);
		else if AscW(tegn) < 2097152 then	'/* 2^21 */
			'putchar (0xF0 | c>>18);
			'putchar (0x80 | c>>12 & 0x3F);
			'putchar (0x80 | c>>6 & 0x3F);
			'putchar (0x80 | c & 0x3F);
		end if
		end if
		end if
		end if

	Next
	' return res
	utf8 = res

End function


function underscore(name)
		' make name legal NCName by replacing each bad character with a "_", typically used for codelist with proper names.)

    Dim txt, res, tegn, i, u
    u=0
	txt = Trim(name)
	' loop gjennom alle tegn
	For i = 1 To Len(txt)
		' blank, komma, !, ", #, $, %, &, ', (, ), *, +, /, :, ;, <, =, >, ?, @, [, \, ], ^, `, {, |, }, ~
		' (tatt med flere fnuttetyper, men hva med "."?)
		tegn = Mid(txt,i,1)
		if tegn = " " or tegn = "," or tegn = """" or tegn = "#" or tegn = "$" or tegn = "%" or tegn = "&" or tegn = "(" or tegn = ")" or tegn = "*" Then
			'Repository.WriteOutput "Script", "Bad1: " & tegn,0
			u=1
		Else
			if tegn = "+" or tegn = "/" or tegn = ":" or tegn = ";" or tegn = "<" or tegn = ">" or tegn = "?" or tegn = "@" or tegn = "[" or tegn = "\" Then
				'Repository.WriteOutput "Script", "Bad2: " & tegn,0
				u=1
			Else
				If tegn = "]" or tegn = "^" or tegn = "`" or tegn = "{" or tegn = "|" or tegn = "}" or tegn = "~" or tegn = "'" or tegn = "´" or tegn = "¨" Then
					'Repository.WriteOutput "Script", "Bad3: " & tegn,0
					u=1
				else
					'Repository.WriteOutput "Script", "Good: " & tegn,0
					If u = 1 Then
						res = res + "_"
						'res = res + UCase(tegn)
						u=0
						'else
					End If
					res = res + tegn
				End If
		    End If
		End If
	Next
	tegn = Mid(res,1,1)
	'if Ucase(res) = "CON" or Ucase(res) = "PRN" or Ucase(res) = "AUX" or Ucase(res) = "NUL" or tegn = "-" or tegn = "0" or tegn = "1" or tegn = "2" or tegn = "3" or tegn = "4" or tegn = "5" or tegn = "6" or tegn = "7" or tegn = "8" or tegn = "9" Then
	if Ucase(res) = "CON" or Ucase(res) = "PRN" or Ucase(res) = "AUX" or Ucase(res) = "NUL" Then
		res = "_" + res
'	if Ucase(res) = "COM1" then  res = "_" + res
'COM1, COM2, COM3, COM4, COM5, COM6, COM7, COM8, COM9.
'LPT1, LPT2, LPT3, LPT4, LPT5, LPT6, LPT7, LPT8, LPT9
		Repository.WriteOutput "Script", "Trying to make legal filename out of this code by prefixing an underscore : [" & res &"]",0
	end if
	underscore = res

End function


'Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
'    ByVal CodePage As Long, _
'    ByVal dwFlags As Long, _
'    ByVal lpWideCharStr As Long, _
'    ByVal cchWideChar As Long, _
'    ByVal lpMultiByteStr As Long, _
'    ByVal cbMultiByte As Long, _
'    ByVal lpDefaultChar As Long, _
'    ByVal lpUsedDefaultChar As Long) As Long
'    
'' CodePage constant for UTF-8
'Private Const CP_UTF8 = 65001''''''
'
'''' Return byte array with VBA "Unicode" string encoded in UTF-8
'Public Function Utf8BytesFromString(strInput As String) As Byte()
'    Dim nBytes As Long
'    Dim abBuffer() As Byte
'    ' Get length in bytes *including* terminating null
'    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
'    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
'    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
'    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
'    Utf8BytesFromString = abBuffer
'End Function

'CStringA ConvertUnicodeToUTF8(const CStringW& uni)
'{
'    if (uni.IsEmpty()) return ""; // nothing to do
'    CStringA utf8;
'    int cc=0;
'    // get length (cc) of the new multibyte string excluding the \0 terminator first
'    if ((cc = WideCharToMultiByte(CP_UTF8, 0, uni, -1, NULL, 0, 0, 0) - 1) > 0)
'    { 
'        // convert
'        char *buf = utf8.GetBuffer(cc);
'        if (buf) WideCharToMultiByte(CP_UTF8, 0, uni, -1, buf, cc, 0, 0);
'        utf8.ReleaseBuffer();
'    }
'    return utf8;
'}

listKoderForEnValgtKodeliste
