option explicit

!INC Local Scripts.EAConstants-VBScript

' script:			listSKOSfraKodeliste
' description:		Skriver en kodeliste til egne SKOS-filer under samme sti som .eap-fila ligger.
' author:			Kent
' date:				2017-06-29,07-07,09-08,09-14,11-09,12-05, 2918-02-20
' date:				2018-10-05 html5, 2018-12-18 feilretting (samisk tegn eng)
' date:				2019-05-23 index.html, 05-27 +/-style?
' date:				2020-05-11 code html points to skos-file for the code
	DIM objFSO
	DIM outFile
	DIM objFile
	DIM htmFile
	DIM idxFile

	DIM codeFSO
	DIM outCodeFile
	DIM objCodeFile
	DIM htmlFSO
	DIM outHtmlFile
	DIM outIdxFile

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
		 		Session.Output("Debug: ------------ Start namespace: [" &namespace& "]. "&Len(namespace)&"  "&Len(theElement.Name)+1&"")
				if namespace <> "" and namespace <> theElement.Name and Len(namespace) > Len(theElement.Name)+1 then
					nsp = Mid(namespace,Len(namespace)-Len(theElement.Name)+1,Len(theElement.Name))
					Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
					if nsp = theElement.Name and nsp <> namespace then
						Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
						namespace = Mid(namespace,1,Len(namespace)-Len(nsp)-1)
						Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
					end if
				end if
				if namespace = "" then
					namespace = getPackageTaggedValue(getAppSchPackage(theElement),"targetNamespace")
				end if

				namespace = InputBox("Please select the namespace name for the codelist.", "namespace", namespace)
				if Mid(namespace,Len(namespace),1) = "/" then
					namespace = Mid(namespace,1,Len(namespace)-1)
				end if
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
	Set idxFile = objFSO.CreateTextFile(el.Name & "\index.html",True,False)
	
	Repository.WriteOutput "Script", "Writes Codelist Name: " & el.Name & " to file " & outfile& " and subfolder " & el.Name,0
	Repository.WriteOutput "Script", "With namespace: " & namespace,0

	objFile.Write"<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
	objFile.Write"<?xml-stylesheet type='text/xsl' href='./CodelistDictionary-skosrdfxml.xsl'?>" & vbCrLf
	objFile.Write"<rdf:RDF" & vbCrLf
    objFile.Write"  xmlns:skos=""http://www.w3.org/2004/02/skos/core#""" & vbCrLf
	objFile.Write"  xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""" & vbCrLf
	objFile.Write"  xml:base="""&utf8(namespace)&"/"">" & vbCrLf

	idxFile.Write"<!DOCTYPE html>" & vbCrLf
	idxFile.Write"<html lang=""no"">" & vbCrLf
	idxFile.Write"	<head>" & vbCrLf
	idxFile.Write"	  <meta charset=""utf-8""/>" & vbCrLf
	idxFile.Write"	  <title>" & utf8(el.Name) & "</title>" & vbCrLf
	' Style?
'	idxFile.Write"	  <style>" & vbCrLf
'	idxFile.Write"	  	table, th, td {" & vbCrLf
'	idxFile.Write"	  	border: 1px solid black;" & vbCrLf
'	idxFile.Write"	  	border-collapse: collapse;}" & vbCrLf
'	idxFile.Write"	  </style>" & vbCrLf
	'idxFile.Write"	  " & vbCrLf

	idxFile.Write"	</head>" & vbCrLf
	idxFile.Write"	<body>" & vbCrLf
	idxFile.Write"    <p>xml:base=" & utf8(namespace) & "</p>" & vbCrLf
	idxFile.Write"    <p>http-URI=" & utf8(namespace) & "/" & utf8(el.Name) & "</p>" & vbCrLf
	'idxFile.Write"    <p>kodens navn=" & utf8(uricode) & "</p>" & vbCrLf
	if getTaggedValue(el,"SOSI_presentasjonsnavn") <> "" then
		idxFile.Write"    <p>presentasjonsnavn=" & utf8(getTaggedValue(el,"SOSI_presentasjonsnavn")) & "</p>" & vbCrLf
	end if
	idxFile.Write"    <p>kodelistas definisjon=" & utf8(getCleanDefinitionText(el)) & "</p>" & vbCrLf

	'idxFile.Write"   <table border=""1"">" & vbCrLf
	idxFile.Write"   <table>" & vbCrLf
	idxFile.Write"   <tbody align=""left"">" & vbCrLf
	idxFile.Write"   <tr>" & vbCrLf

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
		if el.Name = "Kommunenummer" or el.Name = "Fylkesnummer" then
			Repository.WriteOutput "Script", Now & "  " & attr.Name & "." & attr.Notes, 0
			if InStr(LCASE(attr.Notes),"utgått") then 
				Repository.WriteOutput "Script", Now & " utgått: " & attr.Name & "." & attr.Notes, 0
				call listSKOSfraKode(attr,el.Name,namespace)
			else
			if Int(attr.Name) > 2099 and Int(attr.Name) < 2400 then 
				Repository.WriteOutput "Script", Now & " svalb.: " & attr.Name & "." & attr.Notes, 0
				call listSKOSfraKode(attr,el.Name,namespace)
			else
			if Int(attr.Name) > 20 and Int(attr.Name) < 24 then 
				Repository.WriteOutput "Script", Now & " Svalb.: " & attr.Name & "." & attr.Notes, 0
				call listSKOSfraKode(attr,el.Name,namespace)
			else
				call listSKOSfraKode(attr,el.Name,namespace)
			end if
			end if
			end if
		else
			call listSKOSfraKode(attr,el.Name,namespace)
		end if
	next
	'Repository.WriteOutput "Script", "</rdf:RDF>",0
	objFile.Write"</rdf:RDF>" & vbCrLf
	objFile.Close
	
	
	idxFile.Write"  </tr>" & vbCrLf
	idxFile.Write"  </tbody>" & vbCrLf
	idxFile.Write"  </table>" & vbCrLf
	idxFile.Write"  </body>" & vbCrLf
	idxFile.Write"</html>" & vbCrLf
	idxFile.Close

	' Release the file system object
    Set objFSO= Nothing
	Repository.WriteOutput "Script", "html5/SKOS/RDF/xml-file: "&outFile&" written",0
	
end sub

Sub listSKOSfraKode(attr, codelist, namespace)

	dim presentasjonsnavn, uricode, fy, tegn
	if attr.Default <> "" then
		'filter mot å legge meningsløse tallkoder (i initialverdier) inn i filer laget for semantiske søk
		tegn = Mid(attr.Default,1,1)
		if tegn = "0" or tegn = "1" or tegn = "2" or tegn = "3" or tegn = "4" or tegn = "5" or tegn = "6" or tegn = "7" or tegn = "8" or tegn = "9" Then
			uricode = underscore(attr.Name)
		else
			uricode = underscore(attr.Default)
			if attr.Default <> getNCNameX(attr.Default) then
				Repository.WriteOutput "Script", "Trying to make legal http-URI out of initial value for this code: [" & attr.Name & " = " & attr.Default & "] -> [" & uricode & "]",0
			end if
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
	if codelist = "Kommunenummer" then
		fy = Mid(uricode,1,2)
		objFile.Write"    <skos:broader rdf:resource="""&utf8(namespace)&"/Fylkesnummer/"&fy&"""/>" & vbCrLf
	end if
	if getTaggedValue(attr,"designation") <> "" then
		objFile.Write"    <skos:hiddenLabel xml:lang=""en"">" & Mid(utf8(getTaggedValue(attr,"designation")),2,Len(getTaggedValue(attr,"designation"))-4) & "</skos:hiddenLabel>" & vbCrLf
	end if
	if getTaggedValue(attr,"definition") <> "" then
		objFile.Write"    <skos:definition xml:lang=""en"">" & Mid(utf8(getTaggedValue(attr,"definition")),2,Len(getTaggedValue(attr,"definition"))-4) & "</skos:definition>>" & vbCrLf
	end if
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
		fy = Mid(uricode,1,2)
		objCodeFile.Write"    <skos:broader rdf:resource="""&utf8(namespace)&"/Fylkesnummer/"&fy&"""/>" & vbCrLf
	end if
	if getTaggedValue(attr,"designation") <> "" then
		objCodeFile.Write"    <skos:hiddenLabel xml:lang=""en"">" & Mid(utf8(getTaggedValue(attr,"designation")),2,Len(getTaggedValue(attr,"designation"))-4) & "</skos:hiddenLabel>" & vbCrLf
	end if
	if getTaggedValue(attr,"definition") <> "" then
		objCodeFile.Write"    <skos:definition xml:lang=""en"">" & Mid(utf8(getTaggedValue(attr,"definition")),2,Len(getTaggedValue(attr,"definition"))-4) & "</skos:definition>>" & vbCrLf
	end if
	objCodeFile.Write"  </skos:Concept>" & vbCrLf
	objCodeFile.Write"</rdf:RDF>" & vbCrLf

	objCodeFile.Close

    Set codeFSO= Nothing

	Set htmlFSO=CreateObject("Scripting.FileSystemObject")
	outHtmlFile = codeList & "\" & uricode
	'Repository.WriteOutput "Script", Now & " outHtmlFile: " & outHtmlFile, 0
	Set htmFile = objFSO.CreateTextFile(outHtmlFile,True,False)
	htmFile.Write"<!DOCTYPE html>" & vbCrLf
	htmFile.Write"<html lang=""no"">" & vbCrLf
	htmFile.Write"	<head>" & vbCrLf
	htmFile.Write"	  <meta charset=""utf-8""/>" & vbCrLf
	htmFile.Write"	  <title>" & utf8(codelist) & " " & utf8(uricode) & "</title>" & vbCrLf
	htmFile.Write"	</head>" & vbCrLf
	htmFile.Write"	<body>" & vbCrLf
	htmFile.Write"    <p>xml:base = " & utf8(namespace) & "/" & utf8(codelist) & "</p>" & vbCrLf
	htmFile.Write"    <p>http-URI = " & utf8(namespace) & "/" & utf8(codelist) & "/" & utf8(uricode) & "</p>" & vbCrLf
	htmFile.Write"    <p>kodens navn = " & utf8(uricode) & "</p>" & vbCrLf
	htmFile.Write"    <p>presentasjonsnavn = " & utf8(presentasjonsnavn) & "</p>" & vbCrLf
	htmFile.Write"    <p>kodens definisjon = " & utf8(getCleanDefinitionText(attr)) & "</p>" & vbCrLf
	'htmFile.Write"    <p>code description=" & attr.Notes & "</p>" & vbCrLf
	if getTaggedValue(attr,"SOSI_verdi") <> "" then
		htmFile.Write"    <p>SOSI_verdi = " & utf8(getTaggedValue(attr,"SOSI_verdi")) & "</p>" & vbCrLf
	end if
	if getTaggedValue(attr,"designation") <> "" then
		htmFile.Write"    <p>kodens engelske tekniske navn = " & Mid(utf8(getTaggedValue(attr,"designation")),2,Len(getTaggedValue(attr,"designation"))-4) & "</p>" & vbCrLf
	end if
	if getTaggedValue(attr,"definition") <> "" then
		htmFile.Write"    <p>kodens engelske definisjon = " & Mid(utf8(getTaggedValue(attr,"definition")),2,Len(getTaggedValue(attr,"definition"))-4) & "</p>" & vbCrLf
	end if
	htmFile.Write"    <p>lenke til SKOS-fil: <a href=" & utf8(uricode) & ".rdf>" & utf8(uricode) & ".rdf</a></p>" & vbCrLf
	htmFile.Write"  </body>" & vbCrLf
	htmFile.Write"</html>" & vbCrLf

	htmFile.Close

    Set htmlFSO= Nothing

	'add one line in index.htm <a href="land/nasjk.m4a" alt="nasjk.m4a">nÃ¥sjk</a>
	'idxFile.Write"    <p>kode <a href=" & utf8(namespace) & "/" & utf8(codelist) & "/" & utf8(uricode) & "/a>	" & utf8(uricode) & "</p> - <p>" & utf8(getCleanDefinitionText(attr)) & "</p>"& vbCrLf
	'idxFile.Write"    <p>kode <a href=" & utf8(namespace) & "/" & utf8(codelist) & "/" & utf8(uricode) & "/a>	" & utf8(presentasjonsnavn) & "</p> - <p>" & utf8(getCleanDefinitionText(attr)) & "</p>"& vbCrLf
	idxFile.Write"    <td>presentasjonsnavn: <a href=" & utf8(namespace) & "/" & utf8(codelist) & "/" & utf8(uricode) & ">	" & utf8(presentasjonsnavn) & "</a></td><td>" & utf8(getCleanDefinitionText(attr)) & "</td></tr><tr>" & vbCrLf

	
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
		'txt = Trim(currentElement.Notes)
		txt = currentElement.Notes
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
		  ' (tatt med flere fnuttetyper, men hva med "."?) (‘'«»’)
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
	Dim txt, res, tegn, utegn, vtegn, wtegn, xtegn, i
	
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
			'putchar (0xE0 | c>>12);  E0=224, 2^12=4096
			'putchar (0x80 | c>>6 & 0x3F);  80=128, 2^6=64
			'putchar (0x80 | c & 0x3F);  80=128
		else if AscW(tegn) < 2097152 then	'/* 2^21 */
			utegn = Chr((int(AscW(tegn) / 262144) or 240) )
			res = res + utegn
			vtegn = Chr((int(AscW(tegn) / 4096) or 128) )
			res = res + vtegn
			wtegn = Chr((int(AscW(tegn) / 64) or 128) )
			res = res + wtegn
			xtegn = Chr((AscW(tegn) and 63) or 128)
			res = res + xtegn
			'putchar (0xF0 | c>>18);  F0=240, 2^18=262144
			'putchar (0x80 | c>>12 & 0x3F); 80=128, 2^12=4096
			'putchar (0x80 | c>>6 & 0x3F);  80=128, 2^6=64
			'putchar (0x80 | c & 0x3F);  80=128, 3F=63
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
