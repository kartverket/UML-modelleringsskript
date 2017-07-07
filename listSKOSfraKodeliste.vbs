option explicit

!INC Local Scripts.EAConstants-VBScript

' script:			listSKOSfraKodeliste
' description:		Skriver til egne SKOS-filer under samme sti som .eap-fila ligger.
' date  :			2017-06-29,07-07
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
			"Please select a  CodeList class in the Project Browser and try once more." )
		end if
		'Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
		Repository.EnsureOutputVisible "Script"
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires a CodeList class to be selected in the Project Browser." & vbCrLf & _
	  "Please select a  CodeList class in the Project Browser and try again." )
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
	outFile = getNCNameX(el.Name)&".skos.xml"
	Repository.WriteOutput "Script", Now & " outFile: " & outFile, 0
	Set objFile = objFSO.CreateTextFile(outFile,True,True)
	if not objFSO.FolderExists(el.Name) then
		objFSO.CreateFolder el.Name
	end if

	Repository.WriteOutput "Script", "Writes Codelist Name: " & el.Name & " to file " & outfile& " and subfolder " & el.Name,0
	Repository.WriteOutput "Script", "With namespace: " & namespace,0

	objFile.Write"<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
	objFile.Write"<rdf:RDF" & vbCrLf
    objFile.Write"  xmlns:skos=""http://www.w3.org/2004/02/skos/core#""" & vbCrLf
	objFile.Write"  xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""" & vbCrLf
	objFile.Write"  xml:base="""&namespace&"/"">" & vbCrLf


    objFile.Write"  <skos:ConceptScheme rdf:about="""&el.Name&""">" & vbCrLf
	presentasjonsnavn = getTaggedValue(el,"SOSI_presentasjonsnavn") 
	if presentasjonsnavn = "" then presentasjonsnavn = el.Name
    objFile.Write"    <skos:prefLabel xml:lang=""no"">"&presentasjonsnavn&"</skos:prefLabel>" & vbCrLf
    'objFile.Write"    <skos:prefLabel xml:lang=""en"">"&getTaggedValue(el,"definition")&"</skos:prefLabel>" & vbCrLf
    objFile.Write"    <skos:definition xml:lang=""no"">"&getCleanDefinitionText(el)&"</skos:definition>" & vbCrLf
    'objFile.Write"    <skos:definition xml:lang=""en"">"&getTaggedValue(el,"definition")&"</skos:definition>" & vbCrLf
    objFile.Write"  </skos:ConceptScheme>" & vbCrLf


	dim attr as EA.Attribute
	for each attr in el.Attributes
		'Repository.WriteOutput "Script", "Debug:: [" & attr.Name & "] - " & getNCName(attr.Name) & " ",0
		if attr.Name = getNCNameX(attr.Name) then
			call listSKOSfraKode(attr,el.Name,namespace)
		else
			Repository.WriteOutput "Script", "Error trying to make http-URI out of this code: [" & attr.Name & "] - not a valid NCName!",0
		end if
	next
	'Repository.WriteOutput "Script", "</rdf:RDF>",0
	objFile.Write"</rdf:RDF>" & vbCrLf
	objFile.Close


	' Release the file system object
    Set objFSO= Nothing
	Repository.WriteOutput "Script", "SKOS/RDF/xml-file: "&outFile&" written",0
	
end sub

Sub listSKOSfraKode(attr, codelist, namespace)

	dim presentasjonsnavn
 	objFile.Write"  <skos:Concept rdf:about="""&codelist&"/"&attr.Name&""">" & vbCrLf
 	objFile.Write"    <skos:inScheme rdf:resource="""&codelist&"""/>" & vbCrLf
	presentasjonsnavn = getTaggedValue(attr,"SOSI_presentasjonsnavn") 
	if presentasjonsnavn = "" then presentasjonsnavn = attr.Name
	objFile.Write"    <skos:prefLabel xml:lang=""no"">"&presentasjonsnavn&"</skos:prefLabel>" & vbCrLf
        '<skos:prefLabel xml:lang=""en""">"&getTaggedValue(el,"SOSI_presentasjonsnavn")&"</skos:prefLabel>
    objFile.Write"    <skos:definition xml:lang=""no"">"&getCleanDefinitionText(attr)&"</skos:definition>" & vbCrLf
        '<skos:definition xml:lang="en">Measured in terrain</skos:definition>
    objFile.Write"  </skos:Concept>" & vbCrLf

		'<skos:broader rdf:resource="Målemetode/terrengmåltUspesifisertMåleinstrument"/>
		
		
	' write each code to a to separate filer in a subfolder
	Set codeFSO=CreateObject("Scripting.FileSystemObject")
	outCodeFile = codeList&"\"&getNCNameX(attr.Name)&".skos.xml"
	Set objCodeFile = codeFSO.CreateTextFile(outCodeFile,True,True)

	objCodeFile.Write"<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
	objCodeFile.Write"<rdf:RDF" & vbCrLf
    objCodeFile.Write"  xmlns:skos=""http://www.w3.org/2004/02/skos/core#""" & vbCrLf
	objCodeFile.Write"  xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""" & vbCrLf
	objCodeFile.Write"  xml:base="""&namespace&"/"&codelist&"/"&""">" & vbCrLf

 	objCodeFile.Write"  <skos:Concept rdf:about="""&attr.Name&""">" & vbCrLf
 	objCodeFile.Write"    <skos:inScheme rdf:resource="""&namespace&"/"&codelist&"""/>" & vbCrLf

	objCodeFile.Write"    <skos:prefLabel xml:lang=""no"">"&presentasjonsnavn&"</skos:prefLabel>" & vbCrLf
    objCodeFile.Write"    <skos:definition xml:lang=""no"">"&getCleanDefinitionText(attr)&"</skos:definition>" & vbCrLf
	if codelist = "Kommunenummer" then
		dim fy
		fy = Mid(attr.Name,1,2)
		objCodeFile.Write"    <skos:broader rdf:resource="""&namespace&"/Fylkesnummer/"&fy&"""/>" & vbCrLf
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

listKoderForEnValgtKodeliste
