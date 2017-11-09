option explicit

!INC Local Scripts.EAConstants-VBScript

' skriptnavn:       listGMLDICTfraKodeliste
' description:		Skriver en kodeliste til egen gml:Dictionary.xml fil. på samme sti som .eap-fila ligger.
' author:			Kent
' date  :			2017-06-29, 07-07,09-08,11-09
	DIM objFSO
	DIM outFile
	DIM objFile

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
			box = Msgbox ("List class : [«" & theElement.Stereotype &"» "& theElement.Name & "]. to gml:Dictionary.xml format."& vbCrLf & "Creates one file with all codes in the same folder as the .eap-file.",1)
			select case box
			case vbOK
		 		'Session.Output("Debug: ------------ Start class: [«" &theElement.Stereotype& "» " &theElement.Name& "] of type. [" &theElement.Type& "]. ")
				'inputBoxGUI to receive user input regarding the namespace
				'if namespace = "" and getTaggedValue(theElement, "asDictionary") = "true" then
				dim namespace, nsp
				'namespace = ""
				namespace = getTaggedValue(theElement, "codeList")
				if namespace <> "" then
					nsp = Mid(namespace,Len(namespace)-Len(theElement.Name)+1,Len(theElement.Name))
					'Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
					if nsp = theElement.Name and nsp <> namespace then
						'Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
						namespace = Mid(namespace,1,Len(namespace)-Len(nsp)-1)
						'Repository.WriteOutput "Script"," namespace shortened:"&namespace &" to "&nsp, 0
					end if
				end if
				if namespace = "" then
					namespace = getPackageTaggedValue(getAppSchPackage(theElement),"targetNamespace")
				end if

				namespace = InputBox("Please select the codespace name for the codelist.", "namespace", namespace)
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
	dim presentasjonsnavn

	Set objFSO=CreateObject("Scripting.FileSystemObject")
	outFile = el.Name&".xml"
	Set objFile = objFSO.CreateTextFile(outFile,True,False)
	'  får ut 16-bits unicode ved å sette True som siste flagg i kallet over.
	Repository.WriteOutput "Script", "Writes Codelist Name: " & el.Name & " to file " & outfile,0
	Repository.WriteOutput "Script", "With codespace: " & namespace,0

	objFile.Write"<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
	objFile.Write"<?xml-stylesheet type='text/xsl' href='./CodelistDictionary-v32.xsl'?>" & vbCrLf
	objFile.Write"<Dictionary xmlns=""http://www.opengis.net/gml/3.2""" & vbCrLf
    objFile.Write"  xmlns:gml=""http://www.opengis.net/gml/3.2""" & vbCrLf
    objFile.Write"  xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & vbCrLf
    objFile.Write"  gml:id="""&utf8(el.Name)&"""" & vbCrLf
    objFile.Write"  xsi:schemaLocation=""http://www.opengis.net/gml/3.2 http://schemas.opengis.net/gml/3.2.1/gml.xsd"">" & vbCrLf
	objFile.Write"  <description>"&utf8(getCleanDefinitionText(el))&"</description>" & vbCrLf
	objFile.Write"  <identifier codeSpace="""&utf8(namespace)&""">"&utf8(el.Name)&"</identifier>" & vbCrLf




	dim attr as EA.Attribute
	for each attr in el.Attributes
		'Repository.WriteOutput "Script", Now & " " & el.Name & "." & attr.Name, 0

		call listDICTfraKode(attr,el.Name,namespace)
	next
	objFile.Write"</Dictionary>" & vbCrLf
	objFile.Close


	' Release the file system object
    Set objFSO= Nothing
	Repository.WriteOutput "Script", "gml:Dictionary.xml-file: "&outFile&" written",0

end sub

Sub listDICTfraKode(attr, codelist, namespace)

	dim presentasjonsnavn
	presentasjonsnavn = getTaggedValue(attr,"SOSI_presentasjonsnavn") 
	
	
	objFile.Write"  <dictionaryEntry>" & vbCrLf
    if attr.Default <> "" then
		objFile.Write"    <Definition gml:id="""&utf8(codelist)&"."&utf8(getNCNameX(attr.Default))&""">" & vbCrLf
		objFile.Write"      <description>"&utf8(getCleanDefinitionText(attr))&"</description>" & vbCrLf
		objFile.Write"      <identifier codeSpace="""&utf8(namespace)&"/"&utf8(codelist)&""">"&utf8(attr.Default)&"</identifier>" & vbCrLf
		if presentasjonsnavn <> "" then
			objFile.Write"      <name>"&utf8(presentasjonsnavn)&"</name>" & vbCrLf
		end if
  		objFile.Write"      <name>"&utf8(attr.Name)&"</name>" & vbCrLf
	else
		objFile.Write"    <Definition gml:id="""&utf8(codelist)&"."&utf8(getNCNameX(attr.Name))&""">" & vbCrLf
		objFile.Write"      <description>"&utf8(getCleanDefinitionText(attr))&"</description>" & vbCrLf
		objFile.Write"      <identifier codeSpace="""&utf8(namespace)&"/"&utf8(codelist)&""">"&utf8(attr.Name)&"</identifier>" & vbCrLf
		if presentasjonsnavn <> "" then
			objFile.Write"      <name>"&utf8(presentasjonsnavn)&"</name>" & vbCrLf
		end if
 	end if
 

	objFile.Write"    </Definition>" & vbCrLf
    objFile.Write"  </dictionaryEntry>" & vbCrLf

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
		'if Mid(txt,1,1) < ":" then
		'	res = "_" + Mid(txt,1,1)
		'else
			res = Mid(txt,1,1)
		'end if
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

		if AscW(tegn) < 128 then
			res = res + tegn
		else if AscW(tegn) < 2048 then
			utegn = Chr((int(AscW(tegn) / 64) or 192) )
			res = res + utegn
			'               c          63=3F/0011 1111
			vtegn = Chr((AscW(tegn) and 63) or 128)
			res = res + vtegn
		else if AscW(tegn) < 65536 then
			utegn = Chr((int(AscW(tegn) / 4096) or 224) )
			res = res + utegn
			vtegn = Chr((int(AscW(tegn) / 64) or 128) )
			res = res + vtegn
			wtegn = Chr((AscW(tegn) and 63) or 128)
			res = res + wtegn
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


listKoderForEnValgtKodeliste
