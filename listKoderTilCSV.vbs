option explicit

!INC Local Scripts.EAConstants-VBScript

 ' Script Name: listKoderTilCSV
 ' Author: Kent Jonsrud Kartverket
 ' Purpose: list code names in system output
 ' Date: 2022-11-02 Smårettinger, hvis Utgått finnes i navnet settes status til Utgått
 ' Date: 2021-10-01 Navn;Kodeverdi;Eier;Status;Oppdatert;Versjons ID;Beskrivelse;Gyldig fra;Gyldig til;ID
 ' Date: 2021-09-29 Endra ledetekster 
 ' Date: 2020-11-21
'
 '
 ' TBD: Kodenavn;Utvekslingsalias;Eier;Status;oppdateringsdato;Versjon;Definisjon;Gyldig fra;Gyldig til;ID;BredereID

sub fixOldCodelists(el)
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
'	Repository.WriteOutput "Script", "-----------------------" & Now & " Script listKoder running on CodeList: " & el.Name, 0
	'Repository.WriteOutput "Script", Now & " " & el.Stereotype & " " & el.Name, 0
	'Repository.WriteOutput "Script", "id, kode, definisjon, initialverdi, SOSI_verdi, SOSI_presentasjonsnavn",0
    
	dim id
	id = 1
	Repository.WriteOutput "Script",  "Navn;Kodeverdi;Eier;Status;Oppdatert;Versjons ID;Beskrivelse;Gyldig fra;Gyldig til;ID" ,0
	dim attr as EA.Attribute
	for each attr in el.Attributes
	'	Repository.WriteOutput "Script", Now & " " & el.Name & "." & attr.Name, 0

	'If the old codes are in a form suitable as initial draft of proper definition:
	'kopierKodensNavnTilTomDefinisjon(attr)
	'remember to refrase these definitions according to the rules for developing definitions

	' kopierKodensNavnTilTagSOSI_presentasjonsnavn(attr)

    'flyttInitialverdiTilTagSOSI_verdi(attr)

    'settKodensNavnTilNCName(attr)
    'eller
    'settKodensNavnTilEgen_Navn(attr)

   Call listKodeliste(el, id, attr)
   'id = id + 1

	'listKodeliste(attr)

	next

end sub



Sub listKodeliste(codelist, id, attr)

	dim kodenavn, definisjon, SOSI_verdi, SOSI_presentasjonsnavn, navnerom, gyldig
	'dim codelist
	if getTaggedValue(codelist,"codeList") <> "" then
		navnerom = getTaggedValue(codelist,"codeList")
	else
		navnerom = "http://skjema.geonorge.no/sosi/basistype"
	end if
	kodenavn = attr.Name
	definisjon = attr.Notes
	definisjon = Replace(definisjon,"""","")
'	Repository.WriteOutput "Script",  "Navn;Kodeverdi;Eier;Status;Oppdatert;Versjons ID;Beskrivelse;Gyldig fra;Gyldig til;ID" ,0
	'check if the element has a tagged value with the required name
		dim existingTaggedValue AS EA.TaggedValue
		'dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		SOSI_presentasjonsnavn = ""
		SOSI_verdi = ""
		for taggedValuesCounter = 0 to attr.TaggedValues.Count - 1
			set existingTaggedValue = attr.TaggedValues.GetAt(taggedValuesCounter)

			if existingTaggedValue.Name = "SOSI_verdi" then
				SOSI_verdi = existingTaggedValue.Value
			end if
			if existingTaggedValue.Name = "SOSI_presentasjonsnavn" then
				SOSI_presentasjonsnavn = existingTaggedValue.Value
			end if
		next
'   		Repository.WriteOutput "Script", "INSERT INTO databaseskjema." & codelist & " VALUES (" & id & ",'"& kodenavn & "','" & attr.Default & "','" & ikkelinjeskiftEllerEnkeltfnutt(definisjon) & "','" & SOSI_verdi & "','" & SOSI_presentasjonsnavn & "');",0
	if InStr(kodenavn,"Utgått") then
		gyldig="Utgått"
	else
		gyldig="Gyldig"
	end if
	Repository.WriteOutput "Script",  kodenavn &";"& attr.Default &";SOSI;"&gyldig&";"& date &";0.1;"& ikkelinjeskiftEllerEnkeltfnutt(attr.Notes) &";;;"& navnerom & "/" & attr.Name  ,0

End Sub



Sub kopierKodensNavnTilTomDefinisjon(attr)
		if attr.Notes = "" then
			dim notestring
		  ' Move ALL (old) names to START of definition by commenting out the if/endif around
		  ' and use this "notestring =" instead
			' notestring = attr.Name & " " & attr.Notes
			notestring = attr.Name
			Repository.WriteOutput "Script", "New notestring: " & notestring,0
			attr.Notes = notestring
			attr.Update()
		end if

End Sub

Sub kopierKodensNavnTilTagSOSI_presentasjonsnavn(attr)
  		'Repository.WriteOutput "Script", "SOSI_presentasjonsnavn: " & attr.Name,0
      Call TVSetElementTaggedValue(attr, "SOSI_presentasjonsnavn", attr.Name)

End Sub


Sub flyttInitialverdiTilTagSOSI_verdi(attr)
		If attr.Default <> "" then
  		Repository.WriteOutput "Script", "Initial value moved: " & attr.Default,0

      Call TVSetElementTaggedValue(attr, "SOSI_verdi", attr.Default)

      attr.Default = ""
      attr.Update()

		End if
End Sub

sub TVSetElementTaggedValue( theElement, taggedValueName, taggedValue)
	'Repository.WriteOutput "Script", "  Checking if tagged value [" & taggedValueName & "] exists",0
	if not theElement is nothing and Len(taggedValueName) > 0 then
		dim newTaggedValue as EA.TaggedValue
		set newTaggedValue = nothing
		dim taggedValueExists
		taggedValueExists = False

		'check if the element has a tagged value with the provided name
		dim existingTaggedValue AS EA.TaggedValue
		dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set existingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			if existingTaggedValue.Name = taggedValueName then
				taggedValueExists = True
				set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			end if
		next

		'if the element does not contain a tagged value with the provided name, create a new one
		if not taggedValueExists = True then
			set newTaggedValue = theElement.TaggedValues.AddNew( taggedValueName, taggedValue )
			newTaggedValue.Update()
			'Repository.WriteOutput "Script", "    ADDED tagged value ["& taggedValueName & " " & taggedValue & "]",0
		Else
		  If currentExistingTaggedValue.Value = "" Then
		    currentExistingTaggedValue.Value = taggedValue
		    currentExistingTaggedValue.Update()
			  ' Repository.WriteOutput "Script", "    ADDED value ["& taggedValueName & " " & taggedValue& "]",0
		  End If
			'Repository.WriteOutput "Script", "    FOUND tagged value ["& taggedValueName & " " & currentExistingTaggedValue.Value & "]",0
		end if
	end if
end Sub

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


Sub settKodensNavnTilNCName(attr)
		' make name legal NCName
		' (alternatively replace each bad character with a "_", typically used for codelist with proper names.)
		' (Sub settBlankeIKodensNavnTil_(attr))
    Dim txt, txt1, txt2, res, tegn, i, u
    u=0
		'Repository.WriteOutput "Script", "Old code: " & attr.Name,0
		txt = Trim(attr.Name)
		res = ""
			'Repository.WriteOutput "Script", "New NCName: " & txt & " " & res,0

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
				if res = "" then
					if tegn = "1" or tegn = "2" or tegn = "3" or tegn = "4" or tegn = "5" or tegn = "6" or tegn = "7" or tegn = "8" or tegn = "9" or tegn = "0" or tegn = "-" or tegn = "." Then
						' NCNames can not start with any of these characters, skip this
					else
						If u = 1 Then
							res = res + UCase(tegn)
							u=0
						else
							res = res + tegn
						end if
					end if
				else
					'Repository.WriteOutput "Script", "Good: " & tegn & "  " & i & " " & u,0
					If u = 1 Then
						res = res + UCase(tegn)
						u=0
					else
						res = res + tegn
					End If
		        End If
		      End If
		    End If
		  End If
		Next
		txt1 = LCase( Mid(res,1,1) )
		i = Len(res) - 1
		if i < 0 then
			Repository.WriteOutput "Script", "Error: Unable to construct NCName for code: [" & attr.Name & "]",0
		else
			txt2 =Mid(res,2,i)
			txt = txt1 + txt2
			if txt <> attr.Name then
				Repository.WriteOutput "Script", "Change: Old code: [" & attr.Name & "] changed to new NCName: [" & txt & "]",0
				' return txt
				attr.Name = txt
				attr.Update()
			end if
		end if

End Sub

Sub settKodensNavnTilEgen_Navn(attr)
		' make name legal NCName by replacing each bad character with a "_", typically used for codelist with proper names.)

    Dim txt, res, tegn, i, u
    u=0
		txt = Trim(attr.Name)
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
		Repository.WriteOutput "Script", "New NCName: " & res,0
		' return res
		attr.Name = res
		attr.Update()

End Sub

Function ikkelinjeskiftEllerEnkeltfnutt(streng)
    Dim txt, res, tegn, i, u
        u = 0
		txt = ""
		' loop gjennom alle tegn og ta bort linjeskift og lignende lavverditegn
		For i = 1 To Len(streng)
            tegn = Mid(streng,i,1)
			if tegn >  " " then
			    if tegn =  "'" or tegn = "’" then
			       txt = txt + """"
			    else
			       txt = txt + tegn
			    end if
			    u = 0
			else
			    if u = 0 then
			      txt = txt + " "
			    end if
			    u = 1
			end if
        next
  ikkelinjeskiftEllerEnkeltfnutt = txt
End Function

sub oppdaterKoderForEnValgtKodeliste()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
	Repository.CreateOutputTab "Error"
	Repository.ClearOutput "Error"

	'Get the currently selected CodeList in the tree to work on

	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()

	
	if not theElement is nothing  then
		if (theElement.ObjectType = otElement) then
			dim box, mess
			mess = 	"List codelist codes to CSV. Script version 2021-10-01." & vbCrLf
			mess = mess + "Element: "& vbCrLf & "[«" & theElement.Stereotype & "» " & theElement.Name & "]."

			box = Msgbox (mess, vbOKCancel)
			select case box
			case vbOK
				if theElement.Type="Class" and LCase(theElement.Stereotype) = "codelist" Or LCase(theElement.Stereotype) = "enumeration" then
					'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
					fixOldCodelists(theElement)
				Else
					'Other than CodeList selected in the tree
					MsgBox( "This script requires a class with stereotype «CodeList» to be selected in the Project Browser." & vbCrLf & _
					"Please select a  CodeList class in the Project Browser and try once more." )
				end If
			'	Repository.WriteOutput "Script", "-----------------------" & Now & " Finished, check the Error and Types tabs", 0
				Repository.EnsureOutputVisible "Script"
			case VBcancel
						
			end select 
		else
			MsgBox( "This script requires a CodeList class element to be selected in the Project Browser." & vbCrLf & _
			"Please select a  CodeList class in the Project Browser and try once more." )
		end if
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires some CodeList class to be selected in the Project Browser." & vbCrLf & _
	  "Please select a  CodeList class in the Project Browser and try again." )
	end if

	
	
end sub

oppdaterKoderForEnValgtKodeliste
