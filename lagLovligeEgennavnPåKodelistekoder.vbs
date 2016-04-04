option explicit

!INC Local Scripts.EAConstants-VBScript

' skriptnavn:         lagLovligeEgennavnPÂKodelistekoder

sub fixOldCodelists(el)
	'Repository.WriteOutput "Script", Now & " CodeList: " & el.Name, 0
	Repository.WriteOutput "Script", Now & " " & el.Stereotype & " " & el.Name, 0

	dim attr as EA.Attribute
	for each attr in el.Attributes
		Repository.WriteOutput "Script", Now & " " & el.Name & "." & attr.Name, 0

    kopierKodensNavnTilTomDefinisjon(attr)

		kopierKodensNavnTilTagSOSI_presentasjonsnavn(attr)

    'flyttInitialverdiTilTagSOSI_verdi(attr)

    'settKodensNavnTilNCName(attr)
    'eller
    settKodensNavnTilEgen_Navn(attr)


	next

end sub

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

Sub settKodensNavnTilNCName(attr)
		' make name legal NCName
		' (alternatively replace each bad character with a "_", typically used for codelist with proper names.)
		' (Sub settBlankeIKodensNavnTil_(attr))
    Dim txt, res, tegn, i, u
    u=0
		txt = Trim(attr.Name)
		res = LCase( Mid(txt,1,1) )
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
		      If tegn = "]" or tegn = "^" or tegn = "`" or tegn = "{" or tegn = "|" or tegn = "}" or tegn = "~" or tegn = "'" or tegn = "¥" or tegn = "®" Then
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
		Repository.WriteOutput "Script", "New NCName: " & res,0
		' return res
		attr.Name = res
		attr.Update()

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
		      If tegn = "]" or tegn = "^" or tegn = "`" or tegn = "{" or tegn = "|" or tegn = "}" or tegn = "~" or tegn = "'" or tegn = "¥" or tegn = "®" Then
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
			if ((theElement.Type = "Class") and (theElement.Stereotype = "codeList" Or theElement.Stereotype = "CodeList" Or theElement.Stereotype = "enumeration")) then
				'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
				fixOldCodelists(theElement)
			else
				MsgBox( "This script requires a CodeList class to be selected in the Project Browser." & vbCrLf & _
				"Please select a  CodeList class in the Project Browser and try once more." )
			end if
		Else
		  'Other than CodeList selected in the tree
		  MsgBox( "This script requires a CodeList class to be selected in the Project Browser." & vbCrLf & _
			"Please select a  CodeList class in the Project Browser and try once more." )
		end If
		Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
		Repository.EnsureOutputVisible "Script"
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires a CodeList class to be selected in the Project Browser." & vbCrLf & _
	  "Please select a  CodeList class in the Project Browser and try again." )
	end if
end sub

oppdaterKoderForEnValgtKodeliste