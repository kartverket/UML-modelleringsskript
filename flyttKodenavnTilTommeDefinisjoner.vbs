option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		flyttKodenavnTilTommeDefinisjoner
' purpose:		Flytter tekst generert fra kodens navn inn i tomme definisjoner til kodelistekoder
' version:		2019-07-05, 07-09
' author: 		Kent Jonsrud
'
' ex:
' Kodelistekode [«codeList» VA_RørmaterialeAlle.rustfrittStål] har fått definisjon [rustfritt stål].
' Kodelistekode [«codeList» VA_RørmaterialeAlle.PVC] har fått definisjon [p v c].

sub flyttKodenavnTilTommeDefinisjoner()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"

	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		dim message, indent
		dim box
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		if Repository.GetTreeSelectedItemType() = otPackage then
			if UCASE(theElement.Element.Stereotype) = "APPLICATIONSCHEMA" then
			Repository.WriteOutput "Script", Now & " " & theElement.Element.Stereotype & " " & theElement.Name, 0
				message = "Script flyttKodenavnTilTommeDefinisjoner" & vbCrLf & vbCrLf & "Scriptversion 2019-07-09" & vbCrLf & "Fyller alle tomme definisjoner i kodelistekoder fra kodenes navn, for pakke: " & vbCrLf & "[" & theElement.Name & "]."
				message = message & vbCrLf & vbCrLf & "NB: Ekspanderer tekniske navn til setninger med blanke, men tar ikke hensyn til forkortelser!!!"
				box = Msgbox (message,1)
				select case box
				case vbOK
					SessionOutput("Script flyttKodenavnTilTommeDefinisjoner, Scriptversion 2019-07-09. ")
					SessionOutput("Fyller alle tomme kodelistekodedefinisjoner for pakke: [" & theElement.Name & "] " & Now() & ".")
				
					call flyttAlleKodenavnIPakke(theElement)
				
					SessionOutput("Finished filling empty definitions in codelists.")

				case VBcancel

				end select
			else
				MsgBox( "This script requires an «ApplicationSchema» Package or a «CodeList» Class to be selected in the Project Browser." & vbCrLf & _
				"Please select an «ApplicationSchema» Package or a «CodeList» Class in the Project Browser and try again." )
			end if
		Else
			if Repository.GetTreeSelectedItemType() = otElement then
				if theElement.Type="Class" and UCASE(theElement.Stereotype) = "CODELIST" then
					message = "Script flyttKodenavnTilTommeDefinisjoner" & vbCrLf & vbCrLf & "Scriptversion 2019-07-09" & vbCrLf & "Fyller alle tomme definisjoner i kodelistekoder fra kodenes navn, for klasse: " & vbCrLf & "[" & theElement.Name & "]."
					message = message & vbCrLf & vbCrLf & "NB: Ekspanderer tekniske navn til setninger med blanke, men tar ikke hensyn til forkortelser!!!"
					box = Msgbox (message,1)
					'box = Msgbox ("Script flyttKodenavnTilTommeDefinisjoner" & vbCrLf & vbCrLf & "Scriptversion 2019-07-09" & vbCrLf & "Fyller alle tomme definisjoner i kodelistekoder fra kodenes navn, for klasse: " & vbCrLf & "[" & theElement.Name & "].",1)
					select case box
					case vbOK
						SessionOutput("Script flyttKodenavnTilTommeDefinisjoner, Scriptversion 2019-07-09. ")
						SessionOutput("Fyller alle tomme kodelistekodedefinisjoner for klasse: [" & theElement.Name & "] " & Now() & ".")
					
						call flyttAlleKodenavnIKlasse(theElement)
				
					SessionOutput("Finished filling empty definitions in codelist.")
					
					case VBcancel

					end select
				else
					MsgBox( "This script requires an «ApplicationSchema» Package or a «CodeList» Class to be selected in the Project Browser." & vbCrLf & _
					"Please select an «ApplicationSchema» Package or a «CodeList» Class in the Project Browser and try again." )
				end if
			else
				'Other than «ApplicationSchema» Package or a «FeatureType» Class selected in the tree
				MsgBox( "Element type selected: " & theElement.Type & vbCrLf & _
				"This script requires an «ApplicationSchema» Package or a «CodeList» Class to be selected in the Project Browser." & vbCrLf & _
				"Please select an «ApplicationSchema» Package or a «CodeList» Class in the Project Browser and try again." )
			end If
		end if
	end if
end sub


sub flyttAlleKodenavnIPakke(pkg)
 	dim elements as EA.Collection 
	set elements = pkg.Elements 
	dim i
	dim indent, ftname
	for i = 0 to elements.Count - 1 
		dim currentElement as EA.Element 
		set currentElement = elements.GetAt( i ) 
		'SessionOutput( "Debug: Class [«" & currentElement.Stereotype & "» " & currentElement.Name & "] currentElement.Type [" & currentElement.Type & "].")
		if currentElement.Type = "Class" and LCase(currentElement.Stereotype) = "codelist" then
			
			call flyttAlleKodenavnIKlasse(currentElement)
			
		end if
	next

	dim subP as EA.Package
	for each subP in pkg.packages
	    call flyttAlleKodenavnIPakke(subP)
	next

end sub

sub flyttAlleKodenavnIKlasse(currentElement)
	dim def
	dim attr as EA.Attribute
	for each attr in currentElement.Attributes
		'SessionOutput( "Debug:                  attr.Name [«" & attr.Stereotype & "» " & attr.Name & "] attr.Type [" & attr.Type & "] attr.Notes [" & attr.Notes & "].")

		if attr.Notes = "" then
			'def = attr.Name
			def = getCleanDefinitionText(attr.Name)
			if getTaggedValue(attr,"NVDB_navn") <> "" then
				def = getTaggedValue(attr,"NVDB_navn")
			end if
			if getTaggedValue(attr,"SOSI_presentasjonsnavn") <> "" then
				def = getTaggedValue(attr,"SOSI_presentasjonsnavn")
			end if
			
			if attr.Stereotype <> "" then
				def = def & " (" & attr.Stereotype & ")"
			end if
			attr.Notes = def
			attr.Update()
			'SessionOutput( "Debug:             -->  attr.Notes [" & def & "].")
			SessionOutput( "Kodelistekode [«" & currentElement.Stereotype & "» " & currentElement.Name & "." & attr.Name & "] har fått definisjon [" & attr.Notes & "].")
		end if
	next

end sub


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

sub SessionOutput(text)
	Session.Output(text)
end sub

function getCleanDefinitionText(txt)
	'expands NCName to normal text for definition
    Dim res, tegn, i
	getCleanDefinitionText = ""
		res = ""
		' loop gjennom alle tegn
		For i = 1 To Len(txt)
			tegn = Mid(txt,i,1)
			If tegn <> LCase(tegn) Then
				if i > 1 then
					res = res + " "
				end if
				res = res + LCase(tegn)
			Else 
				res = res + tegn
			End If
		Next
		
	getCleanDefinitionText = res

end function

flyttKodenavnTilTommeDefinisjoner
