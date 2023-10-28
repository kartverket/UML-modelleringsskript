option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		makeAbstractSchema
' purpose:		switch known stereotypes, from FeatureType to GI_Class etc.
' formål:		endre kjente stereotyper for revisjon av sosi fagområder
' author:		Kent
' version:		2023-10-27	classes and attributes handled, coldelists with codes become GI_Enumerations
' version:		2023-10-28	roles with role names handled
'				TBD: complete enumerations (<memo>)
'				TBD: remove old stereotype from package and class! (only possible with sql?)
'				TBD: tagged values on GI_Property and GI_Enum not visible under Properties tab, just stereotype name (however scripts will work)
'				TBD: fix æøå in Notes also at this stage?

		DIM debug,txt
		debug = true

sub makeAbstract()
	Repository.EnsureOutputVisible "Script"

	Dim theElement as EA.Element
	dim conn as EA.Connector
	dim connend as EA.ConnectorEnd
	Dim newETag as EA.TaggedValue
	Dim newATag as EA.AttributeTag
	Dim newCTag as EA.RoleTag
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then

		dim box
		box = Msgbox ("Skript makeAbstractSchema" & vbCrLf & vbCrLf & "Skriptversjon 2023-10-27" & vbCrLf & _
			"Endrer kjente stereotyper for modernisering av elementene i sosi fagområde: [" & theElement.Name & "]."  & vbCrLf & _
			"NOTE! This script will automatically change your model!",1)
		select case box
			case vbOK
			
				Repository.ClearOutput "Script"
				Repository.CreateOutputTab "Error"
				Repository.ClearOutput "Error"			
				
				if Repository.GetTreeSelectedItemType() = otPackage then
					makeAbstractSchema(theElement)
				else 
					if Repository.GetTreeSelectedItemType() = otElement and theElement.Type = "Class" then
						makeAbstractClass(theElement)
					else
						MsgBox( "This script requires a package or a class to be selected in the Project Browser." & vbCrLf & _
							"Please select this and try again." )			
					end if
				end if
				Repository.WriteOutput "Script", Now & " End of processing.",0

			case VBcancel
		end select

	else
		MsgBox( "This script requires a package or a class to be selected in the Project Browser." & vbCrLf & _
	  "Please select this and try again." )
	end if

end sub

sub makeAbstractSchema(pkg)
	dim elements as EA.Collection
	dim i
	if debug then Repository.WriteOutput "Script", Now & " Stereotypes will be changed for Package : [«" & pkg.element.Stereotype & "» " & pkg.Name & "].",0 end if
'	Repository.WriteOutput "Script", Now & " Stereotype LCase : " & LCase(pkg.element.Stereotype) & ".",0

	if LCase(pkg.element.Stereotype) = "applicationschema" then
	'	theElement.GetStereotypeList()
'		Repository.WriteOutput "Script", "FQ  Stereotype [" & pkg.element.FQStereotype & "]" & vbCrLf ,0
	'?	Repository.WriteOutput "Script", "Has Stereotype [" & pkg.element.HasStereotype() & "]" & vbCrLf ,0
'?		pkg.element.Stereotype = ""
	'?	pkg.element.Stereotype.Remove()
		pkg.element.Stereotype = "AbstractSchema"
		pkg.element.Update()
		pkg.element.Refresh()
		Repository.WriteOutput "Script", " Stereotype changed for Package : [«" & pkg.element.Stereotype & "» " & pkg.Name & ".",0
	end if
	
 	set elements = pkg.Elements
 	for i = 0 to elements.Count - 1 
		dim element as EA.Element 
		set element = elements.GetAt( i ) 	
		if element.Type = "Class" then
			makeAbstractClass(element)
			element.Refresh()
		end if
	next
	
	
	dim subP as EA.Package
	
	for each subP in pkg.packages
	    makeAbstractSchema(subP)
	next
end sub


sub makeAbstractClass(theElement)

'	Dim theElement as EA.Element
	dim attr as EA.Attribute
	dim conn as EA.Connector
	dim connend as EA.ConnectorEnd
	Dim newETag as EA.TaggedValue
	Dim newATag as EA.AttributeTag
	Dim newCTag as EA.RoleTag

	if debug then Repository.WriteOutput "Script", " Stereotypes will be changed for Class : [«" & theElement.Stereotype & "» " & theElement.Name & "].",0 end if
			
	'			Repository.WriteOutput "Script", "Old Stereotype [" & theElement.Stereotype & "]" & vbCrLf ,0
	'			Repository.WriteOutput "Script", "FQ  Stereotype [" & theElement.FQStereotype & "]" & vbCrLf ,0
	'			Repository.WriteOutput "Script", "StereotypeList [" & theElement.GetStereotypeList() & "]" & vbCrLf ,0
	'			Repository.WriteOutput "Script", "StereotypeEx   [" & theElement.StereotypeEx & "]" & vbCrLf ,0
				
	if LCase(theElement.Stereotype) = "featuretype" then
	'	theElement.GetStereotypeList()
'		Repository.WriteOutput "Script", "FQ featuretype Stereotype [" & theElement.FQStereotype & "]" & vbCrLf ,0
		theElement.Stereotype = "GI_Class"
		for each attr in theElement.Attributes
			makeAbstractAttribute(attr)
			theElement.Update()
			theElement.Refresh()
		next
	end if
	if LCase(theElement.Stereotype) = "datatype" then
	'	theElement.GetStereotypeList()
'		Repository.WriteOutput "Script", "FQ datatype Stereotype [" & theElement.FQStereotype & "]" & vbCrLf ,0
		theElement.Stereotype = "GI_DataType"
		for each attr in theElement.Attributes
			makeAbstractAttribute(attr)
		next
		theElement.Update()
		theElement.Refresh()
	end if
	if LCase(theElement.Stereotype) = "codelist" then
	'	theElement.GetStereotypeList()
'		Repository.WriteOutput "Script", "FQ codelist Stereotype [" & theElement.FQStereotype & "]" & vbCrLf ,0
		if theElement.Attributes.Count() > 0 then
			theElement.Stereotype = "enumeration"
		else
			theElement.Stereotype = "GI_CodeSet"
		end if
		if theElement.Attributes.Count() > 0 then
			Repository.WriteOutput "Script", "Warning: Class with codes not empty! ["  & theElement.Name & " has "& theElement.Attributes.Count() & " codes]" & vbCrLf ,0
			for each attr in theElement.Attributes
				makeAbstractAttribute(attr)
			next
		end if
		theElement.Update()
		theElement.Refresh()
	end if
'	Repository.WriteOutput "Script", "New Stereotype [" & theElement.Stereotype & "]" & vbCrLf ,0			
	Repository.WriteOutput "Script", " Stereotype changed for Class : [«" & theElement.Stereotype & "» " & theElement.Name & "].",0




				
	for each conn in theElement.Connectors
		'''Repository.WriteOutput "Script", "Connector found [" & theElement.Name & "]"  & " Connector.Name [" & conn.Name & "]" & " Connector.Type [" & conn.Type & "]" & vbCrLf ,0
		if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then
		else
		'	find roles referring to the other class
			if conn.ClientID = theElement.ElementID then
				if debug then Repository.WriteOutput "Script", "Connector to supplier [" & theElement.Name & "]"  & " conn.SupplierEnd.Role [" & conn.SupplierEnd.Role & "] to Class [" & Repository.GetElementByID(conn.SupplierID).Name & "]" & vbCrLf ,0 end if
				set connend = conn.SupplierEnd
			else
				if debug then Repository.WriteOutput "Script", "Connector to client [" & theElement.Name & "]"  & " conn.ClientEnd.Role [" & conn.ClientEnd.Role & "]" & " to Class [" & Repository.GetElementByID(conn.ClientID).Name & "]" & vbCrLf ,0 end if
				set connend = conn.ClientEnd
			end if

			if debug then Repository.WriteOutput "Script", "Class [" & theElement.Name & "]"  & " role.Name [" & connend.Role & "]"  & " role.Stereotype [" & connend.Stereotype & "]" & " role.StereotypeEx [" & connend.StereotypeEx & "]" & vbCrLf ,0  end if
			if connend.Role <> "" then
				connend.Stereotype = "GI_Property"
				connend.Update()
			end if
		end if
	next
	theElement.Update()
	theElement.Refresh()
			
			


end sub

sub makeAbstractAttribute(attr)

'	Dim theElement as EA.Element
'	dim attr as EA.Attribute
	dim conn as EA.Connector
	dim connend as EA.ConnectorEnd
	Dim newETag as EA.TaggedValue
	Dim newATag as EA.AttributeTag
	Dim newCTag as EA.RoleTag

			
	'			Repository.WriteOutput "Script", "Old Stereotype [" & attr.Stereotype & "]" & vbCrLf ,0
	'			Repository.WriteOutput "Script", "FQ  Stereotype [" & attr.FQStereotype & "]" & vbCrLf ,0
	'			Repository.WriteOutput "Script", "StereotypeList [" & theElement.GetStereotypeList() & "]" & vbCrLf ,0
	'			Repository.WriteOutput "Script", "StereotypeEx   [" & attr.StereotypeEx & "]" & vbCrLf ,0
				
	if LCase(attr.Stereotype) = "egenskap" then
	'	theElement.GetStereotypeList()
'		Repository.WriteOutput "Script", "FQ egenskap Stereotype [" & attr.FQStereotype & "]" & vbCrLf ,0
		attr.Stereotype = "GI_Property"
		attr.Update()
	end if
	if LCase(attr.Stereotype) = "enum" or LCase(attr.Type) = "<undefined>" or LCase(attr.Type) = "" then
	'	theElement.GetStereotypeList()
'		Repository.WriteOutput "Script", "FQ enum Stereotype [" & attr.FQStereotype & "]" & vbCrLf ,0
		attr.Stereotype = "GI_EnumerationLiteral"
		attr.Update()
	end if
	if LCase(attr.Stereotype) = "" and LCase(attr.Type) <> "" then
	'	theElement.GetStereotypeList()
'		Repository.WriteOutput "Script", "FQ  Stereotype [" & attr.FQStereotype & "]" & vbCrLf ,0
		attr.Stereotype = "GI_Property"
		attr.Update()
	end if

'	Repository.WriteOutput "Script", "New Stereotype [" & attr.Stereotype & "]" & vbCrLf ,0			
	if LCase(attr.Stereotype) = "enum" or LCase(attr.Stereotype) = "GI_EnumerationLiteral" or LCase(attr.Type) = "<undefined>" or LCase(attr.Type) = "" then
	else
		Repository.WriteOutput "Script", " Stereotype changed for Atribute : [«" & attr.Stereotype & "» " & attr.Name & "].",0
	end if

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

function getConnectorEndTaggedValue(connectorEnd,taggedValueName)
	getConnectorEndTaggedValue = ""
	if not connectorEnd is nothing and Len(taggedValueName) > 0 then
		dim existingTaggedValue as EA.RoleTag 
		dim i
		for i = 0 to connectorEnd.TaggedValues.Count - 1
			set existingTaggedValue = connectorEnd.TaggedValues.GetAt(i)
			if existingTaggedValue.Tag = taggedValueName then
				getConnectorEndTaggedValue = existingTaggedValue.Value
			end if 
		next
	end if 
end function 


function iso8601date(str)
	iso8601date = ""
	if Len(str) >= 10 then
		iso8601date = Mid(str,7,4) & "-" & Mid(str,4,2) & "-" & Mid(str,1,2)
	end if
end function


function readutf8(str)
	' make string utf-8
	Dim txt, res, tegn, utegn, vtegn, wtegn, xtegn, i
	Dim c3, c4, c5, c6, c7, ca, e2
	
	readutf8 = ""
	'exit function
	
	'txt = Trim(str)
	txt = str
		c3 = 195 ' vanligste startbyte for utf8
			c4 = 196
			c5 = 197
			c6 = 198
			c7 = 199
			ca = 202
			e2 = 226
'		c3 = 195 b6 = 182 ö
'		c3 = 195 a1 = 161 á
'		c5 = 197 8b = 139 ŋ
		
	' loop gjennom alle tegn
	res = ""
	i = 0
	While i < Len(txt)
		i = i + 1
		tegn = Mid(txt,i,1)
'		Repository.WriteOutput "Script", "readutf8: i=" & i & " tegn=" & tegn,0
		
		Select case int(AscW(tegn))
			Case c3
				'Repository.WriteOutput "Script", "readutf8: c3A i=" & i & " tegn=" & tegn & "int(AscW(tegn))=" & int(AscW(tegn)) ,0
				'Repository.WriteOutput "Script", "readutf8: c3B int(AscW((Mid(txt,i+1,1)))=" & int(AscW(Mid(txt,i+1,1))) ,0
				
				i = i + 1
				Select case int(AscW(Mid(txt,i,1)))
					Case 134
						res=res+"Æ"
						
					Case 152
						res=res+"Ø"
					Case 732
						res=res+"Ø"

					Case 133
						res=res+"Å"
					Case 8230
						res=res+"Å"

					Case 166
						res=res+"æ"

					Case 184
						res=res+"ø"

					Case 165
						res=res+"å"
						
					Case 129
						res=res+"Á"
						
					Case 161
						res=res+"á"
						
					Case 182
						res=res+"ö"

					Case else
						utegn = int(AscW(tegn)) & 511
						vtegn = utegn * 64
						wtegn = int(AscW(Mid(txt,i+1,1))) & 1023
						xtegn = wtegn
						Repository.WriteOutput "Script", "readutf8: c3 i=" & i & " tegn=" & tegn & " " & int(AscW(Mid(txt,i,1))) & " -> " & utegn & " + " & vtegn & " + " & wtegn & " + " & xtegn & " + " & res & " " & AscW(vtegn)+AscW(xtegn),0
		'				res=res+Chr(AscW(vtegn)+AscW(xtegn))
						res=res+"Ã"+Mid(txt,i,1)
				End Select
			'Case c4
			
			Case c5
				'Repository.WriteOutput "Script", "readutf8: c5A i=" & i & " tegn=" & tegn & "int(AscW(tegn))=" & int(AscW(tegn)) ,0
				'Repository.WriteOutput "Script", "readutf8: c5B int(AscW((Mid(txt,i+1,1)))=" & int(AscW(Mid(txt,i+1,1))) ,0
			
				i = i + 1
				Select case int(AscW(Mid(txt,i,1)))

					Case 160
						res=res+"Š"
						
					Case 161
						res=res+"š"
						
					Case 138
						res=res+"Ŋ"
						
					Case 139
						res=res+"ŋ"
						
		'			Case 8249
		'				res=res+"ŋ"
						
					Case else
						utegn = int(AscW(tegn)) & 511
						vtegn = utegn * 64
						wtegn = int(AscW(Mid(txt,i+1,1))) & 1023
						xtegn = wtegn
						Repository.WriteOutput "Script", "readutf8: c5 i=" & i & " tegn=" & tegn & " " & Mid(txt,i,1) & " " & int(AscW(Mid(txt,i,1))) & " -> " & utegn & " + " & vtegn & " + " & wtegn & " + " & xtegn & " + " & res & " " & AscW(vtegn)+AscW(xtegn),0
		'				res=res+Chr(AscW(vtegn)+AscW(xtegn))
						res=res+"Ã"+Mid(txt,i,1)
						
					End Select
				
			'Case c6
			
			'Case c7
			
			'Case ca

			Case e2
				' + 80 + 93 = "en dash" = halv fonthøyde, ( + 80 + 94 = "em dash" = full fonthøyde)
				i = i + 2
				res=res+"–"
			Case else
				res=res+tegn
		
		End Select

	Wend
	readutf8 = res



End function


makeAbstract
