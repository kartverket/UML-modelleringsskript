option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		backupTaggedValues
' purpose:		make backup of all tagged values for classes so changing stereotype will not loose the values
' formål:		ta vare på tagged values for klasser så ikke bytte av stereotype vil miste dem
' author:		Kent
' version:		2023-11-02	fix bugs in connection end
' version:		2023-09-27, 10-27, 10-29 refactoring into subs that later can be combined into one script
'				TBD: multipple runs
'				TBD: multipple tags with same name
'				TBD: packages
'				TBD: logic for empty tags (no value)

		DIM debug,txt
		debug = false

sub backupTaggedValues()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"

	Dim theElement as EA.Element
	dim conn as EA.Connector
	dim connend as EA.ConnectorEnd
	Dim newETag as EA.TaggedValue
	Dim newATag as EA.AttributeTag
	Dim newCTag as EA.RoleTag
	dim existingRoleTaggedValue as EA.RoleTag 
	Set theElement = Repository.GetTreeSelectedObject()
	dim message, box
	if not theElement is nothing  then
		if Repository.GetTreeSelectedItemType() = otPackage then
			'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
				'tømmer System Output for lettere å fange opp eventuelle feilmeldinger der
				Repository.ClearOutput "Script"
				Repository.CreateOutputTab "Error"
				Repository.ClearOutput "Error"

			box = Msgbox ("Skript backupTaggedValues" & vbCrLf & vbCrLf & "Skriptversjon 2023-09-29" & vbCrLf & "Lager kopier av alle tagged values for pakken: [" & theElement.Name & "].",1)
			select case box
			case vbOK
				backupTaggedValuesPackage(theElement)
			case VBcancel

			end select
			Repository.WriteOutput "Script", Now & " End of processing.",0

		else
		
			if Repository.GetTreeSelectedItemType() = otElement and theElement.Type = "Class" then
				'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
					'tømmer System Output for lettere å fange opp eventuelle feilmeldinger der
					Repository.ClearOutput "Script"
					Repository.CreateOutputTab "Error"
					Repository.ClearOutput "Error"

				box = Msgbox ("Skript backupTaggedValues" & vbCrLf & vbCrLf & "Skriptversjon 2023-09-29" & vbCrLf & "Lager kopier av alle tagged values for klassen: [" & theElement.Name & "].",1)
				select case box
				case vbOK
					backupTaggedValuesClass(theElement)
				case VBcancel

				end select
				Repository.WriteOutput "Script", Now & " End of processing.",0
			else
			
		  'Other than CodeList selected in the tree
		  MsgBox( "This script requires a package class to be selected in the Project Browser." & vbCrLf & _
			"Please select in the Project Browser and try again." )
			end If
		end If
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires a package or class to be selected in the Project Browser." & vbCrLf & _
	  "Please select in the Project Browser and try again." )
	end if		
		
end sub


sub backupTaggedValuesPackage(pkg)
	' Show and clear the script output window
'	Repository.EnsureOutputVisible "Script"

'	Dim theElement as EA.Element
'	dim x as EA.Package
	dim elements as EA.Collection
	dim existingTaggedValue as EA.TaggedValue
	dim i, backupValue
	Dim newETag as EA.TaggedValue
	dim theElement as EA.Element 

	Repository.WriteOutput "Script", Now & " Tagged values will be backed up for Package : " & pkg.Name & ".",0

	for i = 0 to pkg.Element.TaggedValues.Count - 1
		backupValue = ""
		set existingTaggedValue = pkg.Element.TaggedValues.GetAt(i)
		backupValue = existingTaggedValue.Value
		' not copying already backed up tags
		if Mid(existingTaggedValue.Name,1,7) <> "backup-" then
			'only one unique tag name backup-xxx with content
			if getTaggedValue(pkg.Element,"backup-" + existingTaggedValue.Name) = "" then
				set newETag = pkg.Element.TaggedValues.AddNew("backup-" + existingTaggedValue.Name,"Tag")
				newETag.Value = backupValue
				newETag.Update()
				Repository.WriteOutput "Script", "New Tag added to Package [" & pkg.Name & "]" & " New TaggedValue.Name [" & existingTaggedValue.Name & "]" & " Value [" & getTaggedValue(pkg.Element,"backup-" + existingTaggedValue.Name) & "]" & vbCrLf ,0
			else
				Repository.WriteOutput "Script", "Warning: Duplicate tag in Package [" & pkg.Name & "]"  & " TaggedValue.Name [" & "backup-" + existingTaggedValue.Name & "]" & " Value [" & backupValue & "]" & vbCrLf ,0
			end if
		end if
	next
	pkg.Element.Refresh()


 	set elements = pkg.Elements
 	for i = 0 to elements.Count - 1 
		set theElement = pkg.Elements.GetAt( i ) 	
		if theElement.Type = "Class" then
			backupTaggedValuesClass(theElement)
			theElement.Refresh()
		end if
	next
	
	
	dim subP as EA.Package
	
	for each subP in pkg.packages
	    backupTaggedValuesPackage(subP)
	next

end sub

sub backupTaggedValuesClass(theElement)
	' Show and clear the script output window
'	Repository.EnsureOutputVisible "Script"

'	Dim theElement as EA.Element
	dim conn as EA.Connector
	dim connend as EA.ConnectorEnd
	Dim newETag as EA.TaggedValue
	Dim newATag as EA.AttributeTag
	Dim newCTag as EA.RoleTag
	dim existingRoleTaggedValue as EA.RoleTag 


	Repository.WriteOutput "Script", Now & " Tagged values will be backed up for Class : " & theElement.Name & ".",0

	dim i, existingTaggedValue, backupValue
	for i = 0 to theElement.TaggedValues.Count - 1
		backupValue = ""
		set existingTaggedValue = theElement.TaggedValues.GetAt(i)
		backupValue = existingTaggedValue.Value
		' not copying already backed up tags
		if Mid(existingTaggedValue.Name,1,7) <> "backup-" then
			'only one unique tag name backup-xxx with content
			if getTaggedValue(theElement,"backup-" + existingTaggedValue.Name) = "" then
				set newETag = theElement.TaggedValues.AddNew("backup-" + existingTaggedValue.Name,"Tag")
				newETag.Value = backupValue
				newETag.Update()
				Repository.WriteOutput "Script", "New Tag added [" & theElement.Name & "]" & " New TaggedValue.Name [" & existingTaggedValue.Name & "]" & " Value [" & getTaggedValue(theElement,"backup-" + existingTaggedValue.Name) & "]" & vbCrLf ,0
			else
				Repository.WriteOutput "Script", "Warning: Duplicate tag in Class [" & theElement.Name & "]"  & " TaggedValue.Name [" & "backup-" + existingTaggedValue.Name & "]" & " Value [" & backupValue & "]" & vbCrLf ,0
			end if
		end if
	next
	theElement.Refresh()


	dim attr as EA.Attribute
	for each attr in theElement.Attributes
		for i = 0 to attr.TaggedValues.Count - 1
			backupValue = ""
			set existingTaggedValue = attr.TaggedValues.GetAt(i)
			backupValue = existingTaggedValue.Value
			Repository.WriteOutput "Script", "Class [" & theElement.Name & "]"  & " attr.Name [" & attr.Name & "]"  & " TaggedValue.Name [" & existingTaggedValue.Name & "]" & " Value [" & backupValue & "]" & vbCrLf ,0
			' not copying already backed up tags
			if Mid(existingTaggedValue.Name,1,7) <> "backup-" then
				' only one unique tag name backup-xxx with content
				if getTaggedValue(attr,"backup-" + existingTaggedValue.Name) = "" then
					set newATag = attr.TaggedValues.AddNew("backup-" + existingTaggedValue.Name,"Tag")
					newATag.Value = backupValue
					newATag.Update()
					Repository.WriteOutput "Script", "New Tag added [" & theElement.Name & "]" & " attr.Name [" & attr.Name & "]" & " New TaggedValue.Name [" & newATag.Name & "]" & " Value [" & backupValue & "]" & vbCrLf ,0
				else
					Repository.WriteOutput "Script", "Warning: Duplicate tag in Class [" & theElement.Name & "]"  & " attr.Name [" & attr.Name & "]"  & " TaggedValue.Name [" & existingTaggedValue.Name & "]" & " Value [" & backupValue & "]" & vbCrLf ,0
				end if
			end if
		next
	next
		
	for each conn in theElement.Connectors
		'''Repository.WriteOutput "Script", "Connector found [" & theElement.Name & "]"  & " Connector.Name [" & conn.Name & "]" & " Connector.Type [" & conn.Type & "]" & vbCrLf ,0
		if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then
		else
		'find roles referring to the other class
			if conn.ClientID = theElement.ElementID then
				if debug then Repository.WriteOutput "Script", "Connector to supplier [" & theElement.Name & "]"  & " conn.SupplierEnd.Role [" & conn.SupplierEnd.Role & "] to Class [" & Repository.GetElementByID(conn.SupplierID).Name & "]" & vbCrLf ,0 end if
				set connend = conn.SupplierEnd
			else
				if debug then Repository.WriteOutput "Script", "Connector to client [" & theElement.Name & "]"  & " conn.ClientEnd.Role [" & conn.ClientEnd.Role & "]" & " to Class [" & Repository.GetElementByID(conn.ClientID).Name & "]" & vbCrLf ,0 end if
				set connend = conn.ClientEnd
			end if
			for i = 0 to connend.TaggedValues.Count - 1
				backupValue = ""
				set existingRoleTaggedValue = connend.TaggedValues.GetAt(i)
				backupValue = existingRoleTaggedValue.Value
				if debug then Repository.WriteOutput "Script", "Class [" & theElement.Name & "]"  & " role.Name [" & connend.Role & "]"  & " existingRoleTaggedValue.Tag [" & existingRoleTaggedValue.Tag & "]" & " Value [" & backupValue & "]" & vbCrLf ,0  end if
				' not copying already backed up tags
				if Mid(existingRoleTaggedValue.Tag,1,7) <> "backup-" then			
					'only one unique tag name backup-xxx with content
					if getConnectorEndTaggedValue(connend, "backup-" & existingRoleTaggedValue.Tag) = "" then
						set newCTag = connend.TaggedValues.AddNew("backup-" + existingRoleTaggedValue.Tag,"Tag")
						newCTag.Value = backupValue
						newCTag.Update()
								connend.Update()
								conn.Update()
								theElement.Update()
								theElement.Refresh()
						Repository.WriteOutput "Script", "New Tag added [" & theElement.Name & "]" & " role.Name [" & connend.Role & "]" & " New RoleTaggedValue.Tag [" & newCTag.Tag & "]" & " Value [" & backupValue & "]" & vbCrLf ,0
					else
						Repository.WriteOutput "Script", "Warning: Duplicate tag in theElement.Name [" & theElement.Name & "]"  & " role.Name [" & connend.Role & "]"  & " existingRoleTaggedValue.Tag [" & existingRoleTaggedValue.Tag & "]" & " Value [" & backupValue & "]" & vbCrLf ,0
					end if
					
				end if
			next
				
		end if
	next
	theElement.Refresh()
			
			

	Repository.WriteOutput "Script", Now & " Tagged values backed up for Class : " & theElement.Name & ".",0


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


backupTaggedValues
