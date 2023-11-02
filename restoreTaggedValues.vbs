option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		restoreTaggedValues
' purpose:		restores backup of all tagged values for classes so new stereotype will get the old values
' formål:		legger tilbake tagged values for klasser etter bytte av stereotype
' author:		Kent
' version:		2023-11-01 remove backup tags
' version:		2023-09-29
'				TBD: refactoring into subs that later can be combined into one script
'				TBD: make new tag x directly in class, with value from backup-x, for when tag x is not in the new stereotype / ?
'				TBD: if value <> "" rename backup-xxx tag to xxx?

	DIM debug,txt
	debug = false

sub restoreTaggedValues()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"

	Dim theElement as EA.Element
	dim conn as EA.Connector
	dim connend as EA.ConnectorEnd
	Dim newETag as EA.TaggedValue
	Dim newATag as EA.AttributeTag
	Dim newCTag as EA.RoleTag
	dim i, j, k, txt
	dim existingTaggedValue as EA.TaggedValue
	dim existingRoleTaggedValue as EA.RoleTag 

	dim	backupValue	
	Set theElement = Repository.GetTreeSelectedObject()
	
	if not theElement is nothing  then
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		'f Repository.GetTreeSelectedItemType() = otPackage then
	'	if Repository.GetTreeSelectedItemType() = otElement and UCASE(theElement.Stereotype) = "CODELIST" then
		if Repository.GetTreeSelectedItemType() = otElement and theElement.Type = "Class" then
			'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
				'tømmer System Output for lettere å fange opp eventuelle feilmeldinger der
				Repository.ClearOutput "Script"
				Repository.CreateOutputTab "Error"
				Repository.ClearOutput "Error"

			dim message
			dim box
			box = Msgbox ("Skript restoreTaggedValues" & vbCrLf & vbCrLf & "Skriptversjon 2023-09-29" & vbCrLf & "Legge tilbake kopier av alle tagged values for klassen: [" & theElement.Name & "].",1)
			select case box
			case vbOK

			Repository.WriteOutput "Script", Now & " Tagged values will be restored for Class : " & theElement.Name & ".",0

			'	copy backed up values to same named tags in new stereotype in new profile (ex: description)
			for i = 0 to theElement.TaggedValues.Count - 1
'				backupValue = ""
				set existingTaggedValue = theElement.TaggedValues.GetAt(i)
'				backupValue = existingTaggedValue.Value
				' not copying already filled tags
				if existingTaggedValue.Value = "" then
'				if Mid(existingTaggedValue.Name,1,7) <> "backup-" then
					if getTaggedValue(theElement,"backup-" + existingTaggedValue.Name) <> "" then
'						set newETag = theElement.TaggedValues.AddNew("backup-" + existingTaggedValue.Name,"Tag")
						existingTaggedValue.Value = getTaggedValue(theElement,"backup-" + existingTaggedValue.Name)
						existingTaggedValue.Update()
						theElement.Update()
						theElement.Refresh()
						Repository.WriteOutput "Script", "Old value restored on Tag " & i & " in [" & theElement.Name & "]" & " Restored TaggedValue.Name [" & existingTaggedValue.Name & "]" & " Value [" & getTaggedValue(theElement,existingTaggedValue.Name) & "]" & vbCrLf ,0
					end if
				end if
			next
			'	if value <> "" rename backup-xxx tag to xxx?
			for i = 0 to theElement.TaggedValues.Count - 1
				set existingTaggedValue = theElement.TaggedValues.GetAt(i)
				if Mid(existingTaggedValue.Name,1,7) = "backup-" and existingTaggedValue.Value <> "" then
					txt = Mid(existingTaggedValue.Name,8,Len(existingTaggedValue.Name))
					if debug then Repository.WriteOutput "Script", "May rename Tag " & i & " " & existingTaggedValue.Name & " in Class [" & theElement.Name & "] to : [" & txt & "]" & vbCrLf ,0 end if
					if getTaggedValue(theElement,txt) = "" then
						if debug then Repository.WriteOutput "Script", "Will rename Tag " & i & " " & existingTaggedValue.Name & " in Class [" & theElement.Name & "] value : [" & existingTaggedValue.Value & "]" & vbCrLf ,0 end if
						existingTaggedValue.Name = txt
						existingTaggedValue.Update()
						theElement.Update()
						theElement.Refresh()
					end if
				end if
			next
			' remove backup tags
			i = 0
			k = theElement.TaggedValues.Count - 1
			while i < k and i < theElement.TaggedValues.Count
				if debug then Repository.WriteOutput "Script", "May Delete Tag " & i & " in Class [" & theElement.Name & "]" & vbCrLf ,0 end if
				set existingTaggedValue = theElement.TaggedValues.GetAt(i)
				if Mid(existingTaggedValue.Name,1,7) = "backup-" then
					if debug then Repository.WriteOutput "Script", " Will delete tag i= [" & i & "] Name [" & existingTaggedValue.Name & "]" & vbCrLf ,0 end if
					call theElement.TaggedValues.DeleteAt(i,0)
					theElement.TaggedValues.Refresh()
					theElement.Update()
					theElement.Refresh()
					k = k - 1
				else
					i = i + 1
				end if
				if debug then Repository.WriteOutput "Script", " wend i= [" & i & "] k= [" & k & "] Name [" & existingTaggedValue.Name & "]" & vbCrLf ,0 end if
			wend
			
'			j = 0
'			k = theElement.TaggedValues.Count - 1
'			for i = 0 to k
'				if debug then Repository.WriteOutput "Script", "May Delete Tag " & i & " in Class [" & theElement.Name & "]" & vbCrLf ,0 end if
'				set existingTaggedValue = theElement.TaggedValues.GetAt(i)
'				if Mid(existingTaggedValue.Name,1,7) = "backup-" then
'					if debug then Repository.WriteOutput "Script", "Delete Tag " & i & " in Class [" & theElement.Name & "]" & " TaggedValue.Name  [" & existingTaggedValue.Name & "]" & " existingTaggedValue.Value [" & existingTaggedValue.Value & "]" & vbCrLf ,0 end if
'					call theElement.TaggedValues.DeleteAt(j,0)
'					theElement.TaggedValues.Refresh()
'					theElement.Update()
'					theElement.Refresh()
'				else
'					j = j + 1
'				end if
'				if i >= k then exit for
'			next

'			theElement.Refresh()


			dim attr as EA.Attribute
			for each attr in theElement.Attributes
				for i = 0 to attr.TaggedValues.Count - 1
'					backupValue = ""
					set existingTaggedValue = attr.TaggedValues.GetAt(i)
'					backupValue = existingTaggedValue.Value
					if debug then Repository.WriteOutput "Script", "Class [" & theElement.Name & "]"  & " attr.Name [" & attr.Name & "]"  & " TaggedValue.Name [" & existingTaggedValue.Name & "]" & " existingTaggedValue.Value [" & existingTaggedValue.Value & "]" & vbCrLf ,0 end if
					' not copying already backed up tags
					if existingTaggedValue.Value = "" then
'					if Mid(existingTaggedValue.Name,1,7) <> "backup-" then
						if getTaggedValue(attr,"backup-" + existingTaggedValue.Name) <> "" then
							existingTaggedValue.Value = getTaggedValue(theElement,"backup-" + existingTaggedValue.Name)
							existingTaggedValue.Update()
							attr.Update()
							theElement.Update()
							theElement.Refresh()
							Repository.WriteOutput "Script", "Old value restored on Tag " & i & " in [" & theElement.Name & "]" & " attr.Name [" & attr.Name & "]" & " New TaggedValue.Name [" & existingTaggedValue.Name & "]" & " existingTaggedValue.Value [" & existingTaggedValue.Value & "]" & vbCrLf ,0
						end if
					end if
				next
			'	if value <> "" rename backup-xxx tag to xxx?
				i = 0
				k = attr.TaggedValues.Count - 1
				if debug then Repository.WriteOutput "Script", " while 0 i= [" & i & "] k= [" & k & "] attr.TaggedValues.Count= [" & attr.TaggedValues.Count & "]" & vbCrLf ,0 end if
				while i < k and i < attr.TaggedValues.Count
					if debug then Repository.WriteOutput "Script", " while 1 i= [" & i & "] k= [" & k & "] attr.TaggedValues.Count= [" & attr.TaggedValues.Count & "]" & vbCrLf ,0 end if
					set existingTaggedValue = attr.TaggedValues.GetAt(i)
					if Mid(existingTaggedValue.Name,1,7) = "backup-" then
						if debug then Repository.WriteOutput "Script", "Delete Tag " & i & " in Attribute [" & theElement.Name & "]"  & " attr.Name [" & attr.Name & "]"  & " TaggedValue.Name [" & existingTaggedValue.Name & "]" & " existingTaggedValue.Value [" & existingTaggedValue.Value & "]" & vbCrLf ,0 end if
						call attr.TaggedValues.DeleteAt(i,0)
						attr.Update()
						theElement.Update()
						theElement.Refresh()
				'		k = k - 1
						i = i + 1
					else
						i = i + 1
					end if
					if debug then Repository.WriteOutput "Script", " wend 2 i= [" & i & "] k= [" & k & "] attr.TaggedValues.Count= [" & attr.TaggedValues.Count & "]" & vbCrLf ,0 end if
				wend
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
'						backupValue = ""
						set existingRoleTaggedValue = connend.TaggedValues.GetAt(i)
'						backupValue = existingTaggedValue.Value
						if debug then Repository.WriteOutput "Script", "Class [" & theElement.Name & "]"  & " role.Name [" & connend.Role & "]"  & " existingRoleTaggedValue.Tag [" & existingRoleTaggedValue.Tag & "]" & " Value [" & existingRoleTaggedValue.Value   & "]" & vbCrLf ,0  end if
						' not copying already backed up tags
						if existingRoleTaggedValue.Value = "" then
							if getConnectorEndTaggedValue(connend,"backup-" + existingRoleTaggedValue.Tag) <> "" then
								existingRoleTaggedValue.Value = getTaggedValue(theElement,"backup-" + existingRoleTaggedValue.Tag)
								existingRoleTaggedValue.Update()
								connend.Update()
								conn.Update()
								conn.Refresh()
								theElement.Update()
								theElement.Refresh()
								Repository.WriteOutput "Script", "Old value restored on Tag  [" & theElement.Name & "]" & " role.Name [" & connend.Role & "]" & " existingRoleTaggedValue.Name [" & existingRoleTaggedValue.Tag & "]" & " Value [" & existingRoleTaggedValue.Value & "]" & vbCrLf ,0
							else
								'make new tag x in class from backup-x, when tag x is not in new stereotype?
							end if
						end if
					next
			'	if value <> "" rename backup-xxx tag to xxx?
					i = 0
					k = connend.TaggedValues.Count - 1
					while i < k and i < connend.TaggedValues.Count
						set existingRoleTaggedValue = connend.TaggedValues.GetAt(i)
						if Mid(existingRoleTaggedValue.Tag,1,7) = "backup-" then
							if debug then Repository.WriteOutput "Script", "Delete Tag " & i & " in Role [" & theElement.Name & "]"  & " role.Name [" & connend.Role & "]" & vbCrLf ,0 end if
							call connend.TaggedValues.DeleteAt(i,0)
							connend.Update()
							conn.Update()
							theElement.Update()
							theElement.Refresh()
							i = i + 1
						else
							i = i + 1
						end if
						if debug then Repository.WriteOutput "Script", " wend i= [" & i & "] k= [" & k & "]" & vbCrLf ,0 end if
					wend

				end if
			next
			
			

			Repository.WriteOutput "Script", Now & " Tagged values restored for Class : " & theElement.Name & ".",0

			case VBcancel

			end select
	

		Else
		  'Other than CodeList selected in the tree
		  MsgBox( "This script requires a class to be selected in the Project Browser." & vbCrLf & _
			"Please select a class in the Project Browser and try again." )
		end If
		'Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires a class to be selected in the Project Browser." & vbCrLf & _
	  "Please select a class in the Project Browser and try again." )
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


restoreTaggedValues
