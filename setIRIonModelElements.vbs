option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		setIRIonModelElements
' purpose:		sets a targetNamespace-based http-uri into empty tagged value IRI of iso19103:2024 standard UML profil
' formål:		setter targetNamespace-basert http-uri inn i tomme tagged value IRI fra ny iso19103:2024 standard UML profil
' author:		Kent Jonsrud
' version:		2024-12-23	class handled
' version:		2024-12-19	package handled
'				TBD: roles and attributes TB handled, codelists with codes TB ...
'				TBD: better logic for setting namespace
'				TBD: create tag IRI if missing
'				TBD: cleaning up

		DIM debug, txt, targetnamespace, pkgncname
		debug = false
		
	Dim tnplist, tvplist
	Dim tnelist, tvelist
	Dim tnalist, tvalist
	Dim tnrlist, tvrlist

sub makeNewApplicationSchema()
	Repository.EnsureOutputVisible "Script"

	Dim i,o
	Dim theElement as EA.Element
	dim conn as EA.Connector
	dim connend as EA.ConnectorEnd
	Dim newETag as EA.TaggedValue
	Dim newATag as EA.AttributeTag
	Dim newCTag as EA.RoleTag
	
	Set tnplist = CreateObject("System.Collections.ArrayList")
	Set tvplist = CreateObject("System.Collections.ArrayList")
	Set tnelist = CreateObject("System.Collections.ArrayList")
	Set tvelist = CreateObject("System.Collections.ArrayList")
	Set tnalist = CreateObject("System.Collections.ArrayList")
	Set tvalist = CreateObject("System.Collections.ArrayList")
	Set tnrlist = CreateObject("System.Collections.ArrayList")
	Set tvrlist = CreateObject("System.Collections.ArrayList")

	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		o = 0
		if debug then
	'		Repository.ClearOutput "Script"
			Repository.WriteOutput "Script", "Repository.ConnectionString [" &Repository.ConnectionString & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Repository.RepositoryType() [" &Repository.RepositoryType() & "]" & vbCrLf ,0
	'		Repository.WriteOutput "Script", "Repository.EAEditionEx [" &Repository.EAEditionEx & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Repository.LibraryVersion [" &Repository.LibraryVersion & "]" & vbCrLf ,0
	'		Repository.WriteOutput "Script", "Repository.LastUpdate [" &Repository.LastUpdate & "]" & vbCrLf ,0
	'		Repository.WriteOutput "Script", "Repository.GetCounts() [" &Repository.GetCounts() & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Repository.IsTechnologyLoaded (GML) [" &Repository.IsTechnologyLoaded ("GML") & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Repository.IsTechnologyLoaded (ISO19109) [" &Repository.IsTechnologyLoaded ("ISO19109") & "]" & vbCrLf ,0
		 
			Repository.WriteOutput "Script", "theElement.element.GetStereotypeList() [" &theElement.element.GetStereotypeList() & "]" & vbCrLf ,0
	'		Repository.WriteOutput "Script", "theElement.element.GetStereotypeList() [" &theElement.element.   .GetStereotypeList() & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "FQName [" & theElement.element.FQName & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "FQ Stereotype [" & theElement.element.FQStereotype & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Has Stereotype SOSI-UML-profil 5.1::ApplicationSchema [" & theElement.element.HasStereotype("SOSI-UML-profil 5.1::ApplicationSchema") & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Has Stereotype ISO19109::ApplicationSchema [" & theElement.element.HasStereotype("ISO19109::ApplicationSchema") & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Has Stereotype ISO19103::AbstractSchema [" & theElement.element.HasStereotype("ISO19103::AbstractSchema") & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Stereotype : [«" & theElement.element.Stereotype & "»]" & ".",0
			Repository.WriteOutput "Script", "StereotypeEx : [«" & theElement.element.StereotypeEx & "»]" & ".",0
		'?		Repository.WriteOutput "Script", "TaggedValuesEX : [«" & theElement.element.TaggedValues.   & "»]" & ".",0
		end if 'debug

		if Repository.IsTechnologyLoaded ("ISO19103") then
			dim box
			box = Msgbox ("Skript setIRIonModelElements" & vbCrLf & vbCrLf & "Skriptversjon 2024-12-23" & vbCrLf & _
				"Setter http-uri-er i tagged value IRI fra ny 2024 standard UML profil: [" & theElement.Name & "]."  & vbCrLf & _
				"NOTE! This script will make CHANGES to ALL elements in your model!",1)
			select case box
				case vbOK
				
					Repository.ClearOutput "Script"
					Repository.CreateOutputTab "Error"
					Repository.ClearOutput "Error"			
					
			'		if debug then
			'			Repository.WriteOutput "Script", "Repository.GetTreeSelectedItemType() [" &Repository.GetTreeSelectedItemType() & "]" & " theElement.Type [" &theElement.Type & "]" & vbCrLf ,0
			'		end if

					dim box2
					box2 = Msgbox ("ALL ELEMENTS IN THE MODEL MAY NOW BE CHANGED !",1)
					select case box2
						case vbOK

							if Repository.GetTreeSelectedItemType() = otPackage then
								Repository.WriteOutput "Script", Now & " Start processing of package. Skriptversjon 2024-12-23",0
								makeApplicationSchema(theElement)
							else 
								if Repository.GetTreeSelectedItemType() = otElement and theElement.Type = "Class" or theElement.Type = "DataType" or theElement.Type = "Enumeration" then
									Repository.WriteOutput "Script", Now & " Start processing of class. Skriptversjon 2024-12-23",0
									makeClass(theElement)
								else
									MsgBox( "This script requires a package or a class to be selected in the Project Browser." & vbCrLf & _
										"Please select this and try again." )			
								end if
							end if
							Repository.WriteOutput "Script", Now & " End of processing.",0
						case VBcancel
					end select

				case VBcancel
			end select
		else
			Repository.WriteOutput "Script", Now & " MDG Technology (UML Profile) ISO19103 not found. End of processing.",0
		end if

	else
		MsgBox( "This script requires a package or a class to be selected in the Project Browser." & vbCrLf & _
	  "Please select this and try again." )
	end if

end sub

sub makeApplicationSchema(pkg)
	dim elements as EA.Collection
	dim i
	Repository.WriteOutput "Script", "  ----------IRI may be updated for Package : [«" & pkg.element.FQStereotype & "» " & pkg.Name & "].",0


'	if LCase(pkg.element.Stereotype) = "applicationschema" or LCase(pkg.element.Stereotype) = "abstractschema" then
		if debug then
			Repository.WriteOutput "Script", "pkg.element.GetStereotypeList() [" &pkg.element.GetStereotypeList() & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "pkg.element.GetStereotypeList() [" &pkg.element.GetStereotypeList() & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "FQName [" & pkg.element.FQName & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "FQ Stereotype [" & pkg.element.FQStereotype & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Has Stereotype SOSI-UML-profil 5.1::ApplicationSchema [" & pkg.element.HasStereotype("SOSI-UML-profil 5.1::ApplicationSchema") & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Stereotype : [«" & pkg.element.Stereotype & "»]" & ".",0
			Repository.WriteOutput "Script", "StereotypeEx : [«" & pkg.element.StereotypeEx & "»]" & ".",0
		end if 'debug

		if LCase(pkg.element.Stereotype) = "applicationschema" or LCase(pkg.element.Stereotype) = "abstractschema" then
	'		if targetnamespace = "" then
				targetnamespace = getPackageTaggedValue(pkg,"targetNamespace")
	'		end if
			if targetnamespace = "" then targetnamespace = "http://my.site.domain/my/landing/page" end if
	'		if mid(targetnamespace,len(targetnamespace)-1, len(targetnamespace)) = "/" then targetnamespace = mid(targetnamespace,1, len(targetnamespace)-1) end if
			pkgncname = getNCNameX(pkg.Name)
			if getPackageTaggedValue(pkg,"IRI") = "" then
				dim existingTaggedValue
				for i = 0 to pkg.element.TaggedValues.Count - 1
					set existingTaggedValue = pkg.element.TaggedValues.GetAt(i)
					if existingTaggedValue.Name = "IRI" then
				'		existingTaggedValue.Value = getPackageTaggedValue(pkg,"targetNamespace") & "/" & getNCNameX(pkg.Name)
						existingTaggedValue.Value = targetnamespace & "/" & pkgncname
						existingTaggedValue.Update()
						pkg.element.TaggedValues.Refresh()
						pkg.element.Refresh()
'						Repository.WriteOutput "Script", " IRI tag set for Package : [«" & pkg.element.Stereotype & "» " & pkg.element.Name & "] - on tag number [" & i & "].",0
						Repository.WriteOutput "Script", " IRI tag set for Package : [«" & pkg.element.Stereotype & "» " & pkg.element.FQName & "] - on tag number [" & i & "].",0
					end if
				next
				if getPackageTaggedValue(pkg,"IRI") = "" then
					' add tag IRI
				end if
			else
				Repository.WriteOutput "Script", " IRI tag already set to: [" & getPackageTaggedValue(pkg,"IRI") & "] for Package : [«" & pkg.element.Stereotype & "» " & pkg.element.FQName & "].",0
			end if
		end if
'	end if		

	
'	if false then ' temporary
		set elements = pkg.Elements
		for i = 0 to elements.Count - 1 
			dim element as EA.Element 
			set element = elements.GetAt( i ) 	
			if debug then Repository.WriteOutput "Script", "Element Name [" & element.Name & "]"  & " Element Type [" & element.Type & "]" ,0 end if
			if element.Type = "Class" or element.Type = "DataType" or element.Type = "Enumeration" or element.Type = "Type" then
				makeClass(element)
				element.Refresh()
			end if
		next
'	end if 'false temporary
		
	dim subP as EA.Package
		
	for each subP in pkg.packages
		makeApplicationSchema(subP)
	next

	
end sub


sub makeClass(theElement)

'	Dim theElement as EA.Element
	dim attr as EA.Attribute
	dim conn as EA.Connector
	dim connend as EA.ConnectorEnd
	Dim newETag as EA.TaggedValue
	Dim newATag as EA.AttributeTag
	Dim newCTag as EA.RoleTag
	dim i, c

	Repository.WriteOutput "Script", " ------------------ Tagged value IRI  will be added for Class : [«" & theElement.FQStereotype & "» " & theElement.Name & "].",0		
	if debug then
		Repository.WriteOutput "Script", "Stereotype [" & theElement.Stereotype & "]" ,0
		Repository.WriteOutput "Script", "FQ  Stereotype [" & theElement.FQStereotype & "]" ,0
		Repository.WriteOutput "Script", "StereotypeList [" & theElement.GetStereotypeList() & "]" ,0
		Repository.WriteOutput "Script", "StereotypeEx   [" & theElement.StereotypeEx & "]" ,0
		Repository.WriteOutput "Script", "MetaType       [" & theElement.MetaType & "]" ,0
	end if
	
	
	if getTaggedValue(theElement,"IRI") = "" then
		dim existingTaggedValue
		for i = 0 to theElement.TaggedValues.Count - 1
			set existingTaggedValue = theElement.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = "IRI" then
		'		existingTaggedValue.Value = getPackageTaggedValue(pkg,"targetNamespace") & "/" & getNCNameX(pkg.Name)
				existingTaggedValue.Value = targetnamespace & "/" & pkgncname & "/" & getNCNameX(theElement.Name)
				existingTaggedValue.Update()
				theElement.TaggedValues.Refresh()
				theElement.Update()
				theElement.Refresh()
'				Repository.WriteOutput "Script", " IRI tag set for Class : [«" & theElement.Stereotype & "» " & theElement.Name & "] - on tag number [" & i & "].",0
				Repository.WriteOutput "Script", " IRI tag set for Class : [«" & theElement.Stereotype & "» " & theElement.FQName & "] - on tag number [" & i & "].",0
			end if
		next
		if getTaggedValue(theElement,"IRI") = "" then
			' add tag IRI
			Repository.WriteOutput "Script", " tagged value IRI added to Class : [«" & theElement.Stereotype & "» " & theElement.Name & "].",0
		end if
	else
		Repository.WriteOutput "Script", " IRI tag already set to: [" & getTaggedValue(theElement,"IRI") & "] for Class : [«" &theElement.Stereotype & "» " & theElement.FQName & "].",0
	end if
	
	
	
	
	

	if false then ' temporary
		
		for each attr in theElement.Attributes
			makeAttribute(attr)
			theElement.Update()
			theElement.Refresh()
			theElement.TaggedValues.Refresh()
			'?
		next
	end if 'false temporary

		



	if false then ' temporary

				
	for each conn in theElement.Connectors
		'''Repository.WriteOutput "Script", "Connector found [" & theElement.Name & "]"  & " Connector.Name [" & conn.Name & "]" & " Connector.Type [" & conn.Type & "]" & vbCrLf ,0
		if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then
		else
			if debug then Repository.WriteOutput "Script", "Class [" & theElement.Name & "]"  & " has Connector Type [" & conn.Type & "]" & vbCrLf ,0 end if
		'	find roles referring to the other class
			if conn.ClientID = theElement.ElementID then
				if debug then Repository.WriteOutput "Script", "Connector to supplier [" & theElement.Name & "]"  & " conn.SupplierEnd.Role [" & conn.SupplierEnd.Role & "] to Target Class [" & Repository.GetElementByID(conn.SupplierID).Name & "]" & vbCrLf ,0 end if
				set connend = conn.SupplierEnd
			else
				if debug then Repository.WriteOutput "Script", "Connector to client [" & theElement.Name & "]"  & " conn.ClientEnd.Role [" & conn.ClientEnd.Role & "]" & " to SourceClass [" & Repository.GetElementByID(conn.ClientID).Name & "]" & vbCrLf ,0 end if
				set connend = conn.ClientEnd
			end if

			if debug then Repository.WriteOutput "Script", "Class [" & theElement.Name & "]"  & " role.Name [" & connend.Role & "]"  & " role.Stereotype [" & connend.Stereotype & "]" & " role.StereotypeEx [" & connend.StereotypeEx & "]" & vbCrLf ,0  end if
			if connend.Role <> "" then
				backupRoleTags(connend)
				'TBD Delete all tags and trust the restore sub ?
				c = connend.TaggedValues.Count
				for i = c - 1 to 0 step -1
					connend.TaggedValues.Delete(i)
				next
				connend.StereotypeEx = ""
				connend.Update()
				connend.TaggedValues.Refresh()
				conn.Update()
				theElement.Update()
				theElement.Refresh()
				connend.StereotypeEx = "ISO19109::GI_Property"
				connend.TaggedValues.Refresh()
				connend.Update()	'This one gives several SQL error messages if the connection end has old (even empty) tags that should belong to the new stereotype
				conn.Update()		'This one seems like it too
				theElement.Update()
				theElement.Refresh()
				restoreRoleTags(connend)
			end if
		end if
	next
'	restoreElementTags(theElement)
'	theElement.TaggedValues.Refresh()
'	theElement.Update()
'	theElement.Refresh()
			
	end if 'false temporary
			


end sub

sub makeAttribute(attr)

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
				
'	if LCase(attr.Stereotype) = "" and LCase(attr.Type) <> "" then
	if debug then
		Repository.WriteOutput "Script", " Tagged value IRI to be changed for Attribute : [" & attr.Name & " : " & attr.Type & "].",0
	end if
	
	
	if getTaggedValue(attr,"IRI") = "" then
		dim existingTaggedValue, i
		for i = 0 to attr.TaggedValues.Count - 1
			set existingTaggedValue = attr.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = "IRI" then
		'		existingTaggedValue.Value = getPackageTaggedValue(pkg,"targetNamespace") & "/" & getNCNameX(attr.Name)
				existingTaggedValue.Value = targetnamespace & "/" & pkgncname & "/" & getNCNameX(attr.Name)
				existingTaggedValue.Update()
				attr.TaggedValues.Refresh()
				attr.Update()
'				Repository.WriteOutput "Script", " IRI tag set for Attribute : [«" & attr.Stereotype & "» " & attr.Name & "] - on tag number [" & i & "].",0
				Repository.WriteOutput "Script", " IRI tag set for Attribute : [«" & attr.Stereotype & "» " & attr.Name & "] - on tag number [" & i & "].",0
			end if
		next
		if getTaggedValue(attr,"IRI") = "" then
			' add tag IRI
			Repository.WriteOutput "Script", " tagged value IRI added to Attribute : [«" & attr.Stereotype & "» " & attr.Name & "].",0
		end if
	else
		Repository.WriteOutput "Script", " IRI tag already set to Attribute : [" & getTaggedValue(attr,"IRI") & "] for Class : [«" & attr.Stereotype & "» " & attr.Name & "].",0
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

function getTaggedValueCount(element,taggedValueName)
		dim i, existingTaggedValue
		getTaggedValueCount = 0
		for i = 0 to element.TaggedValues.Count - 1
			set existingTaggedValue = element.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = taggedValueName then
				getTaggedValueCount = getTaggedValueCount + 1
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

function getPackageTaggedValueCount(package,taggedValueName)
		dim i, existingTaggedValue
		getPackageTaggedValueCount = 0
		for i = 0 to package.element.TaggedValues.Count - 1
			set existingTaggedValue = package.element.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = taggedValueName then
				getPackageTaggedValueCount = getPackageTaggedValueCount + 1
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

sub backupPackageTags(pkg)
	dim existing, i
	tnplist.Clear
	tvplist.Clear
	for i = 0 to pkg.Element.TaggedValues.Count - 1
		set existing = pkg.Element.TaggedValues.GetAt(i)
		if existing.Value <> "" then
			if debug then Repository.WriteOutput "Script", " Backup Tagged value " & i & " for package : [" & pkg.Name & " - tag name: " & existing.Name & " tag value: " & existing.Value & "].",0 end if
			tnplist.Add existing.Name
			tvplist.Add existing.Value
		else
			if debug then Repository.WriteOutput "Script", " Empty Tagged value " & i & " for package : [" & pkg.Name & " - tag name: " & existing.Name & " tag value: " & existing.Value & "].",0 end if
		end if
	next
end sub

sub listAllElementTags(element, stage)
	dim existing, i
	Repository.WriteOutput "Script", " Stage : [" & stage & "].",0 
	for i = 0 to element.TaggedValues.Count - 1
		set existing = element.TaggedValues.GetAt(i)
		Repository.WriteOutput "Script", " --- Tagged value " & i & " for element : [" & element.Name & "] - tag name: [" & existing.Name & "] tag FQname: [" & existing.FQName & "] tag value: [" & existing.Value & "].",0 
	next
end sub

sub deleteAllOldElementTags(element)
	dim existing, i
	if element.TaggedValues.Count > 0 then
		if debug then Repository.WriteOutput "Script", " \\\ Deleting Count " & element.TaggedValues.Count & " for element : [" & element.Name & "] ",0 
		for i = 0 to element.TaggedValues.Count - 1
			if debug then 
				set existing = element.TaggedValues.GetAt(i)
				Repository.WriteOutput "Script", " \\\ Deleting Tagged value " & i & " for element : [" & element.Name & "] - tag name: [" & existing.Name & "] tag FQname: [" & existing.FQName & "] tag value: [" & existing.Value & "].",0 
			end if
			element.TaggedValues.Delete(i)
		next
		element.Update()
		element.TaggedValues.Refresh()
		element.Refresh()
	end if
end sub

sub backupElementTags(element)
	dim existing, i
	tnelist.Clear
	tvelist.Clear
	for i = 0 to element.TaggedValues.Count - 1
		set existing = element.TaggedValues.GetAt(i)
		if existing.Value <> "" then
			if debug then Repository.WriteOutput "Script", " Backup Tagged value " & i & " for element : [" & element.Name & "] - tag name: [" & existing.Name & "] tag value: [" & existing.Value & "].",0 end if
			tnelist.Add existing.Name
			tvelist.Add existing.Value
		end if
	next
end sub

sub backupAttributeTags(attr)
	dim existing, i
	tnalist.Clear
	tvalist.Clear
	for i = 0 to attr.TaggedValues.Count - 1
		set existing = attr.TaggedValues.GetAt(i)
		if existing.Value <> "" then
			if debug then Repository.WriteOutput "Script", " Backup Tagged value " & i & " for attribute : [" & attr.Name & " - tag name: " & existing.Name & " tag value: " & existing.Value & "].",0 end if
			tnalist.Add existing.Name
			tvalist.Add existing.Value
		end if
	next
end sub

sub backupRoleTags(connend)
	dim existing, i
	tnrlist.Clear
	tvrlist.Clear
	for i = 0 to connend.TaggedValues.Count - 1
		set existing = connend.TaggedValues.GetAt(i)
		if existing.Value <> "" then
			if debug then Repository.WriteOutput "Script", " Backup Tagged value " & i & " for role : [" & connend.Role & " - tag name: " & existing.Tag & " tag value: " & existing.Value & "].",0 end if
			tnrlist.Add existing.Tag
			tvrlist.Add existing.Value
		end if
	next
end sub

sub restorePackageTags(pkg)
	dim existingTag 'as EA.TaggedValue
	Dim i, j, hit
	Dim newTag as EA.TaggedValue
	if not pkg.element.Update() then Repository.WriteOutput "Script", "Error before restoring stereotypes on package element: [" & pkg.Name & "]  " & pkg.element.GetLastError() & ".",0 end if
	for i = 0 to tnplist.Count - 1
		if debug then Repository.WriteOutput "Script", " Restore Tagged value " & i & " for package : [" & pkg.Name & " - tag name: " & tnplist.Item(i) & " - tag value: " & tvplist.Item(i) & "].",0 end if
		hit = 0
	if debug then Repository.WriteOutput "Script", " Tagged Values tnlist Index" & i & " Count for package : [" & pkg.Name & " - " & pkg.element.TaggedValues.Count  & "].",0 end if
		for j = 0 to pkg.element.TaggedValues.Count - 1
	if debug then Repository.WriteOutput "Script", " Tagged Values Index " & j & " for package : [" & pkg.Name & "].",0 end if
'	if debug then Repository.WriteOutput "Script", " Tagged Values Name " & j & " for package : [" & pkg.Name & "].",0 end if
		if debug then Repository.WriteOutput "Script", "Before restoring stereotypes on package element: [" & pkg.Name & "]  " & pkg.element.GetLastError() & ".",0 end if
		if debug then Repository.WriteOutput "Script", "  pkg.element.TaggedValues.GetLastError(): " & pkg.element.TaggedValues.GetLastError() & ".",0 end if
			set existingTag = pkg.element.TaggedValues.GetAt(j)
			if existingTag.Name = tnplist.Item(i) then
				hit = hit + 1
				if existingTag.Value = "" then
					'tag exists but empty
					existingTag.Value = tvplist.Item(i)
					existingTag.Update()
				'	existingTag.Refresh()
					pkg.element.Update()
					pkg.element.Refresh()
				else
					if existingTag.Value <> tvplist.Item(i) then
						'tag exist but different value stored
						Repository.WriteOutput "Script", " Multi valued Tagged value " & i & " for package : [" & pkg.Name & " - tag name: " & tnplist.Item(i) & " tag value: " & tvplist.Item(i) & " existing tag value: " & existingTag.Value & "].",0
						set newTag = pkg.Element.TaggedValues.AddNew(tnplist.Item(i),"Tag")
						newTag.Value = tvplist.Item(i)
						newTag.Update()
						pkg.element.Update()
						pkg.element.Refresh()
					else
						'same value found
					end if
				end if
			end if
			
		next

			'
		if hit = 0 then
			'tag missing in new stereotype, insert the old tag
			if debug then Repository.WriteOutput "Script", " Keep old Tagged value " & i & " for package : [" & pkg.Name & " - tag name: " & tnplist.Item(i) & " tag value: " & tvplist.Item(i) & "].",0 end if
			set newTag = pkg.Element.TaggedValues.AddNew(tnplist.Item(i),"Tag")
			newTag.Value = tvplist.Item(i)
			newTag.Update()
			pkg.element.Update()
			pkg.element.Refresh()
		end if
		
	next
	
end sub

sub restoreElementTags(element)
	dim existingTag 'as EA.TaggedValue
	Dim i, j, hit
	Dim newTag as EA.TaggedValue
	for i = 0 to tnelist.Count - 1
		if debug then Repository.WriteOutput "Script", " Restore Tagged value " & i & " for element : [" & element.Name & " - tag name: " & tnelist.Item(i) & " tag value: " & tvelist.Item(i) & "].",0 end if
		hit = 0
	
		for j = 0 to element.TaggedValues.Count - 1
			set existingTag = element.TaggedValues.GetAt(j)
			if existingTag.Name = tnelist.Item(i) then
				if debug then Repository.WriteOutput "Script", " Tag found in " & j & " for element : [" & element.Name & " - tag name: " & existingTag.Name & " tag value: " & existingTag.Value & "].",0 end if
				hit = hit + 1
				if existingTag.Value = "" then
					'tag exists but empty
					existingTag.Value = tvelist.Item(i)
					existingTag.Update()
				'	existingTag.Refresh()
					element.Update()
					element.Refresh()
				else
					if existingTag.Value <> tvelist.Item(i) then
						'tag exist but different value stored
						Repository.WriteOutput "Script", " Multi valued Tagged value " & i & " for element: [" & element.Name & " - tag name: " & tnelist.Item(i) & " tag value: " & tvelist.Item(i) & " existing tag value: " & existingTag.Value & "].",0
						set newTag = element.TaggedValues.AddNew(tnelist.Item(i),"Tag")
						newTag.Value = tvelist.Item(i)
						newTag.Update()
						element.Update()
						element.Refresh()
					else
						'same value found
					end if
				end if
			end if
			
		next

			'
		if hit = 0 then
			'tag missing in new stereotype, insert the old tag
			if debug then Repository.WriteOutput "Script", " Keep old Tagged value " & i & " for element : [" & element.Name & " - tag name: " & tnelist.Item(i) & " tag value: " & tvelist.Item(i) & "].",0 end if
			set newTag = element.TaggedValues.AddNew(tnelist.Item(i),"Tag")
			newTag.Value = tvelist.Item(i)
			newTag.Update()
			element.Update()
			element.Refresh()
		end if
		
	next
	
end sub

sub restoreAttributeTags(attr)
	dim existingTag 'as EA.TaggedValue
	Dim i, j, hit
	'Dim newTag as EA.TaggedValue
	Dim newTag as EA.AttributeTag
	for i = 0 to tnalist.Count - 1
		if debug then Repository.WriteOutput "Script", " Restore Tagged value " & i & " for attribute : [" & attr.Name & " - tag name: " & tnalist.Item(i) & " tag value: " & tvalist.Item(i) & "].",0 end if
		hit = 0
	
		for j = 0 to attr.TaggedValues.Count - 1
			set existingTag = attr.TaggedValues.GetAt(j)
			if existingTag.Name = tnalist.Item(i) then
				hit = hit + 1
				if existingTag.Value = "" then
					'tag exists but empty
					existingTag.Value = tvalist.Item(i)
					existingTag.Update()
				'	existingTag.Refresh()
					attr.Update()
				'	attr.Refresh()
				else
					if existingTag.Value <> tvalist.Item(i) then
						'tag exist but different value stored
						Repository.WriteOutput "Script", " Multi valued Tagged value " & i & " for attribute: [" & attr.Name & " - tag name: " & tnalist.Item(i) & " tag value: " & tvalist.Item(i) & " existing tag value: " & existingTag.Value & "].",0
						set newTag = attr.TaggedValues.AddNew(tnelist.Item(i),"Tag")
						newTag.Value = tvalist.Item(i)
						newTag.Update()
						attr.Update()
					'	attr.Refresh()
					else
						'same value found
					end if
				end if
			end if
			
		next

			'
		if hit = 0 then
			'tag missing in new stereotype, insert the old tag
			if debug then Repository.WriteOutput "Script", " Keep old Tagged value " & i & " for attribute : [" & attr.Name & " - tag name: " & tnalist.Item(i) & " tag value: " & tvalist.Item(i) & "].",0 end if
			set newTag = attr.TaggedValues.AddNew(tnalist.Item(i),"Tag")
			newTag.Value = tvalist.Item(i)
			newTag.Update()
			attr.Update()
		'	attr.Refresh()
		end if
		
	next
	
end sub

sub restoreRoleTags(connend)
	dim existing 'as EA.RoleTag
	Dim i, j, hit
	Dim newTag as EA.RoleTag
	for i = 0 to tnrlist.Count - 1
		if debug then Repository.WriteOutput "Script", " Restore Tagged value " & i & " for role: [" & connend.Role & " - tag name: " & tnrlist.Item(i) & " tag value: " & tvrlist.Item(i) & "].",0 end if
		hit = 0
	
		for j = 0 to connend.TaggedValues.Count - 1
			set existing = connend.TaggedValues.GetAt(j)
			if existing.Tag = tnrlist.Item(i) then
				hit = hit + 1
				if existing.Value = "" then
					'tag exists but empty
					existing.Value = tvrlist.Item(i)
					existing.Update()
				'	existing.Refresh()
					connend.Update()
					connend.Refresh()
				else
					if existing.Value <> tvrlist.Item(i) then
						'tag exist but different value stored
						Repository.WriteOutput "Script", " Multi valued Tagged value " & i & " for role : [" & connend.Role & " - tag name: " & tnrlist.Item(i) & " tag value: " & tvrlist.Item(i) & " existing tag value: " & existing.Value & "].",0
						set newTag = connend.TaggedValues.AddNew(tnrlist.Item(i),"Tag")
						newTag.Value = tvrlist.Item(i)
						newTag.Update()
						connend.Update()
				'		connend.Refresh()
					else
						'same value found
					end if
				end if
			end if
			
		next

			'
		if hit = 0 then
			'tag missing in new stereotype, insert the old tag
			if debug then Repository.WriteOutput "Script", " Keep old Tagged value " & i & " for role: [" & connend.Role & " - tag name: " & tnrlist.Item(i) & " tag value: " & tvrlist.Item(i) & "].",0 end if
			set newTag = connend.TaggedValues.AddNew(tnrlist.Item(i),"Tag")
			newTag.Value = tvrlist.Item(i)
			newTag.Update()
			connend.Update()
		'	connend.Refresh()
		end if
		
	next
	
end sub


makeNewApplicationSchema
