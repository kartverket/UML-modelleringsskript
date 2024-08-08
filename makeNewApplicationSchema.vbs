option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		makeNewApplicationSchema
' purpose:		updates stereotypes into new 2024 iso standard Rules for Application Schema, changing stereotypes to a new profile without loosing old tagged values
' formål:		endre kjente stereotyper til ny iso19109:2024 standard UML profil
' author:		Kent Jonsrud
' version:		2024-08-08	improved script output
' version:		2024-08-07	specialized class types like datatype, codelist and enumeration need some special care
' version:		2024-07-31	testing prerequisits, that the profile "ISO19109" exists before the package is updated, 
' version:		2024-02-20	testing against the new profile "ISO19109", 
' version:		2024-02-15	testing against the new profile "ISO19109 merged"
' version:		2024-02-09	Switches to current national profile SOSI-UML-profil 5.1::
' version:		2024-02-02	Changed switch over from GI_Class and enumeration to GI_Interface and GI_Enumeration
' version:		2023-10-28	association ends with role names handled
' version:		2023-10-27	switch known stereotypes, from FeatureType to GI_Class etc.
' version:		2023-10-27	classes and attributes handled, codelists with codes become GI_Enumerations
'				TBD: verify that specially association ends keep all their original tags, ?
'				TBD: copy definititon text in Notes into tag definition, ?
'				TBD: use of description <memo>, ?
'				TBD: fix erronious æøå in Notes also at this stage?
'				TBD: still ok to keep all other tags directly on element and not in a profile? (like sequenceNumber,SOSI_Navn,++)
'				TBD: move exiting alias names into (EN) designation tags and move --Definition-- parts into separate (EN) definition tags ?
'				TBD: may be problems with attributes that has some other existing stereotypes ?
'
'				TBD: check that all elements are changed, complete or discard description <memo>, ?enumerations (<memo>)
'				TBD: tagged values on GI_Property and GI_Enum not visible under Properties tab, just stereotype name (however scripts will work)
'				TBD: if not pkg.Update() on all Update(), show old profile name, 

		DIM debug,txt
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
		if false then
			Repository.ClearOutput "Script"
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
			Repository.WriteOutput "Script", "Stereotype : [«" & theElement.element.Stereotype & "»]" & ".",0
			Repository.WriteOutput "Script", "StereotypeEx : [«" & theElement.element.StereotypeEx & "»]" & ".",0
			dim ster as EA.Stereotype
			for i = 0 to Repository.Stereotypes.Count - 1 
				set ster = Repository.Stereotypes.GetAt( i ) 
		'		set xxxx = Repository.Resources.GetAt( i ) 
	'			Repository.WriteOutput "Script", i & " Repository.Stereotypes Name: " & ster.Name ,0
	'			Repository.WriteOutput "Script", i & " Repository.Stereotypes AppliesTo: " & ster.AppliesTo ,0
		'		Repository.WriteOutput "Script", i & " Repository.Stereotypes ObjectType: " & ster.ObjectType ,0
			'	if ster.Type = "Class" then
			'		o = o + 1
			'	end if
			next
		'?		Repository.WriteOutput "Script", "TaggedValuesEX : [«" & theElement.element.TaggedValues.   & "»]" & ".",0
		end if 'debug

		if Repository.IsTechnologyLoaded ("ISO19109") then
			dim box
			box = Msgbox ("Skript makeNewApplicationSchema" & vbCrLf & vbCrLf & "Skriptversjon 2024-08-07" & vbCrLf & _
				"Endrer kjente stereotyper til ny 2024 standard UML profil: [" & theElement.Name & "]."  & vbCrLf & _
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
					box2 = Msgbox ("ALL ELEMENTS WILL NOW BE CHANGED!",1)
					select case box2
						case vbOK

							if Repository.GetTreeSelectedItemType() = otPackage then
								Repository.WriteOutput "Script", Now & " Start processing of package.",0
								makeApplicationSchema(theElement)
							else 
								if Repository.GetTreeSelectedItemType() = otElement and theElement.Type = "Class" or theElement.Type = "DataType" or theElement.Type = "Enumeration" then
									Repository.WriteOutput "Script", Now & " Start processing of class.",0
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
			Repository.WriteOutput "Script", Now & " MDG Technology (UML Profile) ISO19109 not found. End of processing.",0
		end if

	else
		MsgBox( "This script requires a package or a class to be selected in the Project Browser." & vbCrLf & _
	  "Please select this and try again." )
	end if

end sub

sub makeApplicationSchema(pkg)
	dim elements as EA.Collection
	dim i
	Repository.WriteOutput "Script", "              All Stereotypes will be updated for Package : [«" & pkg.element.FQStereotype & "» " & pkg.Name & "].",0


	if LCase(pkg.element.Stereotype) = "applicationschema" or LCase(pkg.element.Stereotype) = "abstractschema" then
		if debug then
			Repository.WriteOutput "Script", "pkg.element.GetStereotypeList() [" &pkg.element.GetStereotypeList() & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "pkg.element.GetStereotypeList() [" &pkg.element.GetStereotypeList() & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "FQName [" & pkg.element.FQName & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "FQ Stereotype [" & pkg.element.FQStereotype & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Has Stereotype SOSI-UML-profil 5.1::ApplicationSchema [" & pkg.element.HasStereotype("SOSI-UML-profil 5.1::ApplicationSchema") & "]" & vbCrLf ,0
			Repository.WriteOutput "Script", "Stereotype : [«" & pkg.element.Stereotype & "»]" & ".",0
			Repository.WriteOutput "Script", "StereotypeEx : [«" & pkg.element.StereotypeEx & "»]" & ".",0
		end if 'debug

		backupPackageTags(pkg)
		pkg.element.StereotypeEx = ""
		if not pkg.element.Update() then Repository.WriteOutput "Script", "Error on removing stereotypes on package element: [" & pkg.Name & "]  " & pkg.element.GetLastError() & ".",0 end if
		pkg.element.Refresh()
		pkg.element.TaggedValues.Refresh()
		if not pkg.Update() then Repository.WriteOutput "Script", "Error on removing stereotypes on package : [" & pkg.Name & "]  " & pkg.GetLastError() & ".",0 end if
	'	pkg.Refresh()
		pkg.element.StereotypeEx = "ISO19109::ApplicationSchema"
		pkg.element.Update()
		if not pkg.element.Update() then Repository.WriteOutput "Script", "Error on Update of package element: [" & pkg.Name & "]  " & pkg.element.GetLastError() & ".",0 end if
		pkg.element.Refresh()
		pkg.element.TaggedValues.Refresh()
		if not pkg.Update() then Repository.WriteOutput "Script", "Error on Update of package : [" & pkg.Name & "]  " & pkg.GetLastError() & ".",0 end if
	'	pkg.Refresh()
	
	'		Repository.WriteOutput "Script", "New FQStereotype : [«" & pkg.element.FQStereotype & "»]" & ".",0
			
		restorePackageTags(pkg)
		pkg.Update()
	'	pkg.Refresh()

		Repository.WriteOutput "Script", " Stereotype changed for Package : [«" & pkg.element.Stereotype & "» " & pkg.Name & "].",0
	end if
	
'	if false then
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
	
	
	dim subP as EA.Package
	
	for each subP in pkg.packages
	    makeApplicationSchema(subP)
	next
'	end if 'false
	
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

	Repository.WriteOutput "Script", " ------------------ Stereotypes will be changed for Class : [«" & theElement.FQStereotype & "» " & theElement.Name & "].",0		
	if debug then
		Repository.WriteOutput "Script", "Old Stereotype [" & theElement.Stereotype & "]" ,0
		Repository.WriteOutput "Script", "FQ  Stereotype [" & theElement.FQStereotype & "]" ,0
		Repository.WriteOutput "Script", "StereotypeList [" & theElement.GetStereotypeList() & "]" ,0
		Repository.WriteOutput "Script", "StereotypeEx   [" & theElement.StereotypeEx & "]" ,0
		Repository.WriteOutput "Script", "MetaType       [" & theElement.MetaType & "]" ,0
	end if
	backupElementTags(theElement)
	
	
				
	if LCase(theElement.Stereotype) = "featuretype" or LCase(theElement.Stereotype) = "gi_interface" then
		if debug then 		Repository.WriteOutput "Script", " featuretype Class : [«" & theElement.Stereotype & "» " & theElement.Name & "].",0 end if
		if debug then call listAllElementTags(theElement, "at start") end if
		theElement.Stereotype = ""
		theElement.StereotypeEx = ""
		theElement.Update()
		theElement.TaggedValues.Refresh()
		deleteAllOldElementTags(theElement)
		theElement.Refresh()
		if debug then call listAllElementTags(theElement, "all tags after old stereotype is removed and before new stereotype") end if
		theElement.StereotypeEx = "ISO19109::FeatureType"
		theElement.Update()
		theElement.TaggedValues.Refresh()
		theElement.Refresh()
		if debug then call listAllElementTags(theElement, "tags after some of them are moved to new stereotype") end if
		if debug then call listAllElementTags(theElement, "tags after restore") end if
		
		for each attr in theElement.Attributes
			makeAttribute(attr)
			theElement.Update()
			theElement.Refresh()
			theElement.TaggedValues.Refresh()
			'?
		next
		if debug then call listAllElementTags(theElement, "tags after attributes updated") end if
		theElement.TaggedValues.Refresh()
		theElement.Update()
		theElement.Refresh()
		restoreElementTags(theElement)
		if debug then call listAllElementTags(theElement, "tags after restore") end if
		theElement.TaggedValues.Refresh()
		theElement.Update()
		theElement.Refresh()
		if debug then call listAllElementTags(theElement, "tags after last Refresh") end if
	end if
	if LCase(theElement.Stereotype) = "datatype" or LCase(theElement.Stereotype) = "gi_datatype" then
		if debug then 		Repository.WriteOutput "Script", " datatype Class : [«" & theElement.Stereotype & "» " & theElement.Name & "].",0 end if
		if debug then 		Repository.WriteOutput "Script", " datatype1 Type : [«" & theElement.ClassifierType & "» " & theElement.Type & "].",0 end if
		if debug then call listAllElementTags(theElement, "at start") end if
		theElement.Stereotype = ""
		theElement.StereotypeEx = ""
		theElement.Type = "DataType"
		theElement.Update()
		theElement.TaggedValues.Refresh()
		deleteAllOldElementTags(theElement)
		theElement.Refresh()
		if debug then call listAllElementTags(theElement, "all tags after old stereotype is removed and before new stereotype") end if
		theElement.StereotypeEx = "ISO19109::GI_DataType"
		theElement.Update()
		theElement.TaggedValues.Refresh()
		theElement.Refresh()
		if debug then call listAllElementTags(theElement, "tags after new stereotype") end if
		if debug then 		Repository.WriteOutput "Script", " datatype2 Class : [«" & theElement.Stereotype & "» " & theElement.Name & "].",0 end if

		for each attr in theElement.Attributes
			makeAttribute(attr)
			theElement.Update()
			theElement.Refresh()
	'		theElement.TaggedValues.Refresh()
		next
		if debug then call listAllElementTags(theElement, "tags after attributes updated") end if
		theElement.TaggedValues.Refresh()
		theElement.Update()
		theElement.Refresh()
		restoreElementTags(theElement)
		if debug then call listAllElementTags(theElement, "tags after restore") end if
		theElement.TaggedValues.Refresh()
		theElement.Update()
		theElement.Refresh()
		if debug then call listAllElementTags(theElement, "tags after last Refresh") end if
	end if
	if LCase(theElement.Stereotype) = "codelist" or LCase(theElement.Stereotype) = "gi_codeset" or LCase(theElement.Stereotype) = "enumeration" or LCase(theElement.Stereotype) = "gi_enumeration" or theElement.Type = "Enumeration"then
		if debug then 		Repository.WriteOutput "Script", " codelist Class : [«" & theElement.Stereotype & "» " & theElement.Name & "] codelist Type : [" & theElement.Type & "].",0 end if
		theElement.Stereotype = ""
		theElement.StereotypeEx = ""
		if theElement.Attributes.Count() > 0 then
			theElement.Type = "Enumeration"
			theElement.Update()
			theElement.TaggedValues.Refresh()
			deleteAllOldElementTags(theElement)
			theElement.Refresh()
			theElement.StereotypeEx = "ISO19109::GI_Enumeration"
			theElement.Update()
			theElement.TaggedValues.Refresh()
			theElement.Refresh()
		else
			theElement.Type = "DataType"
			theElement.Update()
			theElement.TaggedValues.Refresh()
			deleteAllOldElementTags(theElement)
			theElement.Refresh()
			theElement.StereotypeEx = "ISO19109::GI_CodeSet"
			theElement.Update()
			theElement.TaggedValues.Refresh()
			theElement.Refresh()
		end if

		if theElement.Attributes.Count() > 0 then
			Repository.WriteOutput "Script", "Info: Element with codes not empty! ["  & theElement.Name & " has "& theElement.Attributes.Count() & " codes]" & vbCrLf ,0
			for each attr in theElement.Attributes
				makeAttribute(attr)
				theElement.Update()
				theElement.Refresh()
				theElement.TaggedValues.Refresh()
			next
		end if
		if debug then call listAllElementTags(theElement, "tags after attributes updated") end if
		theElement.TaggedValues.Refresh()
		theElement.Update()
		theElement.Refresh()
		restoreElementTags(theElement)
		if debug then call listAllElementTags(theElement, "tags after restore") end if
		theElement.TaggedValues.Refresh()
		theElement.Update()
		theElement.Refresh()
		if debug then call listAllElementTags(theElement, "tags after last Refresh") end if
	end if
'	if LCase(theElement.Stereotype) = "enumeration" or LCase(theElement.Stereotype) = "gi_enumeration" or theElement.Type = "Enumeration" then
'		if debug then 		Repository.WriteOutput "Script", " TBD enumeration Class : [«" & theElement.Stereotype & "» " & theElement.Name & "].",0 end if
'		'TBD
'		'for each attr in theElement.Attributes
'			'makeCode(attr)
'		'next
'	end if

'	Repository.WriteOutput "Script", "New Stereotype [" & theElement.Stereotype & "]" & vbCrLf ,0			
	Repository.WriteOutput "Script", " Stereotype changed for Class : [«" & theElement.Stereotype & "» " & theElement.Name & "].",0




				
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
				
	backupAttributeTags(attr)
	
'	if LCase(attr.Stereotype) = "" and LCase(attr.Type) <> "" then
	if debug then
		Repository.WriteOutput "Script", " Stereotype to be changed for Attribute : [" & attr.Name & " : " & attr.Type & "].",0
	end if
	if LCase(attr.Stereotype) = "" or LCase(attr.Stereotype) = "egenskap" or LCase(attr.Stereotype) = "gi_property" then
		attr.StereotypeEx = ""
		attr.Update()
		attr.TaggedValues.Refresh()
		attr.StereotypeEx = "ISO19109::GI_Property"
		attr.Update()
		attr.TaggedValues.Refresh()
	end if
	if LCase(attr.Stereotype) = "enum" or LCase(attr.Stereotype) = "kode" or LCase(attr.Type) = "<undefined>" or LCase(attr.Type) = "" then
		attr.StereotypeEx = ""
		attr.Update()
		attr.TaggedValues.Refresh()
		attr.StereotypeEx = "ISO19109::GI_EnumerationLiteral"
		attr.Update()
		attr.TaggedValues.Refresh()
	end if

	if LCase(attr.Stereotype) = "enum" or LCase(attr.Stereotype) = "gi_enumerationliteral" or LCase(attr.Type) = "<undefined>" or LCase(attr.Type) = "" then
		Repository.WriteOutput "Script", " Stereotype changed for Code : [«" & attr.Stereotype & "» " & attr.Name & "].",0
	else
		Repository.WriteOutput "Script", " Stereotype changed for Attribute : [«" & attr.Stereotype & "» " & attr.Name & "].",0
	end if
	restoreAttributeTags(attr)

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
