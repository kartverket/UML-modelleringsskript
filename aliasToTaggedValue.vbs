option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: aliasToTaggedValue
' Author: Sara Henriksen, Kent Jonsrud
' Purpose: For every package/attribute/element with an Alias, add the Alias-Value to a new designation-tag 
' Date: 12.07.16
' Date: 2020-05-11 added test for existing tagged value designation on attributes, and parsing of Notes for moving text after -- Definition -- keyword into tagged value definition
'
' TBD: Uppercase on first character in package and class designations
' TBD: copy alias and -- Definition -- on association roles, classes and packages into tagged values
' TBD: add test for existing tagged value designation and definition on packages and classes
'
' Project Browser Script main function
'
sub OnProjectBrowserScript()
	
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	
	' Handling Code: Uncomment any types you wish this script to support
	' NOTE: You can toggle comments on multiple lines that are currently
	' selected with [CTRL]+[SHIFT]+[C].
	select case treeSelectedType
	
'		case otElement
'			' Code for when an element is selected
'			dim theElement as EA.Element
'			set theElement = Repository.GetTreeSelectedObject()
'					
		case otPackage
'			' Code for when a package is selected
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			
			'make a msgbox where you can choose OK or Cancel 
			dim message 
			dim box
			box = Msgbox ("The selected package is: [" & thePackage.Name &"]. Starting searching for Alias-tags.", 1)
			select case box
			case vbOK
				findAlias(thePackage)
			case VBcancel
				
			end select 
			
			
'		case otDiagram
'			' Code for when a diagram is selected
'			dim theDiagram as EA.Diagram
'			set theDiagram = Repository.GetTreeSelectedObject()
'			
'		case otAttribute
'			' Code for when an attribute is selected
'			dim theAttribute as EA.Attribute
'			set theAttribute = Repository.GetTreeSelectedObject()
'			
'		case otMethod
'			' Code for when a method is selected
'			dim theMethod as EA.Method
'			set theMethod = Repository.GetTreeSelectedObject()
		
		case else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK
			
	end select
	
end sub


'sub procedure to check if a package/element/attribute has an alias, and if so, copy the Alias-value to a designation tag 
'@param[in]: package (EA.package) The package containing elements with potentially alias tags.
sub findAlias(package)
Session.Output("The current package is: " & package.Name)
			
			dim elements as EA.Collection
			set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages)
			
			dim packages as EA.Collection
			set packages = package.Packages 'collection of packages that belong to this package
			
			dim newTaggedValue as EA.TaggedValue
			set newTaggedValue = nothing
			
			
			'check the chosen package for Alias tag, and if it exists , makes a new designation tag with the same value.  
			if len (package.Element.Alias) > 0 then
						
						Session.Output("   Package [" & package.Element.Name & "] has Alias tag: [" & package.Element.Alias & "]. Copied to designation-tag" )
						'TODO: check for designations already entered
						set newTaggedValue = package.Element.TaggedValues.AddNew ("designation", """"&makeLCName(package.Element.Alias)&""""&"@en")
						newTaggedValue.Update()
						
			end if
			
			' Navigate the package collection and call the findAlias function for each of them
			dim p
			for p = 0 to packages.Count - 1
				dim currentPackage as EA.Package
				set currentPackage = packages.GetAt( p )
				
					
				
				findAlias(currentPackage)
				
			next
				
			' Navigate the elements collection, pick the classes, find the definitions/notes and do sth. with it
			'Session.Output( " number of elements in package: " & elements.Count)
			dim i
			for i = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( i )
				
				
									
					if len (currentElement.Alias) > 0 then
						
						Session.Output("   Class [" & currentElement.Name & "] has Alias tag: [" & currentElement.Alias & "] Copied to designation-tag" )
						'if Alias exist, a new taggedValue ("designation") is made with the same value as Alias
						set newTaggedValue = currentElement.TaggedValues.AddNew( "designation", """"&makeLCName(currentElement.Alias)&""""&"@en" )
						
						newTaggedValue.Update()
						
					end if
					
					
					'check the attributes for Alias-tags.				
				
										
						' Retrieve all attributes for this element
						dim attributesCollection as EA.Collection
						set attributesCollection = currentElement.Attributes
						
						
						
						
						if attributesCollection.Count > 0 then
							 
							dim n
							for n = 0 to attributesCollection.Count - 1 					
								dim currentAttribute as EA.Attribute		
								set currentAttribute = attributesCollection.GetAt(n)
								
								if len (currentAttribute.Alias) > 0 then
									Session.Output( "    Class ["& currentElement.Name &"] \ Attribute [" & currentAttribute.Name & "] has Alias tag [" & currentAttribute.Alias & "] Copied to designation-tag" )
									'if Alias exist, a new taggedValue ("designation") is made with the same value as Alias
									if getTaggedValue(currentAttribute,"designation") = "" then 
										set newTaggedValue = currentAttribute.TaggedValues.AddNew( "designation", """"&makeLCName(currentAttribute.Alias)&""""&"@en" )
										newTaggedValue.Update()
									end if
									
									
								end if
								
								if InStr(currentAttribute.Notes,"-- Definition --") > 0 then
									if getTaggedValue(currentAttribute,"definition") = "" then 
										dim j,l
										j = InStr(currentAttribute.Notes,"-- Definition --")
										Session.Output( "    Class ["& currentElement.Name &"] \ Attribute [" & currentAttribute.Name & "] has -- Definition -- in Notes [" & currentAttribute.Alias & "] Copied to definition-tag" )
										set newTaggedValue = currentAttribute.TaggedValues.AddNew( "definition", """"& Mid(currentAttribute.Notes,j+16) &""""&"@en" )
										currentAttribute.Notes = Trim(Mid(currentAttribute.Notes,1,j-1))
										newTaggedValue.Update()
										currentAttribute.Update()
									end if
									
								end if
							next
						
						end if	
						
						
				
				
			next
		
		Session.Output( "Done with package ["& package.Name &"]")
		Session.Output("---------------------------------------")
		
end sub

function makeLCName(tekst)
		' make name legal NCName
		' (alternatively replace each bad character with a "_", typically used for codelist with proper names.)
		' (Sub settBlankeIKodensNavnTil_(attr))
    Dim txt, txt1, txt2, res, tegn, i, u
    u=0
		'Repository.WriteOutput "Script", "Old code: " & attr.Name,0
		makeLCName = ""
		txt = Trim(tekst)
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
			Repository.WriteOutput "Script", "Error: Unable to construct NCName for code: [" & tekst & "]",0
		else
			txt2 =Mid(res,2,i)
			txt = txt1 + txt2
			if txt <> tekst then
				Repository.WriteOutput "Script", "Change: Old code: [" & tekst & "] changed to new NCName: [" & txt & "]",0
				' return txt
				makeLCName = txt
			end if
		end if

end function

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

'start the main function
OnProjectBrowserScript
