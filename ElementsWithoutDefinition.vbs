option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: ElementsWithoutDefinition
' Author: Magnus Karge
' Purpose: Find elements (classes, attributes) without definition (notes) in the selected package and subpackage
' Date: 15.07.2015
'

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
			' Code for when a package is selected
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			Msgbox "The selected package is: [" & thePackage.Name &"]. Starting search for classes and attributes without definition."
			dim numberOfElementsWithoutDefinition
			numberOfElementsWithoutDefinition = FindElementsWithoutDefinitionInPackage(thePackage)
			Session.Output( "There are " & numberOfElementsWithoutDefinition & " elements without definition in the selected package and subpackages.")
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

function FindElementsWithoutDefinitionInPackage(package)

			
			Session.Output("The current package is: " & package.Name)
			dim localCounter
			localCounter = 0
			dim elements as EA.Collection
			set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages)
			
			dim packages as EA.Collection
			set packages = package.Packages 'collection of packages that belong to this package

			' Navigate the package collection and call the findElementsWithoutDefinitionInPackage function for each of them
			dim p
			for p = 0 to packages.Count - 1
				dim currentPackage as EA.Package
				set currentPackage = packages.GetAt( p )
				localCounter = localCounter + FindElementsWithoutDefinitionInPackage(currentPackage)
			next

			' Navigate the elements collection, pick the classes, find the definitions/notes and do sth. with it
			'Session.Output( " number of elements in package: " & elements.Count)
			dim i
			for i = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( i )
				
				'Is the currentElement of type Class? If so, continue checking definition and it's attributes' definitions. If not continue with the next element.
				if currentElement.Type = "Class" then
									
					'Session.Output( "Found class " & currentElement.Name )
					if currentElement.Notes = "" then
						'Msgbox currentElement.Name & "mangler definisjon"
						Session.Output("   Class [" & currentElement.Name & "] has no definition.")
						localCounter = localCounter + 1
					else
						'definition to system output
						'Session.Output( "  Definition: " & currentElement.Notes)
					end if
					
					
					dim stereotype
					stereotype = currentElement.Stereotype
					'is the class a codeList? If so don't check the attributes' definitions.				
					if stereotype <> "codeList" and stereotype <> "CodeList" then
										
						' Retrieve all attributes for this element
						dim attributesCollection as EA.Collection
						set attributesCollection = currentElement.Attributes
			
						if attributesCollection.Count > 0 then
							'set localAttributesWithoutDefinition 
							dim n
							for n = 0 to attributesCollection.Count - 1 					
								dim currentAttribute as EA.Attribute		
								set currentAttribute = attributesCollection.GetAt(n)
								
								if currentAttribute.Notes = "" then
									'Msgbox currentAttribute.Name & " mangler definisjon"
									Session.Output( "    Class ["& currentElement.Name &"] \ Attribute [" & currentAttribute.Name & "] has no definition")
									'Session.Output("    " & currentAttribute.Name & " has no definition")
									localCounter = localCounter + 1
								else
									'definition to system output
									'Session.Output( "    Definition: " & currentAttribute.Notes  & ".")
								end if
							next
						end if	
					else 
						Session.Output( "The class [" & currentElement.Name & "] has stereotype <<"& stereotype & ">>, attributes will not be checked." )
					end if
					
				end if
				
			next
			'summerization
			'Session.Output( "Found " & localCounter & " elements without definition.")
			FindElementsWithoutDefinitionInPackage = localCounter
		Session.Output( "Done with package ["& package.Name &"]")
		'TODO: check counter for local elements
		'Session.Output( "There are "& localCounter & " elements without definition in this package.")
		
end function

OnProjectBrowserScript
