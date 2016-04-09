option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This script contains code from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: Modellvalidering
' Author: Magnus Karge
' Purpose: Validate model elements according to rules defined in the standard SOSI Regler for UML-modellering 5.0
' 	Implemented rules:
'	- krav/3: 
'			Find elements (classes, attributes, navigable association roles, operations, datatypes) 
'	        without definition (notes/rolenotes) in the selected package and subpackages
'   - krav/definisjoner:
'			Same as krav/3 but checks also for definitions of packages
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
			'check if the selected package has stereotype applicationSchema
			if UCase(thePackage.element.stereotype) = UCase("applicationSchema") then
				Msgbox "Starte modellvalidering for pakke [" & thePackage.Name &"]."
				dim numberOfElementsWithoutDefinition
				numberOfElementsWithoutDefinition = FindElementsWithoutDefinitionInPackage(thePackage)
				Session.Output( "Modellen inneholder " & numberOfElementsWithoutDefinition & " elementer med manglende definisjon.")
			else
				Msgbox "Pakke [" & thePackage.Name &"] har ikke stereotype applicationSchema. Velg en pakke med stereotype applicationSchema for å starte modellvalidering."
			end if
			
			
			
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

			
			'Session.Output("The current package is: " & package.Name)
			dim localCounter
			localCounter = 0
			dim elements as EA.Collection
			set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages)
			
			'check package definition
			if package.Notes = "" then
						Session.Output("Pakke [" & package.Name & "] mangler definisjon. [/krav/definisjoner]")
						localCounter = localCounter + 1
			end if
			
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
						Session.Output("Klasse [" & currentElement.Name & "] mangler definisjon. [krav/3]")
						localCounter = localCounter + 1
					else
						'definition to system output
						'Session.Output( "  Definition: " & currentElement.Notes)
					end if
					
					
					dim stereotype
					stereotype = currentElement.Stereotype
					
										
						' Retrieve all attributes for this element
						dim attributesCollection as EA.Collection
						set attributesCollection = currentElement.Attributes
			
						if attributesCollection.Count > 0 then
							dim n
							for n = 0 to attributesCollection.Count - 1 					
								dim currentAttribute as EA.Attribute		
								set currentAttribute = attributesCollection.GetAt(n)
								
								if currentAttribute.Notes = "" then
									'Msgbox currentAttribute.Name & " mangler definisjon"
									Session.Output( "Klasse ["& currentElement.Name &"] \ Egenskap [" & currentAttribute.Name & "] mangler definisjon. [krav/3]")
									'Session.Output("    " & currentAttribute.Name & " has no definition")
									localCounter = localCounter + 1
								else
									'definition to system output
									'Session.Output( "    Definition: " & currentAttribute.Notes  & ".")
								end if
							next
						end if	
						
						'retrieve all associations for this element
						dim connectors as EA.Collection
						set connectors = currentElement.Connectors
					
						'iterate the connectors
						'Session.Output("Found " & connectors.Count & " connectors for featureType " & currentElement.Name)
						dim connectorsCounter
						for connectorsCounter = 0 to connectors.Count - 1
							dim currentConnector as EA.Connector
							set currentConnector = connectors.GetAt( connectorsCounter )
							
							dim sourceElementID
							sourceElementID = currentConnector.ClientID
							
							dim sourceEndNavigable 
							sourceEndNavigable = currentConnector.ClientEnd.Navigable
							dim sourceEndName
							sourceEndName = currentConnector.ClientEnd.Role
							dim sourceEndDefinition
							sourceEndDefinition = currentConnector.ClientEnd.RoleNote
														
							dim targetEndNavigable 
							targetEndNavigable = currentConnector.SupplierEnd.Navigable
							dim targetEndName
							targetEndName = currentConnector.SupplierEnd.Role
							dim targetEndDefinition
							targetEndDefinition = currentConnector.SupplierEnd.RoleNote
														
							'if the current element is on the connectors client side conduct some tests
							'(this condition is nedded to make sure only associations with 
							'source end connected to elements within this applicationSchema package are 
							'checked. Associations with source end connected to elements outside of this
							'package are possibly locked and not editable)
							dim elementOnOppositeSide as EA.Element
							if currentElement.ElementID = sourceElementID then
								'check if there is a definition on navigable ends of the connector
								'Session.Output( "Tester Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & sourceEndName & "] -- definisjon: "& sourceEndDefinition)
								if sourceEndNavigable = "Navigable" and sourceEndDefinition = "" then
									Session.Output( "Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & sourceEndName & "] mangler definisjon. [krav/3]")
									localCounter = localCounter + 1
								end if
								
								'Session.Output( "Tester Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & targetEndName & "] -- definisjon: "& targetEndDefinition)
								if targetEndNavigable = "Navigable" and targetEndDefinition = "" then
									Session.Output( "Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & targetEndName & "] mangler definisjon. [krav/3]")
									localCounter = localCounter + 1
								end if
								
							end if
																				
						next
						
						' Retrieve all operations for this element
						dim operationsCollection as EA.Collection
						set operationsCollection = currentElement.Methods
			
						if operationsCollection.Count > 0 then
							dim operationCounter
							for operationCounter = 0 to operationsCollection.Count - 1 					
								dim currentOperation as EA.Method		
								set currentOperation = operationsCollection.GetAt(operationCounter)
								
								if currentOperation.Notes = "" then
									'Msgbox currentAttribute.Name & " mangler definisjon"
									Session.Output( "Klasse ["& currentElement.Name &"] \ Operasjon [" & currentOperation.Name & "] mangler definisjon. [krav/3]")
									'Session.Output("    " & currentAttribute.Name & " has no definition")
									localCounter = localCounter + 1
								else
									'definition to system output
									'Session.Output( "    Definition: " & currentAttribute.Notes  & ".")
								end if
							next
						end if					
				end if
				
			next
			'summerization
			'Session.Output( "Found " & localCounter & " elements without definition.")
			FindElementsWithoutDefinitionInPackage = localCounter
		'Session.Output( "Done with package ["& package.Name &"]")
		'TODO: check counter for local elements
		'Session.Output( "There are "& localCounter & " elements without definition in this package.")
		
end function

OnProjectBrowserScript
