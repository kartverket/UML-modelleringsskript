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
'	- /krav/3: 
'			Find elements (classes, attributes, navigable association roles, operations, datatypes) 
'	        without definition (notes/rolenotes) in the selected package and subpackages
'   - /krav/definisjoner (partially implemented except for constraints):
'			Same as krav/3 but checks also for definitions of packages
'	- /krav/10:
'			Check if all navigable association ends have cardinality
'	- /krav/11:
'			Check if all navigable association ends have role names
'	- /krav/flerspråklighet/pakke (partially):
'			Check if there is a tagged value "language" with any content
'	- /krav/12:
'			If datatypes have associations then the datatype shall only be target in a composition
'
' Date: 2016-04-09
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
				dim numberOfErrors
				numberOfErrors = FindNonvalidElementsInPackage(thePackage)
				Session.Output( "Antall feil funnet i modellen: " & numberOfErrors)
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

function FindNonvalidElementsInPackage(package)

			
			'Session.Output("The current package is: " & package.Name)
			dim localCounter
			localCounter = 0
			dim elements as EA.Collection
			set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages)
			
			'check package definition
			if package.Notes = "" then
						Session.Output("FEIL: Pakke [" & package.Name & "] mangler definisjon. [/krav/definisjoner]")
						localCounter = localCounter + 1
			end if
			
			dim packageTaggedValues as EA.Collection
			set packageTaggedValues = package.Element.TaggedValues
			
			'only for applicationSchema packages:
			'iterate the tagged values collection and check if the applicationSchema package has a tagged value "language" with any content
			if UCase(package.element.stereotype) = UCase("applicationSchema") then
				dim taggedValueLanguageMissing
				taggedValueLanguageMissing = true
				dim packageTaggedValuesCounter
				for packageTaggedValuesCounter = 0 to packageTaggedValues.Count - 1
					dim currentTaggedValue as EA.TaggedValue
					set currentTaggedValue = packageTaggedValues.GetAt(packageTaggedValuesCounter)
					if (currentTaggedValue.Name = "language") and not (currentTaggedValue.Value= "") then
						'Session.Output("funnet tagged value"& currentTaggedValue.Name &" = "& currentTaggedValue.Value)
						taggedValueLanguageMissing = false
						exit for
					else 
						if currentTaggedValue.Name = "language" and currentTaggedValue.Value= "" then
							Session.Output("FEIL: Tagged value ["& currentTaggedValue.Name &"] til pakke ["& package.Name & "] mangler verdi. [/krav/flerspråklighet/pakke]")
							localCounter = localCounter + 1
							taggedValueLanguageMissing = false
							exit for
						end if
					end if
				next
				if taggedValueLanguageMissing then
					Session.Output("FEIL: Tagged value [language] mangler på pakke ["& package.Name & "]. [/krav/flerspråklighet/pakke]")
					localCounter = localCounter + 1
				end if
			end if
			
			dim packages as EA.Collection
			set packages = package.Packages 'collection of packages that belong to this package

			' Navigate the package collection and call the FindNonvalidElementsInPackage function for each of them
			dim p
			for p = 0 to packages.Count - 1
				dim currentPackage as EA.Package
				set currentPackage = packages.GetAt( p )
				localCounter = localCounter + FindNonvalidElementsInPackage(currentPackage)
			next

			' Navigate the elements collection, pick the classes, find the definitions/notes and do sth. with it
			'Session.Output( " number of elements in package: " & elements.Count)
			dim i
			for i = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( i )
				
				'Is the currentElement of type Class? If so, continue conducting some tests. If not continue with the next element.
				if currentElement.Type = "Class" then
									
					'check if there is a definition for the class element
					if currentElement.Notes = "" then
						Session.Output("FEIL: Klasse [" & currentElement.Name & "] mangler definisjon. [/krav/3]")
						localCounter = localCounter + 1
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
									Session.Output( "FEIL: Klasse ["& currentElement.Name &"] \ Egenskap [" & currentAttribute.Name & "] mangler definisjon. [/krav/3]")
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
							dim sourceEndCardinality		
							sourceEndCardinality = currentConnector.ClientEnd.Cardinality
							
							dim targetElementID
							targetElementID = currentConnector.SupplierID
							dim targetEndNavigable 
							targetEndNavigable = currentConnector.SupplierEnd.Navigable
							dim targetEndName
							targetEndName = currentConnector.SupplierEnd.Role
							dim targetEndDefinition
							targetEndDefinition = currentConnector.SupplierEnd.RoleNote
							dim targetEndCardinality
							targetEndCardinality = currentConnector.SupplierEnd.Cardinality
							
							'if the current element is on the connectors client side conduct some tests
							'(this condition is needed to make sure only associations with 
							'source end connected to elements within this applicationSchema package are 
							'checked. Associations with source end connected to elements outside of this
							'package are possibly locked and not editable)
							'Session.Output("connectorType: "&currentConnector.Type)
							
							dim elementOnOppositeSide as EA.Element
							if currentElement.ElementID = sourceElementID and not currentConnector.Type = "Realisation" then
								set elementOnOppositeSide = Repository.GetElementByID(targetElementID)
								
								'check if the elementOnOppositeSide has stereotype "dataType" and this side's end is no composition
								if (Ucase(elementOnOppositeSide.Stereotype) = Ucase("dataType")) and not (currentConnector.ClientEnd.Aggregation = 2) then
									Session.Output( "FEIL: Klasse [<<"&elementOnOppositeSide.Stereotype&">>"& elementOnOppositeSide.Name &"] har assosiasjon til klasse [" & currentElement.Name & "] som ikke er komposisjon på "& currentElement.Name &"-siden. [/krav/12]")									
									localCounter = localCounter + 1
								end if
								'check if this side's element has stereotype "dataType" and the opposite side's end is no composition
								if (Ucase(currentElement.Stereotype) = Ucase("dataType")) and not (currentConnector.SupplierEnd.Aggregation = 2) then
									Session.Output( "FEIL: Klasse [<<"&currentElement.Stereotype&">>"& currentElement.Name &"] har assosiasjon til klasse [" & elementOnOppositeSide.Name & "] som ikke er komposisjon på "& elementOnOppositeSide.Name &"-siden. [/krav/12]")									
									localCounter = localCounter + 1
								end if
								
								'check if there is a definition on navigable ends of the connector
								'Session.Output( "Tester Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & sourceEndName & "] -- definisjon: "& sourceEndDefinition)
								if sourceEndNavigable = "Navigable" and sourceEndDefinition = "" then
									Session.Output( "FEIL: Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & sourceEndName & "] mangler definisjon. [/krav/3]")
									localCounter = localCounter + 1
								end if
								
								'Session.Output( "Tester Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & targetEndName & "] -- definisjon: "& targetEndDefinition)
								if targetEndNavigable = "Navigable" and targetEndDefinition = "" then
									Session.Output( "FEIL: Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & targetEndName & "] mangler definisjon. [/krav/3]")
									localCounter = localCounter + 1
								end if
								
								'check if there is multiplicity on navigable ends
								if sourceEndNavigable = "Navigable" and sourceEndCardinality = "" then
									Session.Output( "FEIL: Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & sourceEndName & "] mangler multiplisitet. [/krav/10]")
									localCounter = localCounter + 1
								end if
								
								if targetEndNavigable = "Navigable" and targetEndCardinality = "" then
									Session.Output( "FEIL: Klasse ["& currentElement.Name &"] \ Assosiasjonsrolle [" & targetEndName & "] mangler multiplisitet. [/krav/10]")
									localCounter = localCounter + 1
								end if

								'check if there are role names on navigable ends
								if sourceEndNavigable = "Navigable" and sourceEndName = "" then
									Session.Output( "FEIL: Assosiasjonen mellom klasse ["& currentElement.Name &"] og klasse ["& elementOnOppositeSide.Name & "] mangler rollenavn på navigerbar ende på "& currentElement.Name &"-siden [/krav/11]")
									localCounter = localCounter + 1
								end if
								
								if targetEndNavigable = "Navigable" and targetEndName = "" then
									Session.Output( "FEIL: Assosiasjonen mellom klasse ["& currentElement.Name &"] og klasse ["& elementOnOppositeSide.Name & "] mangler rollenavn på navigerbar ende på "& elementOnOppositeSide.Name &"-siden [/krav/11]")
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
									Session.Output( "FEIL: Klasse ["& currentElement.Name &"] \ Operasjon [" & currentOperation.Name & "] mangler definisjon. [/krav/3]")
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
			FindNonvalidElementsInPackage = localCounter
		'Session.Output( "Done with package ["& package.Name &"]")
		'TODO: check counter for local elements
		'Session.Output( "There are "& localCounter & " elements without definition in this package.")
		
end function

OnProjectBrowserScript
