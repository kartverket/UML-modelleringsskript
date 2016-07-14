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
'	- /krav/flersprÃ¥klighet/pakke (partially):
'			Check if there is a tagged value "language" with any content
'	- /krav/12:
'			If datatypes have associations then the datatype shall only be target in a composition
'	- /krav/navning (partially):
'			Check if names of attributes, operations, roles start with lower case and names of packages, 
'			classes and associations start with upper case
'
' Date: 2016-04-13
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
				numberOfErrors = FindInvalidElementsInPackage(thePackage)
				Session.Output( "Antall feil funnet i modellen: " & numberOfErrors)
			else
				Msgbox "Pakke [" & thePackage.Name &"] har ikke stereotype applicationSchema. Velg en pakke med stereotype applicationSchema for Ã¥ starte modellvalidering."
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
			Session.Prompt "[ADVARSEL] Velg en pakke med stereotype applicationSchema for Ã¥ starte modellvalidering.", promptOK
			
	end select
	
end sub


'Check if the provided argument for input parameter theObject fulfills the requirements in krav/3:
'Find elements (classes, attributes, navigable association roles, operations, datatypes) 
'without definition (notes/rolenotes)
'
' @param[in] theObject (EA.ObjectType) The object to check against krav/3, 
'			supposed to be one of the following types: EA.Attribute, EA.Method, EA.Connector, EA.Element
' @param[out] (Integer) The number of errors found for tested elements.
function Krav3(theObject)
	'Declare local variables
	Dim currentAttribute as EA.Attribute
	Dim currentMethod as EA.Method
	Dim currentConnector as EA.Connector
	Dim currentElement as EA.Element
	Dim localCounter
	
	'initialize the local variables
	localCounter = 0
	'Session.Output("Krav 3 with object type: "& theObject.ObjectType)
	Select Case theObject.ObjectType
		Case otElement
			' Code for when the function's parameter is an element
			'Session.Output("Krav 3 test for Element")
			set currentElement = theObject
			
			If currentElement.Notes = "" then
				Session.Output("Error: Class [" & currentElement.Name & "] has no definition. [/krav/3]")	
				localCounter = localCounter + 1
			end if
		Case otAttribute
			' Code for when the function's parameter is an attribute
			'Session.Output("Krav 3 test for Attribute")
			set currentAttribute = theObject
			
			'get the attribute's parent element
			dim attributeParentElement as EA.Element
			set attributeParentElement = Repository.GetElementByID(currentAttribute.ParentID)
			
			if currentAttribute.Notes = "" then
				Session.Output( "Error: Class ["& attributeParentElement.Name &"] \ attribute [" & currentAttribute.Name & "] has no definition. [/krav/3]")
				localCounter = localCounter + 1
			end if
			
		Case otMethod
			' Code for when the function's parameter is a method
			'Session.Output("Krav 3 test for Method")
			set currentMethod = theObject
			
			'get the method's parent element
			dim methodParentElement as EA.Element
			set methodParentElement = Repository.GetElementByID(currentMethod.ParentID)
			
			if currentMethod.Notes = "" then
				Session.Output( "Error: Class ["& methodParentElement.Name &"] \ operation [" & currentMethod.Name & "] has no definition. [/krav/3]")
				localCounter = localCounter + 1
			end if
		Case otConnector
			' Code for when the function's parameter is a connector
			'Session.Output("Krav 3 test for Connector")
			set currentConnector = theObject
			
			'get the necessary connector attributes
			dim sourceEndElementID
			sourceEndElementID = currentConnector.ClientID 'id of the element on the source end of the connector
			dim sourceEndNavigable 
			sourceEndNavigable = currentConnector.ClientEnd.Navigable 'navigability on the source end of the connector
			dim sourceEndName
			sourceEndName = currentConnector.ClientEnd.Role 'role name on the source end of the connector
			dim sourceEndDefinition
			sourceEndDefinition = currentConnector.ClientEnd.RoleNote 'role definition on the source end of the connector
								
			dim targetEndNavigable 
			targetEndNavigable = currentConnector.SupplierEnd.Navigable 'navigability on the target end of the connector
			dim targetEndName
			targetEndName = currentConnector.SupplierEnd.Role 'role name on the target end of the connector
			dim targetEndDefinition
			targetEndDefinition = currentConnector.SupplierEnd.RoleNote 'role definition on the target end of the connector

			dim sourceEndElement as EA.Element
			
			if sourceEndNavigable = "Navigable" and sourceEndDefinition = "" then
				'get the element on the source end of the connector
				set sourceEndElement = Repository.GetElementByID(sourceEndElementID)
				
				Session.Output( "Error: Class ["& sourceEndElement.Name &"] \ Association role [" & sourceEndName & "] has no definition. [/krav/3]")
				localCounter = localCounter + 1
			end if
			
			if targetEndNavigable = "Navigable" and targetEndDefinition = "" then
				'get the element on the source end of the connector (also source end element here because error message is related to the element on the source end of the connector)
				set sourceEndElement = Repository.GetElementByID(sourceEndElementID)
				
				Session.Output( "Error: Class ["& sourceEndElement.Name &"] \ Association role [" & targetEndName & "] has no definition. [/krav/3]")
				localCounter = localCounter + 1
			end if
			
		Case else		
			Session.Output( "Error: Function [Krav3] started with invalid parameter.")
	End Select
	
	Krav3 = localCounter
end function


'sub procedure to check if the package contains classes with multiple inheritance
'@param[in]: currentElement (EA.Element). The element "classe" is potentially with a multiple inheritance.
sub findMultipleInheritance(currentElement)

	dim connectors as EA.Collection 
 	set connectors = currentElement.Connectors 
 					 
 	'iterate the connectors 
 					
 	dim connectorsCounter 
	dim numberOfSuperClasses 
	numberOfSuperClasses = 0 
	dim theTargetGeneralization as EA.Connector
	set theTargetGeneralization = nothing
					
 		for connectorsCounter = 0 to connectors.Count - 1 
			dim currentConnector as EA.Connector 
			set currentConnector = connectors.GetAt( connectorsCounter ) 
						
						
			'check if the connector type is "Generalization" and if so
			'get the element on the source end of the connector  
			if currentConnector.Type = "Generalization"  then
				if currentConnector.ClientID = currentElement.ElementID then 
					
					'count number of classes with a generalization connector on the source side 
					numberOfSuperClasses = numberOfSuperClasses + 1 
					set theTargetGeneralization = currentConnector 
				end if 
							
							
			end if
			'if theres more than one generalization connecter on the source side the class has multiple inheritance
				if numberOfSuperClasses > 1 then
					Session.Output("Error: Found multiple inheritance for class:  " &startClass& ". [/krav/enkelarv]")
					exit for 
						
							
				end if 
						
		next
					
			' if there is just one generalization connector on the source side, start checking genralization connectors for the superclasses 
			if numberOfSuperClasses = 1 and not theTargetGeneralization is nothing then
						
				dim superClassID 
				dim superClass as EA.Element
				'the elementID of the element at the target end
				superClassID =  theTargetGeneralization.SupplierID 
				set superClass = Repository.GetElementByID(superClassID)
			
		
				'Check level of superClass
				call findMultipleInheritance (superClass)
						
			end if 
					
end sub


function FindInvalidElementsInPackage(package)

			
			'Session.Output("The current package is: " & package.Name)
			dim localCounter
			localCounter = 0
			dim elements as EA.Collection
			set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages)
			Dim myDictionary
			dim errorsInFunctionTests
			
			'check package definition
			if package.Notes = "" then
						Session.Output("FEIL: Pakke [" & package.Name & "] mangler definisjon. [/krav/definisjoner]")
						localCounter = localCounter + 1
			end if
			
			'check if first letter of package name is capital letter
			
			if not Left(package.Name,1) = UCase(Left(package.Name,1)) then
						Session.Output("FEIL: Navnet til pakka [" & package.Name & "] skal starte med stor bokstav. [/krav/navning]")
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
				localCounter = localCounter + FindInvalidElementsInPackage(currentPackage)
			next
			
			'------------------------------------------------------------------
			'---ELEMENTS---
			'------------------------------------------------------------------		
			
			' Navigate the elements collection, pick the classes, find the definitions/notes and do sth. with it
			'Session.Output( " number of elements in package: " & elements.Count)
			dim i
			for i = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( i )
				
				'Is the currentElement of type Class? If so, continue conducting some tests. If not continue with the next element.
				if currentElement.Type = "Class" then
									
					'check if there is a definition for the class element (call Krav3 function)
					errorsInFunctionTests = Krav3(currentElement)
					localCounter = localCounter + errorsInFunctionTests
					
					'check if there is there is multiple inheritance for the class element (/krav/enkelArv)
					'initialize the global variable startClass which is needed in subrutine findMultipleInheritance
					startClass = currentElement.Name
					Call findMultipleInheritance(currentElement)
					
					
					'check if first letter of class name is capital letter
					if not Left(currentElement.Name,1) = UCase(Left(currentElement.Name,1)) then
						Session.Output("FEIL: Navnet til klassen [" & currentElement.Name & "] skal starte med stor bokstav. [/krav/navning]")
						localCounter = localCounter + 1
					end if
					
					dim stereotype
					stereotype = currentElement.Stereotype
							
						'------------------------------------------------------------------
						'---ATTRIBUTES---
						'------------------------------------------------------------------					
						
						' Retrieve all attributes for this element
						dim attributesCollection as EA.Collection
						set attributesCollection = currentElement.Attributes
			
						if attributesCollection.Count > 0 then
							dim n
							for n = 0 to attributesCollection.Count - 1 					
								dim currentAttribute as EA.Attribute		
								set currentAttribute = attributesCollection.GetAt(n)
								'check if the attribute has a definition									
								'Call the subfunction with currentAttribute as parameter
								errorsInFunctionTests = Krav3(currentAttribute)
								localCounter = localCounter + errorsInFunctionTests
								
								'check if the attribute's name starts with lower case
								if not Left(currentAttribute.Name,1) = LCase(Left(currentAttribute.Name,1)) then
									Session.Output("FEIL: Navnet til egenskapen [" & currentAttribute.Name & "] til klassen ["&currentElement.Name&"] skal starte med liten bokstav. [/krav/navning]")
									localCounter = localCounter + 1
								end if
							next
						end if	
					
						'------------------------------------------------------------------
						'---ASSOCIATIONS---
						'------------------------------------------------------------------
						
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
								
								'if the connector has a name (optional according to the rules), check if it starts with capital letter
								if not currentConnector.Name = "" and not Left(currentConnector.Name,1) = UCase(Left(currentConnector.Name,1)) then
									Session.Output("FEIL: Navnet til assosiasjonen [" & currentConnector.Name & "] mellom klasse ["& elementOnOppositeSide.Name &"] og klasse [" & currentElement.Name & "] skal starte med stor bokstav. [/krav/navning]")
									localCounter = localCounter + 1
								end if
								
								'check if the elementOnOppositeSide has stereotype "dataType" and this side's end is no composition
								if (Ucase(elementOnOppositeSide.Stereotype) = Ucase("dataType")) and not (currentConnector.ClientEnd.Aggregation = 2) then
									Session.Output( "FEIL: Klasse [<<"&elementOnOppositeSide.Stereotype&">>"& elementOnOppositeSide.Name &"] har assosiasjon til klasse [" & currentElement.Name & "] som ikke er komposisjon pÃ¥ "& currentElement.Name &"-siden. [/krav/12]")									
									localCounter = localCounter + 1
								end if
								'check if this side's element has stereotype "dataType" and the opposite side's end is no composition
								if (Ucase(currentElement.Stereotype) = Ucase("dataType")) and not (currentConnector.SupplierEnd.Aggregation = 2) then
									Session.Output( "FEIL: Klasse [<<"&currentElement.Stereotype&">>"& currentElement.Name &"] har assosiasjon til klasse [" & elementOnOppositeSide.Name & "] som ikke er komposisjon pÃ¥ "& elementOnOppositeSide.Name &"-siden. [/krav/12]")									
									localCounter = localCounter + 1
								end if
								
								'check if there is a definition on navigable ends (navigable association roles) of the connector
								'Call the subfunction with currentConnector as parameter
								errorsInFunctionTests = Krav3(currentConnector)
								localCounter = localCounter + errorsInFunctionTests
																
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
								
								'if there are role names on connector ends (regardless of navigability), check if they start with lower case
								if not sourceEndName = "" and not Left(sourceEndName,1) = LCase(Left(sourceEndName,1)) then
									Session.Output("FEIL: Navnet til rollen [" & sourceEndName & "] på assosiasjonsende i tilknytning til klassen ["& currentElement.Name &"] skal starte med liten bokstav. [/krav/navning]")
									localCounter = localCounter + 1
								end if
								if not (targetEndName = "") and not (Left(targetEndName,1) = LCase(Left(targetEndName,1))) then
									Session.Output("FEIL: Navnet til rollen [" & targetEndName & "] på assosiasjonsende i tilknytning til klassen ["& elementOnOppositeSide.Name &"] skal starte med liten bokstav. [/krav/navning]")
									localCounter = localCounter + 1
								end if
							end if
																				
						next
						
						'------------------------------------------------------------------
						'---OPERATIONS---
						'------------------------------------------------------------------
						
						' Retrieve all operations for this element
						dim operationsCollection as EA.Collection
						set operationsCollection = currentElement.Methods
			
						if operationsCollection.Count > 0 then
							dim operationCounter
							for operationCounter = 0 to operationsCollection.Count - 1 					
								dim currentOperation as EA.Method		
								set currentOperation = operationsCollection.GetAt(operationCounter)
								
								'check if the operations's name starts with lower case
								'TODO: this rule does not apply for constructor operation
								if not Left(currentOperation.Name,1) = LCase(Left(currentOperation.Name,1)) then
									Session.Output("FEIL: Navnet til operasjonen [" & currentOperation.Name & "] til klassen ["&currentElement.Name&"] skal starte med liten bokstav. [/krav/navning]")
									localCounter = localCounter + 1
								end if
								
								'check if there is a definition for the operation (call Krav3 function)
								'call the subroutine with currentOperation as parameter
								errorsInFunctionTests = Krav3(currentOperation)
								localCounter = localCounter + errorsInFunctionTests
								
							next
						end if					
				end if
				
			next
			'summerization
			'Session.Output( "Found " & localCounter & " elements without definition.")
			FindInvalidElementsInPackage = localCounter
		'Session.Output( "Done with package ["& package.Name &"]")
		'TODO: check counter for local elements
		'Session.Output( "There are "& localCounter & " elements without definition in this package.")
		
end function

'global variable 
dim startClass 
OnProjectBrowserScript
