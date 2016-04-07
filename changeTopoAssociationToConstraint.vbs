option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 	changeTopoAssociationToConstraint
' Author: 		Magnus Karge
' Purpose: 		Find associations with stereotype 'topo' between feature types in the model, 
'				add constraint 'KanAvgrensesAv..' to the class from which the association is pointing towards another class
'				(via a navigable end) and remove the topo association afterwards
' Date: 		06.04.2016
'
' Project Browser Script main function
sub OnProjectBrowserScript()
	
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	
	'find out what type is selected
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
			if UCase(thePackage.element.stereotype) = UCase("applicationSchema") then
				Msgbox "The selected package is: [" & thePackage.Name &"]. Starting search for elements with topo association."
				FindElementsWithTopoAssociationInPackage(thePackage)
			else
				Msgbox "The selected package [" & thePackage.Name &"] has no stereotype applicationSchema. Please select a package with stereotype applicationSchema to run this script."
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
			Session.Prompt "This script does not support items of this type. Please choose a package in order to start the script.", promptOK
			
	end select
	
end sub

'sub procedure to check the content of a given package and all its subpackages and add missing tags to elements
'@param[in]: package (EA.package) The package containing elements with potentially missing tags.
sub FindElementsWithTopoAssociationInPackage(package)
			
			dim elements as EA.Collection
			'collection of elements that belong to this package (classes, notes... BUT NO packages)
			set elements = package.Elements 
			
			dim packages as EA.Collection
			'collection of packages that belong to this package
			set packages = package.Packages 
									
			'navigate the package collection and call the FindElementsWithTopoAssociationInPackage 
			'sub procedure for each of them
			dim packageCounter
			for packageCounter = 0 to packages.Count - 1
				dim currentPackage as EA.Package
				set currentPackage = packages.GetAt( packageCounter )
				FindElementsWithTopoAssociationInPackage(currentPackage)
			next
					
			'navigate the elements collection
			dim elementsCounter
			for elementsCounter = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( elementsCounter )
								
				'check if the currentElement has stereotype FeatureType. 
				if UCase(currentElement.Stereotype) = UCase("FeatureType") then
					Session.Output("-----------------------------")
					Session.Output("Current feature type is: "& currentElement.Name &". Start searching for topo association.")
					
					'check if the feature type has associations with stereotype topo
					dim connectors as EA.Collection
					set connectors = currentElement.Connectors
					
					'navigate the connectors
					'Session.Output("Found " & connectors.Count & " connectors for featureType " & currentElement.Name)
					dim connectorsCounter
					for connectorsCounter = 0 to connectors.Count - 1
						dim currentConnector as EA.Connector
						set currentConnector = connectors.GetAt( connectorsCounter )
						
						'check if the connector has stereotype topo if not ignore this one
						if UCase(currentConnector.Stereotype) = UCase("topo") then
							'Session.Output("...with clientID: " & currentConnector.ClientID)
							dim sourceElementID
							sourceElementID = currentConnector.ClientID 
							dim sourceEndNavigable 
							sourceEndNavigable = currentConnector.ClientEnd.Navigable
							dim targetElementID
							targetElementID = currentConnector.SupplierID
							dim targetEndNavigable 
							targetEndNavigable = currentConnector.SupplierEnd.Navigable
							dim oppositeSideNavigable 
							dim currentElementSideNavigable 
							dim oppositeSideSharedAggregation
							dim currentElementSharedAggregation
							
							'find out which side is the opposite one
							dim elementOnOppositeSide as EA.Element
							if currentElement.ElementID = sourceElementID then
								'current element is on the connectors client side
								'get the needed information
								currentElementSideNavigable = sourceEndNavigable
								oppositeSideNavigable = targetEndNavigable
								
								if currentConnector.ClientEnd.Aggregation = 1 then
									currentElementSharedAggregation = true
								else
									currentElementSharedAggregation = false
								end if
								
								if currentConnector.SupplierEnd.Aggregation = 1 then
									oppositeSideSharedAggregation = true
								else
									oppositeSideSharedAggregation = false
								end if
								
								set elementOnOppositeSide = Repository.GetElementByID(targetElementID)
								
							else
								'current element is on the connectors supplier side
								'get the needed information
								currentElementSideNavigable = targetEndNavigable
								oppositeSideNavigable = sourceEndNavigable
								
								if currentConnector.ClientEnd.Aggregation = 1 then
									oppositeSideSharedAggregation = true
								else
									oppositeSideSharedAggregation = false
								end if
								
								if currentConnector.SupplierEnd.Aggregation = 1 then
									currentElementSharedAggregation = true
								else
									currentElementSharedAggregation = false
								end if
								
								set elementOnOppositeSide = Repository.GetElementByID(sourceElementID)
							end if
							Session.Output("Found topo association from " & currentElement.Name & " to " & elementOnOppositeSide.Name)
							
							'check if the current element owns at least one (inherited) attribute of type 'Flate'
							dim hasFlateAttributeType
							hasFlateAttributeType = false
							dim currentElementAttributeCollection as EA.Collection
							set currentElementAttributeCollection = currentElement.AttributesEx
							dim currentElementAttributeCollectionCounter
							for currentElementAttributeCollectionCounter = 0 to currentElementAttributeCollection.Count - 1
								dim currentAttribute as EA.Attribute
								set currentAttribute = currentElementAttributeCollection.GetAt(currentElementAttributeCollectionCounter)
								'Session.Output("Attribute name:"& currentAttribute.Name & "-- Attribute type: "&  currentAttribute.Type)
								if UCase(currentAttribute.Type) = "FLATE" then
									hasFlateAttributeType = true
									'exit for
								end if
							next


							'if the current element has an attribute of type Flate and the topo association is directed towards the opposite side of the current element but not bi-directional
							'or has a shared aggregation on the current element's side and not on the opposite side, start processing
							if hasFlateAttributeType and (((oppositeSideNavigable = "Navigable") and not (currentElementSideNavigable = "Navigable")) or ((currentElementSharedAggregation = true) and (oppositeSideSharedAggregation = false))) then
								Session.Output("Processing topo association from " & currentElement.Name & " to " & elementOnOppositeSide.Name)
								'find out if there already is a constraint 'KanAvgrensesAv..'
								dim currentElementConstraints as EA.Collection
								set currentElementConstraints = currentElement.Constraints
								dim constraintsCounter
								dim constraintExists
								constraintExists = false
								dim currentConstraint as EA.Constraint
								for constraintsCounter = 0 to currentElementConstraints.Count - 1
									
									set currentConstraint = currentElementConstraints.GetAt(constraintsCounter)
									dim currentConstraintName
									currentConstraintName = currentConstraint.Name
									'Session.Output("Found constraint "& currentConstraintName)
									dim firstPartOfCurrentConstraintName
									firstPartOfCurrentConstraintName = Left(currentConstraintName,14)
									'Session.Output("First part of constraint name "& firstPartOfCurrentConstraintName)
									if (firstPartOfCurrentConstraintName = "KanAvgrensesAv") then
										Session.Output("Constraint 'KanAvgrensesAv..' already exists for feature type "& currentElement.Name)
										constraintExists = true
										exit for
									end if	
								next
								
								dim elementNameOnOppositeSide
								elementNameOnOppositeSide = elementOnOppositeSide.Name
								'Session.Output("Name for the element on opposite side: "& elementNameOnOppositeSide)
								if constraintExists then
									'Session.Output("constraint 'KanAvgrensesAv..' already exists")
									'check if it contains the name of the element on the opposite side of the topo association
									If InStr(currentConstraintName, elementNameOnOppositeSide) = 0 Then
										'if not add the element name to the constraint and remove the topo association
										Session.Output(elementNameOnOppositeSide & " not included in constraint")
										currentConstraint.Name = currentConstraintName & ", " & elementNameOnOppositeSide
										currentConstraint.Update()
										Session.Output("added element name to constraint: "& elementNameOnOppositeSide)
										Session.Output("new constraint name: "& currentConstraint.Name)
										currentElement.Connectors.Delete(connectorsCounter)
										Session.Output("Removed topo association from "& currentElement.Name & " to " & elementNameOnOppositeSide)
									else
										'if so just remove the topo association
										Session.Output(elementNameOnOppositeSide & " already included in constraint")
										currentElement.Connectors.Delete(connectorsCounter)
										Session.Output("Removed topo association from "& currentElement.Name & " to " & elementNameOnOppositeSide)
									End If
									
									
								else
									'constraint does not exist - create one containing the name of the element on the opposite side of the association
									dim newConstraint as EA.Constraint
									'newConstraint.Name = "KanAvgrensesAv "& elementNameOnOppositeSide
									'newConstraint.Update()
									'add new constraint to the constraint collection
									set newConstraint = currentElementConstraints.AddNew("KanAvgrensesAv "& elementNameOnOppositeSide,"OCL")
									newConstraint.Update()
									currentElementConstraints.Refresh()
									Session.Output("Added new constraint: "& newConstraint.Name)
									currentElement.Connectors.Delete(connectorsCounter)
									Session.Output("Removed topo association from "& currentElement.Name & " to " & elementNameOnOppositeSide)
								end if
																					
															
							else
								Session.Output("Topo association from "& currentElement.Name &" to "& elementOnOppositeSide.Name &" will be ignored because either has "& currentElement.Name &" no (inherited) attribute type 'Flate' or association's navigability or aggregation does not match search criteria.")
							end if

						end if
					next
					currentElement.Connectors.Refresh()
					Repository.RefreshOpenDiagrams(true)
				end if
				
				'Session.Output("Done with element ["& currentElement.Name &"]")
				
			next
	'Session.Output( "Done with package ["& package.Name &"]")
			
end sub



'start the main function
OnProjectBrowserScript
