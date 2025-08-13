option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 	AddMissingTags
' Author: 		Magnus Karge
' Purpose: 		To add missing tags defined in the Norwegian standard "SOSI regler for UML-modellering" 
' 				to model elements (application schemas, feature types & attributes, data types & attributes,
'				code lists, enumerations)
' Date: 		11.09.2015
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
			Msgbox "The selected package is: [" & thePackage.Name &"]. Starting search for elements with missing tags."
			FindElementsWithMissingTagsInPackage(thePackage)
			
'			
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
sub FindElementsWithMissingTagsInPackage(package)
			
			'Session.Output("The current package is: " & package.Name)
			'if the current package has stereotype applicationSchema then check tagged values
			if package.element.stereotype = "applicationSchema" or package.element.stereotype = "ApplicationSchema" then
				Call TVSetElementTaggedValue(package.element, "SOSI_kortnavn")
				Call TVSetElementTaggedValue(package.element, "SOSI_modellstatus")
				Call TVSetElementTaggedValue(package.element, "targetNamespace")
				Call TVSetElementTaggedValue(package.element, "xmlns")
				Call TVSetElementTaggedValue(package.element, "xsdDocument")
				Call TVSetElementTaggedValue(package.element, "language")
				Call TVSetElementTaggedValue(package.element, "version")
			end if
			
			dim elements as EA.Collection
			'collection of elements that belong to this package (classes, notes... BUT NO packages)
			set elements = package.Elements 
			
			dim packages as EA.Collection
			'collection of packages that belong to this package
			set packages = package.Packages 
									
			'navigate the package collection and call the FindElementsWithMissingTagsInPackage 
			'sub procedure for each of them
			dim packageCounter
			for packageCounter = 0 to packages.Count - 1
				dim currentPackage as EA.Package
				set currentPackage = packages.GetAt( packageCounter )
				FindElementsWithMissingTagsInPackage(currentPackage)
			next
					
			'navigate the elements collection
			dim elementsCounter
			for elementsCounter = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( elementsCounter )
				
				'Session.Output("The current element is: " & currentElement.Name & " [Stereotype: " & currentElement.Stereotype & "]")
				
				'check if the currentElement has stereotype FeatureType. 
				if ((currentElement.Stereotype = "FeatureType") or (currentElement.Stereotype = "featureType")) then
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
					Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
					Call TVSetElementTaggedValue(currentElement, "isCollection")
					Call TVSetElementTaggedValue(currentElement, "byValuePropertyType")					
					Call TVSetElementTaggedValue(currentElement, "noPropertyType")
				end if
				
				'check if the currentElement has stereotype CodeList. 
				if ((currentElement.Stereotype = "CodeList") or (currentElement.Stereotype = "codeList")) then
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
					Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
					Call TVSetElementTaggedValue(currentElement, "SOSI_datatype")
					Call TVSetElementTaggedValue(currentElement, "SOSI_lengde")
					Call TVSetElementTaggedValue(currentElement, "asDictionary")
				end if
				
				'check if the currentElement has stereotype dataType. 
				if ((currentElement.Stereotype = "DataType") or (currentElement.Stereotype = "dataType")) then
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
					Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
				end if
				
				'check if the currentElement has stereotype enumeration. 
				if ((currentElement.Stereotype = "Enumeration") or (currentElement.Stereotype = "enumeration")) then
					Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
				end if
				
				'if the currentElement has stereotype dataType or FeatureType then 
				'navigate the attributes and check for missing tags
				if ((currentElement.Stereotype = "DataType") or (currentElement.Stereotype = "dataType") or (currentElement.Stereotype = "FeatureType") or (currentElement.Stereotype = "featureType")) then
					dim attributesCounter
					for attributesCounter = 0 to currentElement.Attributes.Count - 1
						dim currentAttribute as EA.Attribute
						set currentAttribute = currentElement.Attributes.GetAt ( attributesCounter )
						'Session.Output( "  The current attribute is ["& currentAttribute.Name &"]")
						'call sub function TVSetElementTaggedValue
						'one function call for each of the required tags
						Call TVSetElementTaggedValue(currentAttribute, "SOSI_navn")
						Call TVSetElementTaggedValue(currentAttribute, "SOSI_datatype")
						Call TVSetElementTaggedValue(currentAttribute, "SOSI_lengde")
						Call TVSetElementTaggedValue(currentAttribute, "inLineOrByReference")
						Call TVSetElementTaggedValue(currentAttribute, "isMetadata")
					next
				end if	
				
				'Session.Output( "Done with element ["& currentElement.Name &"]")
			next
	'Session.Output( "Done with package ["& package.Name &"]")
			
end sub


' Sets the specified TaggedValue on the provided element. If the provided element does not already
' contain a TaggedValue with the specified name, a new TaggedValue is created with the requested
' name. If a TaggedValue already exists with the specified name then nothing will be changed.
'
' @param[in] theElement (EA.Element) The element to set the TaggedValue value on
' @param[in] taggedValueName (String) The name of the TaggedValue to set
'
sub TVSetElementTaggedValue( theElement, taggedValueName)
	'Session.Output( "  Checking if tagged value [" & taggedValueName & "] exists")
	if not theElement is nothing and Len(taggedValueName) > 0 then
		dim newTaggedValue as EA.TaggedValue
		set newTaggedValue = nothing
		dim taggedValueExists
		taggedValueExists = False
		
		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			if currentExistingTaggedValue.Name = taggedValueName then
				taggedValueExists = True
			end if
		next
				
		'if the element does not contain a tagged value with the provided name, create a new one
		if not taggedValueExists = True then
			set newTaggedValue = theElement.TaggedValues.AddNew( taggedValueName, "" )
			newTaggedValue.Update()
			'Session.Output( "    ADDED tagged value ["& taggedValueName &"]")
		else 
			'Session.Output( "    FOUND tagged value ["& taggedValueName &"]")
		end if
	end if
end sub

'start the main function
OnProjectBrowserScript
