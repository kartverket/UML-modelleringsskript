option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: deleteTags
' Author: Sara Henriksen
' Purpose: Find and delete all SOSI_melding and RationalRose tags from packages, attributes, elements.. 
' Date: 07.07.16 
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
'			' Code for when a package is selected
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			
			'make a msgbox where you can choose OK or Cancel 
			dim message 
			dim box
			box = Msgbox ("The selected package is: [" & thePackage.Name &"]. Starting searching for tags to delete.", 1)
			select case box
			case vbOK
				FindTagValuesToDelete(thePackage)
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




'sub procedure to check if the tagged value contains the word RationalRose, and if so, deletes it. 
'@param[in]: theElement (Element Class) and TaggedValueName (String)
sub DeleteRationalRoseTag( theElement, taggedValueName)
'Session.Output( "  starter delete RationalRose tag med element [" & theElement.Name & "] og taggedValue [" & taggedValueName & "] ")
	'Session.Output( "  Checking if tagged value [" & taggedValueName & "] exists")
	if not theElement is nothing and Len(taggedValueName) > 0 then
	
		
		
		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			dim currentValue 
			currentValue = currentExistingTaggedValue.Value
			
			
				'check if the first 12 characters of the tag are RATIONALROSE and if so, deletes it. 
				if UCase(Mid(currentExistingTaggedValue.Name,1,12))= taggedValueName then
					Session.Output("[ " &theElement.Name & "] har RationalRose tag : [" & currentExistingTaggedValue.Name & "] "   )
					
				
					theElement.TaggedValues.DeleteAt taggedValuesCounter, FALSE
					Session.Output("RationalRose-tag : [" & currentValue & "]  slettet"   )
				
				end if


			
			

		next
		theElement.TaggedValues.Refresh()
		
		
	end if
	
end sub


'sub procedure to check if the tagged value exist with the provided name exist, and if so, deletes it.  
'@param[in]: theElement (Element Class) and TaggedValueName (String) 
sub DeleteTag( theElement, taggedValueName)
'Session.Output( "  starter delete SOSI_melding tag med element [" & theElement.Name & "] og taggedValue [" & taggedValueName & "] ")
	'Session.Output( "  Checking if tagged value [" & taggedValueName & "] exists")
	if not theElement is nothing and Len(taggedValueName) > 0 then
		
		
		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.TaggedValue 
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			dim currentValue 
			currentValue = currentExistingTaggedValue.Value
			
			
				'check if the 
				if currentExistingTaggedValue.Name = taggedValueName then
				
					
					Session.Output("[ " &theElement.Name & "] har SOSI_melding tag : [" & currentExistingTaggedValue.Name & "] "   )
					theElement.TaggedValues.DeleteAt taggedValuesCounter, FALSE
					Session.Output("SOSI_melding-tag : [" & currentValue & "]  slettet"   )
				
				end if


			
			

		next
		theElement.TaggedValues.Refresh()
		
		
	end if
	
end sub

'sub procedure to navigate trough all the packages, attributes and classes, and calls the DeleteTag, DeleteRationalRoseTag to see if the element has 
' a tag with the provided name to delete. 
'@param[in]: package (EA.package) The package containing elements with potentially SOSI_melding or RationalRose..tags
sub FindTagValuesToDelete(package)
	Session.Output("The current package is: " & package.Name)
	
			dim elements as EA.Collection
			set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages)
			
			dim packages as EA.Collection
			set packages = package.Packages 'collection of packages that belong to this package
			
			
			Call DeleteTag( package.Element , "SOSI_melding" ) 'searching for SOSI_melding tags 
			Call DeleteRationalRoseTag( package.Element , "RATIONALROSE" ) 'searching for RationalRose... tags  
			' Navigate the package collection and call the FindTagValuesToDelete function for each of them
			dim p
			for p = 0 to packages.Count - 1
				dim currentPackage as EA.Package
				set currentPackage = packages.GetAt( p ) 'getAT
				
				
				FindTagValuesToDelete(currentPackage) 'looking for packages in the package 
				
				
				
				
			next
			' Navigate the elements collection, pick the classes, find the taggedValues/designation and do sth. with it
			'Session.Output( " number of elements in package: " & elements.Count)
			dim i
			for i = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( i )
				
				
				'Is the currentElement of type Class? 
				if currentElement.Type =  "Class"  then 	
					
					
					Call DeleteTag( currentElement, "SOSI_melding") 'searching for SOSI_melding tags 
					Call DeleteRationalRoseTag( currentElement, "RATIONALROSE") 'searching for RationalRose.. tags  
				
				
					dim attributesCollection as EA.Collection
					set attributesCollection = currentElement.Attributes
					
					if attributesCollection.Count > 0 then
							 
							dim n
							for n = 0 to attributesCollection.Count - 1 					
								dim currentAttribute as EA.Attribute		
								set currentAttribute = attributesCollection.GetAt(n)
								
								Call DeleteTag( currentAttribute, "SOSI_melding") 'searching for SOSI_melding tags
								Call DeleteRationalRoseTag( currentAttribute, "RATIONALROSE") 'searching for RationalRose...tags
							next
							
						
					end if
					
						
				end if
				
			next
			Session.Output ("ferdig med pakke    " &package.Name)
			Session.Output ("-----------------------------------")
end sub

'start the main function
OnProjectBrowserScript