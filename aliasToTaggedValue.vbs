option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: AliasToTaggedValue
' Author: Sara Henriksen
' Purpose: For every package/attribute/element with an Alias, add the Alias-Value to a new designation-tag 
' Date: 12.07.16
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
							
						set newTaggedValue = package.Element.TaggedValues.AddNew ("designation", """"&package.Element.Alias&""""&"@en")
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
						set newTaggedValue = currentElement.TaggedValues.AddNew( "designation", """"&currentElement.Alias&""""&"@en" )
						
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
									set newTaggedValue = currentAttribute.TaggedValues.AddNew( "designation", """"&currentAttribute.Alias&""""&"@en" )
									newTaggedValue.Update()
									
									
								end if
							next
						
						end if	
						
						
				
				
			next
		
		Session.Output( "Done with package ["& package.Name &"]")
		Session.Output("---------------------------------------")
		
end sub

'start the main function
OnProjectBrowserScript
