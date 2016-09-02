option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: checkLanguageDesignationDefinitionTag
' Author:Sara Henriksen
' Purpose: check package with stereotype ApplicationSchema for language, designation and definition tags. If the package is missing one of the tags, return an error
' if the langauge tag exists, check the value of the tag. 
' Date: 15.08.16
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
			
			finDesignationTAG(thePackage)
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
			Session.Prompt "This script does not support items of this type.", promptOK
			
	end select
	
end sub

'sub procedure to check if an ApplicationSchema got the tags: language, designation and definition, and check if the value of the tags are correct. Retruns an
'error if the tags dont exist, or the value of the tags are not right. 
'@param[in]: theElement (Element Class) and TaggedValueName (String)
sub checkLanguageDesignationDefinitionTag (theElement,  taggedValueName)

'check if the package got a stereotype ApplicationSchema 
if UCase(theElement.Stereotype) = "APPLICATIONSCHEMA" then 
	'searching for definition tag
	if taggedValueName = "definition" then
		if not theElement is nothing and Len(taggedValueName) > 0 then
		
		
				
				dim taggedValueDefinitionMissing
				taggedValueDefinitionMissing = true
				
				'iterate trough all taggedValues for the package 
				dim currentExistingTaggedValue AS EA.TaggedValue 
				dim taggedValuesCounter
				for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
					set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)

					'check if the element has a tagged value with the provided name 
						if currentExistingTaggedValue.Name = taggedValueName then
						
							'remove spaces before and after a string, if the value only contains blanks  the value is empty
							currentExistingTaggedValue.Value = Trim(currentExistingTaggedValue.Value)
							'returns an error if the value of the tag is empty
							if len (currentExistingTaggedValue.Value) = 0 then 
								'Session.Output("Error:  Package [" &theElement.Name & "] has an empty definition-tag [/krav/taggedValueSpråk]"   )
								taggedValueDefinitionMissing = false 
								exit for 
							end if  
							
							'check if the value has the right structure. "{value}"@en, if not, return an error. 	
							if (mid(StrReverse(currentExistingTaggedValue.Value),1,2)) = "ne" or (mid(StrReverse(currentExistingTaggedValue.Value),1,2)) = "on" then 
								'Session.Output("[" &theElement.Name& "] has definition tag :  " &currentExistingTaggedValue.Value)
								taggedValueDefinitionMissing = false 'exit for 
							else	
								'if not the structure above is correct, return an warning. 
								'Session.Output("Warning: Package [" &theElement.Name& "] has definition tag: " &currentExistingTaggedValue.Value& ". The tag is not norwegian or english. [/krav/taggedValueSpråk]")
								taggedValueDefinitionMissing = false
								
							end if
							
						end if
					  	
				next
					'if taggedValueDefinitionMissing = true, then the provided tag doesn't exist and it returns an error 
					if taggedValueDefinitionMissing then
						Session.Output("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name& "] lacks [definition] tag. [/krav/taggedValueSpråk]")
					end if
					
		end if 

	end if 


	'searching for designation tag
	if taggedValueName = "designation" then 
		if not theElement is nothing and Len(taggedValueName) > 0 then	
		
				dim taggedValueDesignationMissing 
				taggedValueDesignationMissing = true
				
				'iterate trough all tagged values for the package
				dim currentExistingTaggedValue1 AS EA.TaggedValue 
				dim taggedValuesCounter1
				
				for taggedValuesCounter1 = 0 to theElement.TaggedValues.Count - 1
					set currentExistingTaggedValue1 = theElement.TaggedValues.GetAt(taggedValuesCounter1)

					'check if the element has a tagged value with the provided name	
					if currentExistingTaggedValue1.Name = taggedValueName then
					
							'remove spaces before and after a string, if the value only contains blanks  the value is empty
							currentExistingTaggedValue1.Value = Trim(currentExistingTaggedValue1.Value)
							'check if the value of the tag is empty, if so, returns an error 
							if len (currentExistingTaggedValue1.Value) = 0 then 
								'Session.Output("Error:  Package [" &theElement.Name & "] has an empty designation tag [/krav/taggedValueSpråk]"   )
								taggedValueDesignationMissing = false 
								exit for 
							end if  
							
							'check that the value of the tag got the correct structure:  "{value}"@en (must be english, if not, returns an warning)
							if (mid(StrReverse(currentExistingTaggedValue1.Value),1,2)) = "ne" or (mid(StrReverse(currentExistingTaggedValue1.Value),1,2)) = "on" then 
								'Session.Output("[" &theElement.Name& "] has designation tag :  " &currentExistingTaggedValue1.Value)
								taggedValueDesignationMissing = false 
							else  
								'Session.Output("Warning: [" &theElement.Name& "] has designation tag: " &currentExistingTaggedValue1.Value& ". The tag is not norwegian or english. [/krav/taggedValueSpråk]")
								taggedValueDesignationMissing = false
							end if
							
						
					  
						
					end if 
				next
'					'if the tag doesn't exist for the package, retruns an error 
					if taggedValueDesignationMissing then 
						Session.Output("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name& "] lacks [designation] tag. [/krav/taggedValueSpråk]")
					end if
		
		
		end if 
	end if 
	
	
	'searching for designation tag
	if taggedValueName = "language" then
		if not theElement is nothing and Len(taggedValueName) > 0 then
		
		
				
				dim taggedValueLanguageMissing
				taggedValueLanguageMissing = true
				
				'iterate trough all tagged values for the package
				dim currentExistingTaggedValue2 AS EA.TaggedValue 
				dim taggedValuesCounter2
				for taggedValuesCounter2 = 0 to theElement.TaggedValues.Count - 1
					set currentExistingTaggedValue2 = theElement.TaggedValues.GetAt(taggedValuesCounter2)
	
'			
'			
'						 'check if the element has a tagged value with the provided name
						if currentExistingTaggedValue2.Name = taggedValueName then
							'remove spaces before and after a string, if the value only contains blanks  the value is empty
							currentExistingTaggedValue2.Value = Trim(currentExistingTaggedValue2.Value)
							'if the value is nothing, return an error 
							if len (currentExistingTaggedValue2.Value) = 0 then 
								Session.Output("ERROR:  Package [" &theElement.Name & "] has an empty language tag [/krav/taggedValueSpråk]"   )
								taggedValueLanguageMissing = false 
								exit for 
							end if  
							
							'check if the value is "no" for norwegian or "en" for english, if not, returns an warning. 	
							if currentExistingTaggedValue2.Value = "no" or currentExistingTaggedValue2.Value = "en" then 
								'Session.Output("Package [" &theElement.Name& "] has language tag :  " &currentExistingTaggedValue2.Value)
								taggedValueLanguageMissing = false 'exit for 
							else  
								Session.Output("Warning: Package [" &theElement.Name& "] has language tag: " &currentExistingTaggedValue2.Value& ". The language tag is not norwegian or english. [/krav/taggedValueSpråk]")
								taggedValueLanguageMissing = false
							end if
							
						end if 
					
'			
'
				next
				'if the package doesn't have a tag with the provided name, return an error. 
				if taggedValueLanguageMissing then
					Session.Output ("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name& "] lacks [language] tag. [/krav/taggedValueSpråk]")
				end if
'		
		
		end if 
	end if 

end if 
end sub 

'sub procedure that find the package name and calls another sub procedure (checkLanguageDesignationDefinitionTag) 
'@param[in]: package (Package Class)
sub finDesignationTAG(package)


	dim packages as EA.Collection
	set packages = package.Packages
	
	Call checkLanguageDesignationDefinitionTag (package.Element, "definition") 
	Call checkLanguageDesignationDefinitionTag (package.Element, "designation")
	Call checkLanguageDesignationDefinitionTag (package.Element, "language")

end sub
OnProjectBrowserScript
