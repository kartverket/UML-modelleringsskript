option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: transelateDesignationName
' Author: Sara Henriksen
' Purpose: Bytter modellelementnavn med designation tags, for engelsk og norsk (kan utvides med språk)
' Date: 19.08.16
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
				
			'if UCase(thePackage.element.stereotype) = UCase("applicationSchema") then
			'sjekker om pakken har en language tag, det må den ha for å oversette elementnavnene. 
				dim languageTag AS EA.TaggedValue
				dim taggedCounter
				
				dim languageTagExisit
				languageTagExisit = false 
				
				
				for taggedCounter = 0 to thePackage.Element.TaggedValues.Count - 1
					set languageTag = thePackage.Element.TaggedValues.GetAt(taggedCounter)

		
					if languageTag.Name = "language" then
						languageTagExisit = true 
						
						
						call findDesignation(thePackage)
					end if 

				next

				if not languageTagExisit = true then
					Msgbox "Error: Package [" &thePackage.Name& "] is missing a language-tag"
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


' finner language tagen og lager en variabel som består av verdien til taggen, som skal brukes senere i skriptet. 
dim thePackage as EA.Package
set thePackage = Repository.GetTreeSelectedObject()

dim languageTag AS EA.TaggedValue
dim taggedCounter
				
dim languageTagExisit
languageTagExisit = false 
				
				
	for taggedCounter = 0 to thePackage.Element.TaggedValues.Count - 1
		set languageTag = thePackage.Element.TaggedValues.GetAt(taggedCounter)

		
		if languageTag.Name = "language" then
			languageTagExisit = true 
			dim TVlanguageValue 
			TVlanguageValue = languageTag.Value
								
		end if 

	next
			


sub transelation(theElement, taggedValueName)
	'Session.Output("theElement:  " &theElement.Name)
	

	
	dim newTaggedValueName as EA.TaggedValue
	set newTaggedValueName = nothing 
	dim newElementName as EA.Element.Name
	set newElementName = nothing 
	dim newLanguageTag as EA.TaggedValue
	set newLanguageTag = nothing
	
	if not theElement is nothing and Len(taggedValueName) > 0 then
	
		dim designationTagExist
		designationTagExist = false
		
		dim designationValueMissing
		designationValueMissing = false 
		
		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			dim currentValue
			
			currentValue = currentExistingTaggedValue.Value
			
			
			if currentExistingTaggedValue.Name = taggedValueName then
				
				designationTagExist = true
				
				
				
				if not currentExistingTaggedValue.Value = "" then 
					'Session.Output( "  Funnet tag med navn [" & taggedValueName & "] og verdi: " & currentValue & "")
					
					'lag ny designation tag
					set newTaggedValueName = theElement.TaggedValues.AddNew("designation", """"&theElement.Name&""""&"@"&TVlanguageValue)
					newTaggedValueName.Update()
					
					'fjern "" og landcode på taggen
					dim start, slutt 
					Start = InStr( currentExistingTaggedValue.Value, """" ) 
					slutt = len(currentExistingTaggedValue.Value)- InStr( StrReverse(currentExistingTaggedValue.Value), """" ) -1
					
					'oppdater navnet til elementet, uten "" og @landcode
					theElement.Name = Mid(currentExistingTaggedValue.Value,start+1,slutt)
					theElement.Update()
					
					'slett gammel designation tag
					theElement.TaggedValues.DeleteAt taggedValuesCounter, TRUE			

				end if 	
					
				if currentExistingTaggedValue.Value = "" then 
					designationValueMissing = true 
				end if 					

				
		
				
			
			end if

		next
			'hvis designation tagen ikke finnes, gir advarsel og gjør ingenting med navnet. 
			if designationTagExist = false  then
				Session.Output("the element [" &theElement.Name& "] mangler designation tag")
			end if 
			
			'hvis designation taggen har tom verdi, gir advarsel og gjør ingenting med navnet. 
			if designationValueMissing = true then 
				Session.Output("the element [" &theElement.Name& "] mangler designation verdi")
			end if 
	
	end if 
	
end sub





sub findDesignation(package)
	Session.Output("The current package is: " & package.Name)
	
			dim elements as EA.Collection
			set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages)
			
			dim packages as EA.Collection
			set packages = package.Packages 'collection of packages that belong to this package
			
			
			
			' Navigate the package collection and call the FindTagValuesDesignationWithBlank function for each of them
			dim p
			for p = 0 to packages.Count - 1
				dim currentPackage as EA.Package
				set currentPackage = packages.GetAt( p ) 'getAT
			
				findDesignation(currentPackage) 'går igjennom pakken for å lete etter underpakker
			next
			' Navigate the elements collection
			dim i
			for i = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( i )
				
				
				'Is the currentElement of type Class? If so, continue checking tags and it's attributes' tags. 
				if currentElement.Type =  "Class"  then 	
						
						Call transelation( currentElement, "designation")
						
					
				
					
						dim attributesCollection as EA.Collection
						set attributesCollection = currentElement.Attributes
					
						if attributesCollection.Count > 0 then
							 
								dim n
								for n = 0 to attributesCollection.Count - 1 					
									dim currentAttribute as EA.Attribute		
									set currentAttribute = attributesCollection.GetAt(n)
									'Call TVRemoveBlank(currentAttribute, "designation")
									Call transelation( currentAttribute, "designation") 
								
								next
							
						
						end if
				 
					
						
					
				end if 
				
			next
			
				dim languageTag AS EA.TaggedValue
				dim taggedCounter
				
				dim languageTagExisit
				languageTagExisit = false 
				
				'oppdater language taggen på pakken etter den har byttet språk
				for taggedCounter = 0 to package.Element.TaggedValues.Count - 1
					set languageTag = package.Element.TaggedValues.GetAt(taggedCounter)

		
					if languageTag.Name = "language" then
						if languageTag.Value = "no" then 
							languageTag.Value = "en" 
							languageTag.Update()
							exit for
						end if 
						if languageTag.Value = "en" then 
							languageTag.Value = "no"
							languageTag.Update()
							exit for
						end if
						
					end if 

				next
			
end sub
'Call the main function
OnProjectBrowserScript