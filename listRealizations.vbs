option explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: 	listRealizations
' version:		0.4.0
' Author: 		Tore Johnsen
' Purpose: 		Checks realizations in the selected package.
'
' Comments:		In early stage of development.
'				
' Date: 2018-06-05


dim globalErrorCounter1
globalErrorCounter1 = 0
dim globalErrorCounter2
globalErrorCounter2 = 0
dim localErrorCounter1
localErrorCounter1 = 0
dim localErrorCounter2
localErrorCounter2 = 0

dim counter1
counter1 = 0
dim counter2
counter2 = 0
dim counter3
counter3 = 0

dim realizedCounter
realizedCounter = 0
dim newOrChangedCounter
newOrChangedCounter = 0
dim notRealizedCounter
notRealizedCounter = 0
dim totalAttributes
totalAttributes = 0

dim packageCollectionAs
set packageCollectionAs = CreateObject("System.Collections.ArrayList")
dim packageCollectionAsGUID
set packageCollectionAsGUID = CreateObject("System.Collections.ArrayList")
dim packageCollectionOf
set packageCollectionOf = CreateObject("System.Collections.ArrayList")
dim packageCollectionOfGUID
set packageCollectionOfGUID = CreateObject("System.Collections.ArrayList")

sub OnProjectBrowserScript()

	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	select case treeSelectedType
		case otPackage
		' Code for when a package is selected
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			Dim StartTime, EndTime, Elapsed
			StartTime = timer 
						
			Session.Output("Start check on package: [«"&thePackage.StereotypeEx&"» "&thePackage.Name&"] "&Now&"")
			checkPackage(thePackage)
			Session.Output(" ")
			Session.Output(" ")
			
			Session.Output("Packages realized from:")
			if packageCollectionOf.Count > 0 then
				dim dX
				for each dX in packageCollectionOf
					Session.Output(" "&dX&"")
				next
			else
				Session.Output("0")
			end if
'			Session.Output("DEBUG GUIDS:")
'			if packageCollectionOfGUID.Count > 0 then
'				dim dX2
'				for each dX2 in packageCollectionOfGUID
'					Session.Output(" "&dX2&"")
'				next
'			else
'				Session.Output("0")
'			end if
			
			Session.Output(" ")
			
			Session.Output("Realized in packages:")
			if packageCollectionAs.Count > 0 then
				dim eX
				for each eX in packageCollectionAs
					Session.Output(" "&eX&"")
				next
			else
				Session.Output("0")
			end if
'			Session.Output("DEBUG GUIDS:")
'			if packageCollectionAsGUID.Count > 0 then
'				dim eX2
'				for each eX2 in packageCollectionAsGUID
'					Session.Output(" "&eX2&"")
'				next
'			else
'				Session.Output("0")
'			end if
			
			Session.Output(" ")
			Session.Output("Lets go reverse!")
			Session.Output(" ")
			Call letsGoReverse
			
			Session.Output("End check on package: [«"&thePackage.StereotypeEx&"» "&thePackage.Name&"] "&Now&" ✓")
			Elapsed = formatnumber((Timer - StartTime),0)
			Session.Output("Run time: " &Elapsed& " seconds" )
			'Session.Output("counter1(actual): "&counter1&"  counter2(reported): "&counter2&"")
		
		case else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK

	end select

end sub

OnProjectBrowserScript

dim p as EA.Package

'---------------------------------------------------------------------------------------------------------------------------
'                                                  main sub
'---------------------------------------------------------------------------------------------------------------------------
sub checkPackage(p)
	
	dim el as EA.Element
	dim elCon as EA.Connector
	
	dim elSource as EA.Element
	dim elTarget as EA.Element
		
	dim packageTarget as EA.Package
	dim packageSource as EA.Package

	dim pathString
	
	dim realStatus
	
	for each el In p.elements

'---------------------------------------------------------------------------------------------------------------------------
'                                           exclude CodeLists and Enumerations
'---------------------------------------------------------------------------------------------------------------------------	
		
	if Ucase(el.Stereotype) <> "CODELIST" AND Ucase(el.Stereotype) <> "ENUMERATION" AND el.Type <> "Enumeration" then
	
		if Ucase(el.Name) <> "KOMMUNENUMMER" then
						
			Session.Output(" ")
			Session.Output("ElementType: "&el.Type&" - «"&el.Stereotype&"» "&el.name&"")
			realStatus = 1
			
			for each elCon in el.Connectors
			
				if elCon.Type = "Realisation" then
					
					realStatus = 0
					set elSource = Repository.GetElementByID(elCon.SupplierID)
					set elTarget = Repository.GetElementByID(elCon.ClientID)
															
					pathString = ""
									
					set packageTarget = Repository.GetPackageByID(elTarget.PackageID)
					set packageSource = Repository.GetPackageByID(elSource.PackageID)
										
					'Doesn't list any realizations within the same package.. TBD?
					
					
					if p.PackageID <> elTarget.PackageID then		
						Do until packageTarget.ParentID = "0"

							if packageTarget.StereotypeEx <> "" then
								
								'dim packageTarget as EA.Package
								'names
								if not packageCollectionAs.Contains("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"") then
									packageCollectionAs.Add("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"")
								end if
								'guids
								if not packageCollectionAsGUID.Contains(packageTarget.PackageGUID) then
									packageCollectionAsGUID.Add(packageTarget.PackageGUID)
								end if
								
								pathString = "«" + packageTarget.StereotypeEx + "» " + packageTarget.Name + "/" + pathString
						
							else
						
								pathString = packageTarget.Name + "/" + pathString
								
							end if
						
							set packageTarget = Repository.GetPackageByID(packageTarget.ParentID)
						
						Loop
					
						if Trim(Ucase(el.Notes)) <> Trim(Ucase(elTarget.Notes)) then
						
							Session.Output("realized as: «"&elTarget.Stereotype&"» "&elTarget.Name&"  (CHECK DEFINITION)  @ "&pathString&"")
						
						else
						
							Session.Output("realized as: «"&elTarget.Stereotype&"» "&elTarget.Name&"  @ "&pathString&"")
									   
						end if
						
						Call attributeComparison(elTarget, elSource)
						
						Session.Output("Attributes realized: "&realizedCounter&" of "&totalAttributes&"   Not realized: "&notRealizedCounter&"  New or changed: "&newOrChangedCounter&"")
						Session.Output(" ")						
					end if
					
					
					if p.PackageID <> elSource.PackageID then
						Do until packageSource.ParentID = "0"
							
							if packageSource.StereotypeEx <> "" then
								
								'name
								if not packageCollectionOf.Contains("«"&packageSource.StereotypeEx&"»"&packageSource.Name&"") then
									packageCollectionOf.Add("«"&packageSource.StereotypeEx&"»"&packageSource.Name&"")
								end if
								'guids
								if not packageCollectionOfGUID.Contains(packageSource.PackageGUID) then
									packageCollectionOfGUID.Add(packageSource.PackageGUID)
								end if
								
								pathString = "«" + packageSource.StereotypeEx + "» " + packageSource.Name + "/" + pathString
							
							else
								
								pathString = packageSource.Name + "/" + pathString
						
							end if
							
							set packageSource = Repository.GetPackageByID(packageSource.ParentID)
						
						Loop
											
						if Trim(Ucase(el.Notes)) <> Trim(Ucase(elSource.Notes)) then
						
							Session.Output("realization of: «"&elSource.Stereotype&"» "&elSource.Name&"  (CHECK DEFINITION)  @ "&pathString&"")
						
						else
						
							Session.Output("realization of: «"&elSource.Stereotype&"» "&elSource.Name&"  @ "&pathString&"")
						
						end if
						
						Call attributeComparison(elTarget, elSource)
					
						Session.Output("Attributes realized: "&realizedCounter&" of "&totalAttributes&"   Not realized: "&notRealizedCounter&"  New or changed: "&newOrChangedCounter&"")
						'Session.Output(" ")
					end if
				end if
			next
			if realStatus = 1 then
				Session.Output("   not realized")
			end if
		end if				
	end if

'---------------------------------------------------------------------------------------------------------------------------
'                                         Only CodeLists and Enumerations
'---------------------------------------------------------------------------------------------------------------------------	
	
	if Ucase(el.Stereotype) = "CODELIST" or Ucase(el.Stereotype) = "ENUMERATION" or el.Type = "Enumeration" then
	
		if Ucase(el.Name) <> "KOMMUNENUMMER" then
						
			Session.Output(" ")
			Session.Output("ElementType: "&el.Type&" - «"&el.Stereotype&"» "&el.name&"")
			realStatus = 1
			
			for each elCon in el.Connectors
			
				if elCon.Type = "Realisation" then
					
					realStatus = 0
					set elSource = Repository.GetElementByID(elCon.SupplierID)
					set elTarget = Repository.GetElementByID(elCon.ClientID)
				
					pathString = ""
									
					set packageTarget = Repository.GetPackageByID(elTarget.PackageID)
					set packageSource = Repository.GetPackageByID(elSource.PackageID)
										
					'Doesn't list any realizations within the same package.. TBD?
					
					
					if p.PackageID <> elTarget.PackageID then		
						Do until packageTarget.ParentID = "0"

							if packageTarget.StereotypeEx <> "" then

								if not packageCollectionAs.Contains("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"") then
									packageCollectionAs.Add("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"")
								end if
								if not packageCollectionAsGUID.Contains(packageTarget.PackageGUID) then
									packageCollectionAsGUID.Add(packageTarget.PackageGUID)
								end if
								
								pathString = "«" + packageTarget.StereotypeEx + "» " + packageTarget.Name + "/" + pathString
						
							else
						
								pathString = packageTarget.Name + "/" + pathString
								
							end if
						
							set packageTarget = Repository.GetPackageByID(packageTarget.ParentID)
						
						Loop
					
						if Trim(Ucase(el.Notes)) <> Trim(Ucase(elTarget.Notes)) then
						
							Session.Output("realized as: «"&elTarget.Stereotype&"» "&elTarget.Name&"  (CHECK DEFINITION)  @ "&pathString&"")
						
						else
						
							Session.Output("realized as: «"&elTarget.Stereotype&"» "&elTarget.Name&"  @ "&pathString&"")
									   
						end if
						
						Call codeComparison(elTarget, elSource)
						
						Session.Output("Codes realized: "&realizedCounter&" of "&totalAttributes&"   Not realized: "&notRealizedCounter&"  New or changed: "&newOrChangedCounter&"")
						Session.Output(" ")						
					end if
					
					
					if p.PackageID <> elSource.PackageID then
						Do until packageSource.ParentID = "0"
							
							if packageSource.StereotypeEx <> "" then
								
								if not packageCollectionOf.Contains("«"&packageSource.StereotypeEx&"»"&packageSource.Name&"") then
									packageCollectionOf.Add("«"&packageSource.StereotypeEx&"»"&packageSource.Name&"")
								end if
								if not packageCollectionOfGUID.Contains(packageSource.PackageGUID) then
									packageCollectionOfGUID.Add(packageSource.PackageGUID)
								end if
								
								pathString = "«" + packageSource.StereotypeEx + "» " + packageSource.Name + "/" + pathString
							
							else
								
								pathString = packageSource.Name + "/" + pathString
						
							end if
							
							set packageSource = Repository.GetPackageByID(packageSource.ParentID)
						
						Loop
											
						if Trim(Ucase(el.Notes)) <> Trim(Ucase(elSource.Notes)) then
						
							Session.Output("realization of: «"&elSource.Stereotype&"» "&elSource.Name&"  (CHECK DEFINITION)  @ "&pathString&"")
						
						else
						
							Session.Output("realization of: «"&elSource.Stereotype&"» "&elSource.Name&"  @ "&pathString&"")
						
						end if
						
						Call codeComparison(elTarget, elSource)
					
						Session.Output("Codes realized: "&realizedCounter&" of "&totalAttributes&"   Not realized: "&notRealizedCounter&"  New or changed: "&newOrChangedCounter&"")
						Session.Output(" ")
					end if
				end if
			next
			if realStatus = 1 then
				Session.Output("   not realized")
			end if
		end if				
	end if
	
	next

	dim subP as EA.Package
	for each subP in p.packages
		checkPackage(subP)
	next

end sub
'---------------------------------------------------------------------------------------------------------------------------
'                                                  end main sub
'---------------------------------------------------------------------------------------------------------------------------

dim subEl as EA.Element
dim superEl as EA.Element

'---------------------------------------------------------------------------------------------------------------------------
'                                         sub for attribute comparison
'---------------------------------------------------------------------------------------------------------------------------
sub attributeComparison(subEl, superEl)

'Session.Output("DEBUG:  mainEl-"&mainEl.Name&"  otherEl-"&otherEl.Name&"")

dim superAtt as EA.Attribute
dim subAtt as EA.Attribute
	
dim attStatus
attStatus = 0
dim attStatus2
attStatus2 = 0
dim attStatus3
attStatus3 = 0

dim attRealList
set attRealList = CreateObject("System.Collections.ArrayList")
dim attNotRealList
set attNotRealList = CreateObject("System.Collections.ArrayList")
dim attNewList
set attNewList = CreateObject("System.Collections.ArrayList")

realizedCounter = 0
newOrChangedCounter = 0
notRealizedCounter = 0
totalAttributes = 0
		
'attributes on the element realized from
for each superAtt in superEl.AttributesEx
	attStatus = 0
	totalAttributes = totalAttributes + 1
	'attributes on the element that realizes
	for each subAtt in subEl.Attributes
		if Ucase(superAtt.Name) = Ucase(subAtt.Name) then
			if superAtt.LowerBound <> subAtt.LowerBound or superAtt.UpperBound <> subAtt.UpperBound then
				if Trim(Ucase(superAtt.Notes)) <> Trim(Ucase(subAtt.Notes)) then
					attRealList.add("   attribute realized: "&superAtt.Name&" ["&superAtt.LowerBound&".."&superAtt.UpperBound&"] as ["&subAtt.LowerBound&".."&subAtt.UpperBound&"]  (CHECK DEFINITION) (MULTIPLICITY CHANGED)")
				else
					attRealList.add("   attribute realized: "&superAtt.Name&" ["&superAtt.LowerBound&".."&superAtt.UpperBound&"] as ["&subAtt.LowerBound&".."&subAtt.UpperBound&"]  (MULTIPLICITY CHANGED)")
				end if
			elseif superAtt.LowerBound = subAtt.LowerBound and superAtt.UpperBound = subAtt.UpperBound then
				if Trim(Ucase(superAtt.Notes)) <> Trim(Ucase(subAtt.Notes)) then
					attRealList.add("   attribute realized: "&superAtt.Name&" ["&superAtt.LowerBound&".."&superAtt.UpperBound&"]  (CHECK DEFINITION)")
				else
					attRealList.add("   attribute realized: "&superAtt.Name&" ["&superAtt.LowerBound&".."&superAtt.UpperBound&"]")
				end if
			end if
			attStatus = 1
			realizedCounter = realizedCounter + 1
		end if
	next



	if attStatus = 0 then
		
		dim testName
'		Session.Output("DEBUG: "&superAtt.ParentID&"")
'		set testName = Repository.GetElementByID(superAtt.ParentID)
'		Session.Output("DEBUG name: "&testName.Name&"")
		
		if superAtt.LowerBound = 1 then
			attNotRealList.add("      attribute not realized: "&superAtt.Name&" ["&superAtt.LowerBound&".."&superAtt.UpperBound&"]  (THIS SHOULD HAVE BEEN REALIZED)")
			notRealizedCounter = notRealizedCounter + 1
		else
			attNotRealList.add("      attribute not realized: "&superAtt.Name&" ["&superAtt.LowerBound&".."&superAtt.UpperBound&"]")
			notRealizedCounter = notRealizedCounter + 1
		end if
	end if
next

if attRealList.Count > 0 then
	dim aX
	for each aX in attRealList
		Session.Output(aX)
	next
'else	
'	Session.Output("   attribute realized: 0")
end if

if AttNotRealList.Count > 0 then
	'Session.Output(" ")
	dim bX
	for each bX in attNotRealList
		Session.Output(bX)
	next
	'Session.Output(" ")
'else	
'	Session.Output(" ")
'	Session.Output("      attribute not realized: 0")
'	Session.Output(" ")
end if




'attributes on the element that realizes
for each subAtt in subEl.Attributes
	attStatus2 = 0
	'attributes on the element realized from
	for each superAtt in superEl.AttributesEx
		if Ucase(subAtt.Name) = Ucase(superAtt.Name) then
			attStatus2 = 1
		end if
	next
					
	if attStatus2 = 0 then
		attNewList.add("         new or changed attribute: "&subAtt.Name&" :"&subAtt.Type&" ["&subAtt.LowerBound&".."&subAtt.UpperBound&"]")
		newOrChangedCounter = newOrChangedCounter + 1
	end if
next



if attNewList.Count > 0 then
	dim cX
	for each cX in attNewList
		Session.Output(cX)
	next
'else	
'	Session.Output("         new or changed attribute: 0")
end if

end sub
'---------------------------------------------------------------------------------------------------------------------------
'                                                 end of sub for attribute comparison
'---------------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------------
'                                                 sub for codes
'---------------------------------------------------------------------------------------------------------------------------
sub codeComparison(subEl, superEl)

'Session.Output("DEBUG:  mainEl-"&mainEl.Name&"  otherEl-"&otherEl.Name&"")

dim superAtt as EA.Attribute
dim subAtt as EA.Attribute
	
dim attStatus
attStatus = 0
dim attStatus2
attStatus2 = 0
dim attStatus3
attStatus3 = 0

dim codeRealList
set codeRealList = CreateObject("System.Collections.ArrayList")
dim codeNotRealList
set codeNotRealList = CreateObject("System.Collections.ArrayList")
dim codeNewList
set codeNewList = CreateObject("System.Collections.ArrayList")

realizedCounter = 0
newOrChangedCounter = 0
notRealizedCounter = 0
totalAttributes = 0
		
'attributes on the element realized from
for each superAtt in superEl.AttributesEx
	attStatus = 0
	totalAttributes = totalAttributes + 1
	'attributes on the element that realizes
	for each subAtt in subEl.Attributes
		if Ucase(superAtt.Name) = Ucase(subAtt.Name) then
			if superAtt.LowerBound <> subAtt.LowerBound or superAtt.UpperBound <> subAtt.UpperBound then
				if Trim(Ucase(superAtt.Notes)) <> Trim(Ucase(subAtt.Notes)) then
					codeRealList.add("   code realized: "&superAtt.Name&"  (CHECK DEFINITION) (CODES SHOULD BE [1..1])")
				else
					codeRealList.add("   code realized: "&superAtt.Name&"  (CODES SHOULD BE [1..1])")
				end if
			elseif superAtt.LowerBound = subAtt.LowerBound and superAtt.UpperBound = subAtt.UpperBound then
				if Trim(Ucase(superAtt.Notes)) <> Trim(Ucase(subAtt.Notes)) then
					codeRealList.add("   code realized: "&superAtt.Name&"  (CHECK DEFINITION)")
				else
					codeRealList.add("   code realized: "&superAtt.Name&"")
				end if
			end if
			attStatus = 1
			realizedCounter = realizedCounter + 1
		end if
	next
					
	if attStatus = 0 then
		codeNotRealList.add("      code not realized: "&superAtt.Name&"")
		notRealizedCounter = notRealizedCounter + 1
	end if
next

if codeRealList.Count > 0 then
	dim aX
	for each aX in codeRealList
		Session.Output(aX)
	next
'else
'	Session.Output("   code realized: 0")
end if

if codeNotRealList.Count > 0 then
	'Session.Output(" ")
	dim bX
	for each bX in codeNotRealList
		Session.Output(bX)
	next
	'Session.Output(" ")
'else	
'	Session.Output(" ")
'	Session.Output("      code not realized: 0")
'	Session.Output(" ")
end if


'attributes on the element that realizes
for each subAtt in subEl.Attributes
	attStatus2 = 0
	'attributes on the element realized from
	for each superAtt in superEl.AttributesEx
		if Ucase(subAtt.Name) = Ucase(superAtt.Name) then
			attStatus2 = 1
		end if
	next
						
	if attStatus2 = 0 then
		codeNewList.add("         new or changed code: "&subAtt.Name&"")
		newOrChangedCounter = newOrChangedCounter + 1
	end if
next

if codeNewList.Count > 0 then
	dim cX
	for each cX in codeNewList
		Session.Output(cX)
	next
'else
'	Session.Output("         new or changed code: 0")
end if


end sub
'---------------------------------------------------------------------------------------------------------------------------
'                                                 end of sub for codes
'---------------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------------
'                                                 sub for reverse check
'---------------------------------------------------------------------------------------------------------------------------
sub letsGoReverse

dim reversePackage as EA.Package



if packageCollectionAsGUID.Count > 0 then
	dim eX2
	for each eX2 in packageCollectionAsGUID

		'session.output(eX2)
		set reversePackage = Repository.GetPackageByGuid(eX2)
		'Session.Output(reversePackage.Name)
		'Session.Output("Start check on package: [«"&thePackage.StereotypeEx&"» "&thePackage.Name&"] "&Now&"")
		Session.Output("Runnning check on: «"&reversePackage.StereotypeEx&"» "&reversePackage.Name&"")
'---------------------------------------------------------------------------------------------------------------------------
'                                                  main sub
'---------------------------------------------------------------------------------------------------------------------------
'packageCollectionAs.Add("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"")
	dim revElListNot
	set revElListNot = CreateObject("System.Collections.ArrayList")
	dim revElListIs
	set revElListIs = CreateObject("System.Collections.ArrayList")
	
	dim el as EA.Element
	dim elCon as EA.Connector
	
	dim elSource as EA.Element
	dim elTarget as EA.Element
		
	dim packageTarget as EA.Package
	dim packageSource as EA.Package

	dim pathString
	
	dim realStatus
	
	for each el In reversePackage.elements

'---------------------------------------------------------------------------------------------------------------------------
'                                           exclude CodeLists and Enumerations
'---------------------------------------------------------------------------------------------------------------------------	
		
	if Ucase(el.Stereotype) <> "CODELIST" AND Ucase(el.Stereotype) <> "ENUMERATION" AND el.Type <> "Enumeration" then
	
		if Ucase(el.Name) <> "KOMMUNENUMMER" then
						
			'Session.Output(" ")
			'Session.Output("ElementType: "&el.Type&" - «"&el.Stereotype&"» "&el.name&"")
			realStatus = 1
			
			for each elCon in el.Connectors
			
				if elCon.Type = "Realisation" then
					
					realStatus = 0
					set elSource = Repository.GetElementByID(elCon.SupplierID)
					set elTarget = Repository.GetElementByID(elCon.ClientID)
															
					pathString = ""
									
					set packageTarget = Repository.GetPackageByID(elTarget.PackageID)
					set packageSource = Repository.GetPackageByID(elSource.PackageID)
										
					'Doesn't list any realizations within the same package.. TBD?
					
					
					if reversePackage.PackageID <> elTarget.PackageID then		
						Do until packageTarget.ParentID = "0"

							if packageTarget.StereotypeEx <> "" then
								
								'dim packageTarget as EA.Package
								'names
								if not packageCollectionAs.Contains("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"") then
									packageCollectionAs.Add("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"")
								end if
								'guids
								if not packageCollectionAsGUID.Contains(packageTarget.PackageGUID) then
									packageCollectionAsGUID.Add(packageTarget.PackageGUID)
								end if
								
								pathString = "«" + packageTarget.StereotypeEx + "» " + packageTarget.Name + "/" + pathString
						
							else
						
								pathString = packageTarget.Name + "/" + pathString
								
							end if
						
							set packageTarget = Repository.GetPackageByID(packageTarget.ParentID)
						
						Loop
					
						if Trim(Ucase(el.Notes)) <> Trim(Ucase(elTarget.Notes)) then
						
							'Session.Output("realized as: «"&elTarget.Stereotype&"» "&elTarget.Name&"  (CHECK DEFINITION)  @ "&pathString&"")
							revElListIs.Add("«"&el.Stereotype&"»"&el.name&"")
						else
						
							'Session.Output("realized as: «"&elTarget.Stereotype&"» "&elTarget.Name&"  @ "&pathString&"")
							revElListIs.Add("«"&el.Stereotype&"»"&el.name&"")	   
						end if
						
						'Call attributeComparison(elTarget, elSource)
						
						'Session.Output("Attributes realized: "&realizedCounter&" of "&totalAttributes&"   Not realized: "&notRealizedCounter&"  New or changed: "&newOrChangedCounter&"")
						'Session.Output(" ")						
					end if
					
					
					if reversePackage.PackageID <> elSource.PackageID then
						Do until packageSource.ParentID = "0"
							
							if packageSource.StereotypeEx <> "" then
								
								'name
								if not packageCollectionOf.Contains("«"&packageSource.StereotypeEx&"»"&packageSource.Name&"") then
									packageCollectionOf.Add("«"&packageSource.StereotypeEx&"»"&packageSource.Name&"")
								end if
								'guids
								if not packageCollectionOfGUID.Contains(packageSource.PackageGUID) then
									packageCollectionOfGUID.Add(packageSource.PackageGUID)
								end if
								
								pathString = "«" + packageSource.StereotypeEx + "» " + packageSource.Name + "/" + pathString
							
							else
								
								pathString = packageSource.Name + "/" + pathString
						
							end if
							
							set packageSource = Repository.GetPackageByID(packageSource.ParentID)
						
						Loop
											
						if Trim(Ucase(el.Notes)) <> Trim(Ucase(elSource.Notes)) then
						
							'Session.Output("realization of: «"&elSource.Stereotype&"» "&elSource.Name&"  (CHECK DEFINITION)  @ "&pathString&"")
							revElListIs.Add("«"&el.Stereotype&"»"&el.name&"")
						else
						
							'Session.Output("realization of: «"&elSource.Stereotype&"» "&elSource.Name&"  @ "&pathString&"")
							revElListIs.Add("«"&el.Stereotype&"»"&el.name&"")
						end if
						
						'Call attributeComparison(elTarget, elSource)
					
						'Session.Output("Attributes realized: "&realizedCounter&" of "&totalAttributes&"   Not realized: "&notRealizedCounter&"  New or changed: "&newOrChangedCounter&"")
						'Session.Output(" ")
					end if
				end if
			next
			if realStatus = 1 then
				'Session.Output("   not realized")
				revElListNot.Add("«"&el.Stereotype&"»"&el.name&"")
				
			end if
		end if				
	end if

'---------------------------------------------------------------------------------------------------------------------------
'                                         Only CodeLists and Enumerations
'---------------------------------------------------------------------------------------------------------------------------	
	
	if Ucase(el.Stereotype) = "CODELIST" or Ucase(el.Stereotype) = "ENUMERATION" or el.Type = "Enumeration" then
	
		if Ucase(el.Name) <> "KOMMUNENUMMER" then
						
			'Session.Output(" ")
			'Session.Output("ElementType: "&el.Type&" - «"&el.Stereotype&"» "&el.name&"")
			realStatus = 1
			
			for each elCon in el.Connectors
			
				if elCon.Type = "Realisation" then
					
					realStatus = 0
					set elSource = Repository.GetElementByID(elCon.SupplierID)
					set elTarget = Repository.GetElementByID(elCon.ClientID)
				
					pathString = ""
									
					set packageTarget = Repository.GetPackageByID(elTarget.PackageID)
					set packageSource = Repository.GetPackageByID(elSource.PackageID)
										
					'Doesn't list any realizations within the same package.. TBD?
					
					
					if reversePackage.PackageID <> elTarget.PackageID then		
						Do until packageTarget.ParentID = "0"

							if packageTarget.StereotypeEx <> "" then

								if not packageCollectionAs.Contains("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"") then
									packageCollectionAs.Add("«"&packageTarget.StereotypeEx&"»"&packageTarget.Name&"")
								end if
								if not packageCollectionAsGUID.Contains(packageTarget.PackageGUID) then
									packageCollectionAsGUID.Add(packageTarget.PackageGUID)
								end if
								
								pathString = "«" + packageTarget.StereotypeEx + "» " + packageTarget.Name + "/" + pathString
						
							else
						
								pathString = packageTarget.Name + "/" + pathString
								
							end if
						
							set packageTarget = Repository.GetPackageByID(packageTarget.ParentID)
						
						Loop
					
						if Trim(Ucase(el.Notes)) <> Trim(Ucase(elTarget.Notes)) then
						
							'Session.Output("realized as: «"&elTarget.Stereotype&"» "&elTarget.Name&"  (CHECK DEFINITION)  @ "&pathString&"")
							revElListIs.Add("«"&el.Stereotype&"»"&el.name&"")
						else
						
							'Session.Output("realized as: «"&elTarget.Stereotype&"» "&elTarget.Name&"  @ "&pathString&"")
							revElListIs.Add("«"&el.Stereotype&"»"&el.name&"")		   
						end if
						
						'Call codeComparison(elTarget, elSource)
						
						'Session.Output("Codes realized: "&realizedCounter&" of "&totalAttributes&"   Not realized: "&notRealizedCounter&"  New or changed: "&newOrChangedCounter&"")
						'Session.Output(" ")						
					end if
					
					
					if reversePackage.PackageID <> elSource.PackageID then
						Do until packageSource.ParentID = "0"
							
							if packageSource.StereotypeEx <> "" then
								
								if not packageCollectionOf.Contains("«"&packageSource.StereotypeEx&"»"&packageSource.Name&"") then
									packageCollectionOf.Add("«"&packageSource.StereotypeEx&"»"&packageSource.Name&"")
								end if
								if not packageCollectionOfGUID.Contains(packageSource.PackageGUID) then
									packageCollectionOfGUID.Add(packageSource.PackageGUID)
								end if
								
								pathString = "«" + packageSource.StereotypeEx + "» " + packageSource.Name + "/" + pathString
							
							else
								
								pathString = packageSource.Name + "/" + pathString
						
							end if
							
							set packageSource = Repository.GetPackageByID(packageSource.ParentID)
						
						Loop
											
						if Trim(Ucase(el.Notes)) <> Trim(Ucase(elSource.Notes)) then
						
							'Session.Output("realization of: «"&elSource.Stereotype&"» "&elSource.Name&"  (CHECK DEFINITION)  @ "&pathString&"")
							revElListIs.Add("«"&el.Stereotype&"»"&el.name&"")
						else
						
							'Session.Output("realization of: «"&elSource.Stereotype&"» "&elSource.Name&"  @ "&pathString&"")
							revElListIs.Add("«"&el.Stereotype&"»"&el.name&"")
						end if
						
						'Call codeComparison(elTarget, elSource)
					
						'Session.Output("Codes realized: "&realizedCounter&" of "&totalAttributes&"   Not realized: "&notRealizedCounter&"  New or changed: "&newOrChangedCounter&"")
						'Session.Output(" ")
					end if
				end if
			next
			if realStatus = 1 then
				'Session.Output("   not realized")
				revElListNot.Add("«"&el.Stereotype&"»"&el.name&"")
			end if
		end if				
	end if
	
	next

	dim subP as EA.Package
	for each subP in reversePackage.packages
		letsGoReverse(subP)
	next


'---------------------------------------------------------------------------------------------------------------------------
'                                                  end main sub
'---------------------------------------------------------------------------------------------------------------------------		
dim yX
dim zX
for each yX in revElListIs
	Session.Output("realized: "&yX&"")
next
for each zX in revElListNot
	Session.Output("not realized: "&zX&"")
next

Session.Output(" ")

next
end if

end sub



















'---------------------------------------------------------------------------------------------------------------------------
'                                                 end of sub for reverse check
'---------------------------------------------------------------------------------------------------------------------------
