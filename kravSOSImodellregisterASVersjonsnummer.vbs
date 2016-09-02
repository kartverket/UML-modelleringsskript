option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: checkEndingOfPackageName
' Author: Sara Henriksen	
' Purpose: check if the package name ends with a version number. The version number could be a date or a serial number. Returns an error if the version 
' number contains anything other than 0-2 dots or numbers. 
' Date: 25.08.16
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
'			
			findASpackage(thePackage)
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


'sub procedure to check if the package name ends with a version number  
'@param[in]: package (Package Class)
sub checkEndingOfPackageName(thePackage)

	'find the last part of the package name, after "-" 
	dim startContent, endContent, stringContent 
	
	startContent = InStr(thePackage.Name, "-") 
	endContent = len(thePackage.Name)
	stringContent = mid(thePackage.Name, startContent+1, endContent) 

	
	dim versionNumberInPackageName
	versionNumberInPackageName = false 
	
	'count number of dots, only allowed to use max two. 
	dim dotCounter
	dotCounter = 0

	
	'check that the package name contains a "-", and thats it is just number(s) and "." after. 
	if InStr(thePackage.Name, "-") then 

			
'		'if the string is numeric or it has dots, set the valueOk true 
		if  InStr(stringContent, ".")  or IsNumeric(stringContent)  then
				
				versionNumberInPackageName = true 
			
				dim i, tegn 
				for i = 1 to len(stringContent) 
					tegn = Mid(stringContent,i,1)
					if tegn = "." then
						dotCounter = dotCounter  + 1 
					end if 
				next 
				'count number of dots. If it's more than 2 return an error. 
				if dotCounter < 3 then 
					versionNumberInPackageName = true
				else 
					'Session.Output("for mange punktum")
					versionNumberInPackageName = false
				end if
			
		 
		end if 
	
	end if 
	
	'check the string for letters and symbols. If the package name contains one of the following, then return an error. 
	if inStr(UCase(stringContent), "A") or inStr(UCase(stringContent), "B") or inStr(UCase(stringContent), "C") or inStr(UCase(stringContent), "D") or inStr(UCase(stringContent), "E") or inStr(UCase(stringContent), "F") or inStr(UCase(stringContent), "G") or inStr(UCase(stringContent), "H") or inStr(UCase(stringContent), "I") or inStr(UCase(stringContent), "J") or inStr(UCase(stringContent), "K") or inStr(UCase(stringContent), "L")  then 
		versionNumberInPackageName = false
	end if 
	
	if inStr(UCase(stringContent), "M") or inStr(UCase(stringContent), "N") or inStr(UCase(stringContent), "O") or inStr(UCase(stringContent), "P") or inStr(UCase(stringContent), "Q") or inStr(UCase(stringContent), "R") or inStr(UCase(stringContent), "S") or inStr(UCase(stringContent), "T") or inStr(UCase(stringContent), "U") or inStr(UCase(stringContent), "V") or inStr(UCase(stringContent), "W") or inStr(UCase(stringContent), "X") then          
		versionNumberInPackageName = false
	end if 
	
	if inStr(UCase(stringContent), "Y") or inStr(UCase(stringContent), "Z") or inStr(UCase(stringContent), "Æ") or inStr(UCase(stringContent), "Ø") or inStr(UCase(stringContent), "Å") then 
		versionNumberInPackageName = false
	end if 
	
	if inStr(stringContent, ",") or inStr(stringContent, "!") or inStr(stringContent, "@") or inStr(stringContent, "%") or inStr(stringContent, "&") or inStr(stringContent, """") or inStr(stringContent, "#") or inStr(stringContent, "$") or inStr(stringContent, "'") or inStr(stringContent, "(") or inStr(stringContent, ")") or inStr(stringContent, "*") or inStr(stringContent, "+") or inStr(stringContent, "/") then        
		versionNumberInPackageName = false
	end if
	
	if inStr(stringContent, ":") or inStr(stringContent, ";") or inStr(stringContent, ">") or inStr(stringContent, "<") or inStr(stringContent, "=") then
		versionNumberInPackageName = false
	end if 
	
	
	
	if versionNumberInPackageName = false  then  
		Session.Output("Error: Package ["&thePackage.Name&"] does not have a name ending with a version number. [/krav/SOSI-modellregister/applikasjonsskjema/versjonsnummer]")
	else 
		'Session.Output("OK")
	end if 
	

end sub 



'sub procedure that find the package and calls another sub procedure (checkNameOfPackage) 
'@param[in]: package (Package Class)
sub findASpackage(package)

	dim packages as EA.Collection
	set packages = package.Packages
	
	if UCase(package.Element.Stereotype) = "APPLICATIONSCHEMA" then 
		
		Call checkEndingOfPackageName(package.Element ) 
		
	end if 

end sub
OnProjectBrowserScript
