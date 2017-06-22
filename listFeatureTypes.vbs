option explicit

!INC Local Scripts.EAConstants-VBScript

' skriptnavn:         listFeatureTypes


sub listFeatureTypesForEnValgtPakke()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
	Repository.CreateOutputTab "Error"
	Repository.ClearOutput "Error"

	'Get the currently selected CodeList in the tree to work on

	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()

  
	'Repository.GetTreeSelectedItemType(

	if not theElement is nothing  then
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		if Repository.GetTreeSelectedItemType() = otPackage then
			'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
					dim message
			dim box
			box = Msgbox ("Start listing of FeatureTypes for package : [" & theElement.Name & "]. ",1)
			select case box
			case vbOK
				dim namespace
				'namespace = getTaggedValue(theElement,"codeList")
				'if namespace = "" then
				'	namespace = InputBox("Please select the codespace name.", "namespace", "http://skjema.geonorge.no/SOSI/produktspesifikasjon/Stedsnavn/5.0/"&theElement.Name)
				'end if
				call listFeatureTypes(theElement,namespace)
			case VBcancel

			end select
	

		Else
		  'Other than CodeList selected in the tree
		  MsgBox( "This script requires a package to be selected in the Project Browser." & vbCrLf & _
			"Please select a package in the Project Browser and try once more." )
		end If
		'Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
		Repository.EnsureOutputVisible "Script"
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires a package to be selected in the Project Browser." & vbCrLf & _
	  "Please select a package in the Project Browser and try again." )
	end if
end sub


sub listFeatureTypes(pkg,namespace)
	dim presentasjonsnavn
	
 	dim elements as EA.Collection 
 	set elements = pkg.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages) 
	if UCase(pkg.Element.Stereotype) = "APPLICATIONSCHEMA" then
		namespace = pkg.Name
	end if
	
	dim i 
	for i = 0 to elements.Count - 1 
		dim currentElement as EA.Element 
		set currentElement = elements.GetAt( i ) 
				
		if currentElement.Type = "Class" and UCase(currentElement.Stereotype) = "FEATURETYPE" then
			
			Repository.WriteOutput "Script", namespace&";"&pkg.Name&";"&currentElement.Name&";"&getDefinitionText(currentElement),0
			
			
		end if
	
	next
	dim subP as EA.Package
	for each subP in pkg.packages
	    call listFeatureTypes(subP,namespace)
	next


end sub

function getDefinitionText(currentElement)

    Dim txt, res, tegn, i, u
    u=0
	getDefinitionText = ""
		txt = Trim(currentElement.Notes)
		res = ""
		' loop gjennom alle tegn
		For i = 1 To Len(txt)
		  tegn = Mid(txt,i,1)
		  If tegn = ";" Then
			  res = res + " "
		  Else 
			If tegn = """" Then
			  res = res + "'"
			Else
			  If tegn < " " Then
			    res = res + " "
			  Else
			    res = res + Mid(txt,i,1)
			  End If
			End If
		  End If
		  
		Next
		
	getDefinitionText = res

end function

listFeatureTypesForEnValgtPakke
