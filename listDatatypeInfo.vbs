option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 	listDatatypeInfo
' Author: 		Kent Jonsrud
' Purpose: 		List in which file an Application Schema package is managed by Package Control
' Date: 		2018-12-27
' Date: 		2023-10-27 added stereotype output on attributes and list of tagged values on elements and attributes
'				TBD: "navigable" roles
'				TBD: packages


' Project Browser Script main function
sub OnProjectBrowserScript()

	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()

	'find out what type is selected
	select case treeSelectedType

		case otElement
'			' Code for when an element is selected
			dim theElement as EA.Element
			set theElement = Repository.GetTreeSelectedObject()
			
			dim box, mess
			mess = 	"listDatatypeInfo. Script version 2023-10-27." & vbCrLf
			mess = mess + "NOTE! This list info on element: "& vbCrLf & "[«" & theElement.Stereotype & "» " & theElement.Name & "]."

			box = Msgbox (mess, vbOKCancel)
			select case box
			case vbOK
				'Repository.ClearOutput "Script"
				elementDatatypeInfo(theElement)
				Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
				Repository.EnsureOutputVisible "Script"
			case VBcancel
						
			end select 

'
'		case otPackage
			' Code for when a package is selected
'			dim thePackage as EA.Package
'			set thePackage = Repository.GetTreeSelectedObject()
'
'
'		case otDiagram
'			' Code for when a diagram is selected
'			dim theDiagram as EA.Diagram
'			set theDiagram = Repository.GetTreeSelectedObject()
'
		case otAttribute
			' Code for when an attribute is selected
			dim theAttribute as EA.Attribute
			set theAttribute = Repository.GetTreeSelectedObject()

			'dim thePackage as EA.Package
			'set thePackage = Repository.GetTreeSelectedObject()
			'Msgbox "The selected package is: [" & thePackage.Name &"]. Starting search for XMI-file name."
			dim box2, mess2
			mess2 = 	"listDatatypeInfo. Script version 2023-10-27." & vbCrLf
			mess2 = mess2 + "NOTE! This list info on attribute: "& vbCrLf & "[«" & theAttribute.Stereotype & "» " & theAttribute.Name & "]."

			box2 = Msgbox (mess2, vbOKCancel)
			select case box2
			case vbOK
				Repository.ClearOutput "Script"
				attrDatatypeInfo(theAttribute)
				Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
				Repository.EnsureOutputVisible "Script"
			case VBcancel
						
			end select 


'		case otMethod
'			' Code for when a method is selected
'			dim theMethod as EA.Method
'			set theMethod = Repository.GetTreeSelectedObject()

		case else
			' Error message
			Session.Prompt "This script does not support items of this type. Please choose a package in order to start the script.", promptOK

	end select

end sub

sub elementDatatypeInfo(element)
			dim xelement as EA.Element
	'Session.Output("parentID: [" & attr.ParentID & "] visibility: [" & attr.Visibility & "] type: [" & attr.Type & "] default: [" & attr.Default & "] ClassifierID: [" & attr.ClassifierID & "] position: [" & attr.Pos & "]")
	dim attr as EA.Attribute
	dim taggedValue as EA.TaggedValue
	dim i
	Session.Output("datatype: [«" & element.Stereotype & "» " & element.Name & "] Notes: [" & element.Notes & "] parentId: [" & element.ParentID & "] Id: [" & element.ElementID & "]")
	set xelement = Repository.GetElementByID(element.ElementID)
	for i = 0 to xelement.TaggedValues.Count - 1
		set taggedValue = xelement.TaggedValues.GetAt(i)
		'	if taggedValue.Value <> "" then
				Session.Output("datatype: [«" & xelement.Stereotype & "» " & xelement.Name & "] tag: " & i & " [" & taggedValue.Name & "] value: [" & taggedValue.Value & "]")
		'	end if
	next
	for each attr in element.Attributes
		attrDatatypeInfo(attr)

if false then
		if attr.ClassifierID <> 0 then
			' TODO may need to test whether this ClassifierID points to a real object first
			dim datatype as EA.Element
			set datatype = Repository.GetElementByID(attr.ClassifierID)
			Session.Output("attribute: [«" & attr.Stereotype &  "» " & attr.Name & "] datatype: [«" & datatype.Stereotype & "» " & datatype.Name & "] type: [" & datatype.Type & "] in package: [" & Repository.GetPackageByID(datatype.PackageID).Name & "]  Created: [" & datatype.Created & "] - modified: [" & datatype.Modified & "]")

			'Session.Output("The package: " & package.Name & " Is version controlled: " & package.IsVersionControlled & " - owner: " & package.Owner & " - UML version: " & package.UMLVersion & " - version: " & package.Version)
			'Session.Output("Created: " & package.Created & " - modified: " & package.Modified & " - last saved: " & package.LastSaveDate & " - last read: " & package.LastLoadDate)
			'Session.Output("Tree position: " & package.TreePos & " - last read: " & package.LastLoadDate)
			'Session.Output("Flags: " & package.Flags & " - notes: " & package.Notes)
			'Session.Output(" ")
		else
			Session.Output("attribute: [" & attr.Name & "] datatype: [" & attr.Type & "] parentID: [" & attr.ParentID & "] visibility: [" & attr.Visibility & "] type: [" & attr.Type & "] default: [" & attr.Default & "] ClassifierID: [" & attr.ClassifierID & "] position: [" & attr.Pos & "] is not connected to any datatype class.")
		
		end if
end if 'false

	next
end sub

sub attrDatatypeInfo(attr)
	dim taggedValue as EA.TaggedValue
	dim i

	Session.Output("parent: [" & Repository.GetElementByID(attr.ParentID).Name & "] parentID: [" & attr.ParentID & "] visibility: [" & attr.Visibility & "] type: [" & attr.Type & "] default: [" & attr.Default & "] ClassifierID: [" & attr.ClassifierID & "] position: [" & attr.Pos & "]")

	for i = 0 to attr.TaggedValues.Count - 1			
		set taggedValue = attr.TaggedValues.GetAt(i)
		'	if taggedValue.Value <> "" then
			Session.Output("attribute: [«" & attr.Stereotype &  "» " & attr.Name & "] tag: [" & taggedValue.Name & "] value: [" & taggedValue.Value & "] number: [" & i & "]")
		'	end if
	next

	if attr.ClassifierID <> 0 then
		dim datatype as EA.Element
		set datatype = Repository.GetElementByID(attr.ClassifierID)
		Session.Output("datatype: [«" & datatype.Stereotype & "» " & datatype.Name & "] type: [" & datatype.Type & "] in package: [" & Repository.GetElementByID(datatype.PackageID).Name & "]  Created: [" & datatype.Created & "] - modified: [" & datatype.Modified & "]")



		'Session.Output("The package: " & package.Name & " Is version controlled: " & package.IsVersionControlled & " - owner: " & package.Owner & " - UML version: " & package.UMLVersion & " - version: " & package.Version)
		'Session.Output("Created: " & package.Created & " - modified: " & package.Modified & " - last saved: " & package.LastSaveDate & " - last read: " & package.LastLoadDate)
		'Session.Output("Tree position: " & package.TreePos & " - last read: " & package.LastLoadDate)
		'Session.Output("Flags: " & package.Flags & " - notes: " & package.Notes)
		'Session.Output(" ")
	else
		Session.Output("attribute: [" & attr.Name & "] is not connected to any datatype class.")
	
	end if

end sub



'start the main function
OnProjectBrowserScript
