option explicit

!INC Local Scripts.EAConstants-VBScript

' skriptnavn:         listAllConnectorsOnClass

sub listClass(el)
	Repository.WriteOutput "Script", Now & " «" & el.Stereotype & "» " & el.Name &  " ElementID: " & el.ElementID, 0

	dim conn as EA.Connector
	dim i
	i = -1
	for each conn in el.Connectors
		i = i + 1
		Repository.WriteOutput "Script", " Connector at [" & i & "] Type [" & conn.Type & "] Connection between client/target/dratt fra [" & Repository.GetPackageByID(Repository.GetElementByID(conn.ClientID).PackageID).Name & "->" & Repository.GetElementByID(conn.ClientID).Name & "] and supplier/source/dratt til [" & Repository.GetPackageByID(Repository.GetElementByID(conn.SupplierID).PackageID).Name & "->" & Repository.GetElementByID(conn.SupplierID).Name & "]", 0	
	next

end sub






sub listValgtKlasse()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
	Repository.CreateOutputTab "Error"
	Repository.ClearOutput "Error"

	'Get the currently selected class in the tree to work on

	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()

	if not theElement is nothing  then
		if (theElement.ObjectType = otElement) then
			dim box, mess
			mess = 	"Lists all connectors on a class. Script version 2017-05-26." & vbCrLf
			mess = mess + "Lists the content of class: "& vbCrLf & "[«" & theElement.Stereotype & "» " & theElement.Name & "]."

			box = Msgbox (mess, vbOKCancel)
			select case box
			case vbOK
				if theElement.Type="Class" then
					'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
					call listClass(theElement)
				Else
					'Other than CodeList selected in the tree
					MsgBox( "This script requires a class to be selected in the Project Browser." & vbCrLf & _
					"Please select a class in the Project Browser and try once more." )
				end If
				Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
				Repository.EnsureOutputVisible "Script"
			case VBcancel
						
			end select 
		else
		  'Other than class selected in the tree
		  MsgBox( "This script requires a class to be selected in the Project Browser." & vbCrLf & _
			"Please select a class in the Project Browser and try once more." )
		end If
	else
		'No class selected in the tree
		MsgBox( "This script requires a class to be selected in the Project Browser." & vbCrLf & _
	  "Please select a class in the Project Browser and try again." )
	end if
end sub

listValgtKlasse
