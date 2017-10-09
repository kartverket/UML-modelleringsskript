option explicit

!INC Local Scripts.EAConstants-VBScript

' skriptnavn:         slettUtoverrettaRealiseringerFraKlasse
' beskrivelse         Sletter alle utoverretta realiseringer fra en valgt klasse.

sub fixClass(el)
	Repository.WriteOutput "Script", Now & " «" & el.Stereotype & "» " & el.Name &  " ElementID: " & el.ElementID, 0

	dim conn as EA.Connector
	dim i
	i = -1
	for each conn in el.Connectors
	i = i + 1
		'Repository.WriteOutput "Script", " Connection Nbr: " & i & " ConnectorID:" & conn.ConnectorID & " Type:" & conn.Type & " Connected to: " & Repository.GetElementByID(conn.ClientID).Name, 0
		'Repository.WriteOutput "Script", " Client: " & Repository.GetElementByID(conn.ClientID).Name & " Supplier: " & Repository.GetElementByID(conn.SupplierID).Name, 0
		'Repository.WriteOutput "Script", " ElementID: " & el.ElementID & " SupplierID:" & conn.SupplierID & " ClientID:" & conn.ClientID, 0
		if conn.Type = "Realisation" and conn.SupplierID = el.ElementID then

		    Repository.WriteOutput "Script", " Delete Outwards Realization Connector at [" & i & "] Connected to [" & Repository.GetElementByID(conn.ClientID).Name & "]", 0
			el.Connectors.Delete(i)
			
		else
		    'Repository.WriteOutput "Script", " Leave Connector at : " & i & " Connection between supplier/source: " & Repository.GetElementByID(conn.SupplierID).Name & " and client/target:" & Repository.GetElementByID(conn.ClientID).Name, 0	
		end if
	next
	el.Connectors.Refresh()

'	dim attr as EA.Attribute
'	for each attr in el.Attributes
'		Repository.WriteOutput "Script", Now & " " & el.Name & "." & attr.Name, 0
'
'		call slettEttellerannet(attr,prefiks)
'
'	next

end sub



sub oppdaterValgtKlasse()
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
			'mess = 	"Removes all outwards directed Realizations. Script version 2017-05-26." & vbCrLf
			mess = 	"Fjerner alle utoverretta realiseringer fra en valgt klasse." & vbCrLf
			mess = mess + "Merknad: Oppryddingsskript som er laget kun for å kjøres på klasser med fellesegenskaper i SOSI-produktspesifikasjoner." & vbCrLf
			mess = mess + "Advarsel! Dette skriptet vil kunne endre på innholdet i valgt element: "& vbCrLf & "[«" & theElement.Stereotype & "» " & theElement.Name & "]."

			box = Msgbox (mess, vbOKCancel,"Script slettUtoverrettaRealiseringerFraKlasse 2017-05-26.")
			select case box
			case vbOK
				if theElement.Type="Class" then
					'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
					call fixClass(theElement)
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

oppdaterValgtKlasse
