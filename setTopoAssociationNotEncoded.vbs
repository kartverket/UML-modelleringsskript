option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Diagram Script template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.
'
' Script Name: set topo association notEncoded
' Author: Magnus Karge
' Purpose: for topo-assosiasjoner. sette inn tagged values notEncoded på alle ender med roller og navigerbarhet.
' Date: 14.08.2015
'

'
' Diagram Script main function
'
sub OnDiagramScript()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()

	if not currentDiagram is nothing then
		' Get a reference to any selected connector
		dim selectedConnector as EA.Connector
		set selectedConnector = currentDiagram.SelectedConnector
		
		if not selectedConnector is nothing then
			' A connector is selected
			dim sourceEnd
			dim targetEnd
			set sourceEnd = selectedConnector.ClientEnd
			set targetEnd = selectedConnector.SupplierEnd
			 
			'Msgbox "selected element = " & selectedConnector.Name
			
			if selectedConnector.Stereotype = "topo" OR selectedConnector.Stereotype = "Topo" then 
				'call sub for targetEnd
				if targetEnd.IsNavigable AND Len(targetEnd.Role) > 0 then 
					Call TVSetElementTaggedValue( targetEnd, "xsdEncodingRule", "notEncoded", true )
					
				else
					Msgbox "Either role or navigability is not set on target end. No tagged value assigned on target end."
				end if
			
				'call sub for sourceEnd
				if sourceEnd.IsNavigable AND Len(sourceEnd.Role) > 0 then 
					Call TVSetElementTaggedValue( sourceEnd, "xsdEncodingRule", "notEncoded", true )
				else
					Msgbox "Either role or navigability is not set on source end. No tagged value assigned on source end." 
				end if
			else 
				Msgbox "The selected connector has not the required stereotype. Please select a connector with stereotype <<topo>>. "
			end if
			
		else
			' No connector is selected
			Msgbox "There is no connector selected, please start the script by right clicking on a connector." 
		end if
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if

end sub

' Sets the specified TaggedValue on the provided element. If the provided element does not already
' contain a TaggedValue with the specified name, a new TaggedValue is created with the requested
' name and value. If a TaggedValue already exists with the specified name then action to take is
' determined by the replaceExisting variable. If replaceExisting is set to true, the existing value
' is replaced with the specified value, if not, a new TaggedValue is created with the new value.
'
' @param[in] theElement (EA.Element) The element to set the TaggedValue value on
' @param[in] taggedValueName (String) The name of the TaggedValue to set
' @param[in] taggedValueValue (variant) The value of the TaggedValue to set
' @param[in] replaceExisting (boolean) If a TaggedValue of the same name already exists, specifies 
' whether to replace it, or create a new TaggedValue.
'
sub TVSetElementTaggedValue( theElement, taggedValueName, taggedValueValue, replaceExisting )

	if not theElement is nothing and Len(taggedValueName) > 0 then
		dim taggedValue as EA.TaggedValue
		set taggedValue = nothing
	
		' If replaceExisting was specified then attempt to get a tagged value from the element
		' with the provided name
		if not replaceExisting = false then
			
			set taggedValue = theElement.TaggedValues.GetByName( taggedValueName )
		end if
		
		if taggedValue is nothing then
			set taggedValue = theElement.TaggedValues.AddNew( taggedValueName, taggedValueValue )
		else
			set taggedValue.Value = taggedValueValue
		end if
		
		taggedValue.Update()
	end if

end sub

OnDiagramScript
