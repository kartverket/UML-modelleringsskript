option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 	leggInnSOSIformat50Tagger (AddMissingTags)
' Author: 		Magnus Karge / Kent Jonsrud
' Purpose: 		To add missing tags on model elements
'               Kun for SOSI-format versjon 5.0
'               (application schemas, feature types & attributes, data types & attributes,code lists, enumerations)
'				for genetrating GML-ApplicationSchema as defined in the Norwegian standard "SOSI regler for UML-modellering"
' Date: 		2016-08-24     Original:11.09.2015   + Moddet av Kent 2016-03-09/08-24: Legger nå inn forslag til verdi i alle taggene!
' Date: 		2016-11-30     Tilpasset forslag til SOSI 5.0 regler:
' Date: 		2018-02-22     Tilpasset vedtatte SOSI 5.0 regler, henter SOSI_navn rett inn fra datatypen dersom denne finnes.
dim debug
debug = false
'set debug = true to get more information during execution
'set debug = false to get less information during execution

' Project Browser Script main function
sub OnProjectBrowserScript()

	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()

	'find out what type is selected
	select case treeSelectedType

'		case otElement
'			' Code for when an element is selected
'			dim theElement as EA.Element
'			set theElement = Repository.GetTreeSelectedObject()
'
		case otPackage
			' Code for when a package is selected
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			'Msgbox "The selected package is: [" & thePackage.Name &"]. Starting search for elements with missing tags."
			dim box, mess
			'mess = 	"Script: leggInnSOSIformat50Tagger" & vbCrLf
			mess =    	  "Generates tags needed for creating SOSI format version 5.0 from model elements." & vbCrLf
			mess = mess + "This script should NOT be run before correcting all conseptual errors found by the script SOSI model validation! "& vbCrLf
			mess = mess + "NOTE! Please be shure to have a backup as this script will add the missing tagged values to many element types in the package: "& vbCrLf & "[«" & thePackage.element.Stereotype & "» " & thePackage.Name & "]."

			box = Msgbox (mess, vbOKCancel,"SOSI 5.0 Script: leggInnSOSIformat50Tagger version: 2018-02-22")
			select case box
			case vbOK
				if LCase(thePackage.element.Stereotype) = "applicationschema" then
					Repository.ClearOutput "Script"
					FindElementsWithMissingTagsInPackage(thePackage)
				Else
					'Other than package selected in the tree
					MsgBox( "This script requires a package with stereotype «ApplicationSchema» to be selected in the Project Browser." & vbCrLf & _
					"Please select this and try once more." )
				end If
				Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
				Repository.EnsureOutputVisible "Script"
			case VBcancel

			end select


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
			Session.Prompt "This script does not support items of this type. Please choose a package in order to start the script.", promptOK

	end select

end sub

'sub procedure to check the content of a given package and all its subpackages and add missing tags to elements
'@param[in]: package (EA.package) The package containing elements with potentially missing tags.
			Dim ASpackage
sub FindElementsWithMissingTagsInPackage(package)

			if debug then Session.Output("The current package is: [" & package.Name & "]")
			'if the current package has stereotype applicationSchema then check tagged values
			if LCase(package.element.stereotype) = "applicationschema" then
				' Kapittel13kravSOSI: language, version, targetNamespace, SOSI_kortnavn, SOSI_modellstatus, SOSI_versjon
				' Kapittel13kravGML: language, version, targetNamespace, xmlns, xsdDocument, SOSI_modellstatus
				' Kapittel13kravAlle: designation og definition for engelsk
				ASpackage = "http://skjema.geonorge.no/SOSI/produktspesifikasjon/" + toNCName(package.Name,"/-DraftName")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "SOSI_kortnavn",toNCName(package.Name,"-"))
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "SOSI_modellstatus","utkastOgSkjult")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "SOSI_versjon","5.0")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "language","no")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "version","0.1")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "targetNamespace",ASpackage)
				'Call TVSetElementTaggedValue("ApplicationSchema",package.element, "xmlns","app")
				'Call TVSetElementTaggedValue("ApplicationSchema",package.element, "xsdDocument",toNCName(package.Name, "-") & "-DraftName.xsd")
				'Call TVSetElementTaggedValue("ApplicationSchema",package.element, "definition","""""@en")
				' TODO: Klipp inn engelsk fra notefeltet dersom du finner --Definition --
				'Call TVSetElementTaggedValue("ApplicationSchema",package.element, "designation","""""@en")
				' TODO: Klipp inn det engelske navnet fra Alias-feltet dersom dette finnes
				' Er denne ogs� praktisk � ha med her n�?
				'Call TVSetElementTaggedValue("ApplicationSchema",package.element, "xsdEncodingRule","sosi")
				'
				'TODO: sette korrekt case p� pakkestereotypen?
				'TODO: package.element.stereotype = "ApplicationSchema"
				'TODO: package.element.stereotype.Update()
				'
			end if

			dim packages as EA.Collection
			'collection of packages that belong to this package
			set packages = package.Packages

			'navigate the package collection and call the FindElementsWithMissingTagsInPackage
			'sub procedure for each of them
			dim packageCounter
			for packageCounter = 0 to packages.Count - 1
				dim currentPackage as EA.Package
				set currentPackage = packages.GetAt( packageCounter )
				FindElementsWithMissingTagsInPackage(currentPackage)
			next

			dim elements as EA.Collection
			'collection of elements that belong to this package (classes, notes... BUT NO packages)
			set elements = package.Elements

			'navigate the elements collection
			dim elementsCounter
			for elementsCounter = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( elementsCounter )

				if debug then Session.Output("The current element is: [" & currentElement.Name & "] Stereotype: [" & currentElement.Stereotype & "] Type: [" & currentElement.Type & "]")

				if LCase(currentElement.Stereotype) = "featuretype" then
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
					'Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
					'Call TVSetElementTaggedValue(currentElement, "isCollection")
					' F�lgende er ikke p�krevet!
					'Call TVSetElementTaggedValue("FeatureType", currentElement, "byValuePropertyType", "false")
					'Call TVSetElementTaggedValue("FeatureType", currentElement, "noPropertyType", "true")
				end if

				if LCase(currentElement.Stereotype) = "codelist" then
					'Call TVSetElementTaggedValue("CodeList", currentElement, "SOSI_navn", UCASE(currentElement.Name))
					'Call TVSetElementTaggedValue(currentElement, "SOSI_datatype")
					'Call TVSetElementTaggedValue("CodeList", currentElement, "SOSI_lengde", "")
					' F�lgende er ikke p�krevet!
					Call TVSetElementTaggedValue("CodeList", currentElement, "asDictionary", "false")
					if ASpackage <> "" then
						Call TVSetElementTaggedValue("CodeList", currentElement, "codeList", ASpackage + "/" + currentElement.Name + "-DraftName")
					end if
				end if

				if LCase(currentElement.Stereotype) = "datatype" then
					'Call TVSetElementTaggedValue("dataType", currentElement, "SOSI_navn", UCASE(currentElement.Name))
				end if

				if LCase(currentElement.Stereotype) = "enumeration" Or currentElement.Type = "Enumeration" then
					'Call TVSetElementTaggedValue("enumeration", currentElement, "SOSI_navn", UCASE(currentElement.Name))
				end if

				'if the currentElement has stereotype dataType, Union or FeatureType then navigate the attributes and check for missing tags
				if LCase(currentElement.Stereotype) = "featuretype" or LCase(currentElement.Stereotype) = "datatype" or LCase(currentElement.Stereotype) = "union" or currentElement.Type = "DataType" then
					dim attributesCounter
					for attributesCounter = 0 to currentElement.Attributes.Count - 1
						dim currentAttribute as EA.Attribute
						set currentAttribute = currentElement.Attributes.GetAt ( attributesCounter )
						'Session.Output( "  The current attribute is ["& currentAttribute.Name &"]" & "["& currentAttribute.Type &"]")
						if LCase(currentAttribute.Type) = "integer" or LCase(currentAttribute.Type) = "real" or LCase(currentAttribute.Type) = "boolean" or LCase(currentAttribute.Type) = "characterstring" or LCase(currentAttribute.Type) = "datetime" or LCase(currentAttribute.Type) = "date" then
							Call TVSetElementTaggedValue("["+currentElement.Name+"] attribute", currentAttribute, "SOSI_navn", currentAttribute.Name)
							'Call TVSetElementTaggedValue("attribute", currentAttribute, "SOSI_datatype", "T")
							'Call TVSetElementTaggedValue("attribute", currentAttribute, "SOSI_lengde", "")
						else
							if LCase(currentAttribute.Type) = "punkt" or LCase(currentAttribute.Type) = "kurve" or LCase(currentAttribute.Type) = "flate" or LCase(currentAttribute.Type) = "sverm" then
								Call TVSetElementTaggedValue("SOSI geometritype", currentElement, "SOSI_geometri", UCase(currentAttribute.Type))
							else
							if LCase(currentAttribute.Type) = "gm_point" then
								Call TVSetElementTaggedValue("ISO geometritype", currentElement, "SOSI_geometri", "PUNKT")
							else
							if LCase(currentAttribute.Type) = "gm_curve" then
								Call TVSetElementTaggedValue("ISO geometritype", currentElement, "SOSI_geometri", "KURVE")
							else
							if LCase(currentAttribute.Type) = "gm_surface" then
								Call TVSetElementTaggedValue("ISO geometritype", currentElement, "SOSI_geometri", "FLATE")
							else
							if LCase(currentAttribute.Type) = "gm_multipoint" then
								Call TVSetElementTaggedValue("ISO geometritype", currentElement, "SOSI_geometri", "SVERM")
							else
							if LCase(currentAttribute.Type) = "gm_object" then
								Call TVSetElementTaggedValue("ISO geometritype", currentElement, "SOSI_geometri", "PUNKT;KURVE;FLATE;SVERM")
							else
								'not a predefined name, must try to fetch the SOSI name from the old SOSI_navn of the datatype class
								if currentAttribute.ClassifierID then
									dim datatypeElement as EA.Element
									set datatypeElement = Repository.GetElementByID(currentAttribute.ClassifierID)
									if debug then Session.Output("   The datatype of attribute [" & currentAttribute.Name &  "] is: [" & datatypeElement.Name & "]")
									dim datatypeElementSOSInavn
									Call TVGetElementTaggedValue("class", datatypeElement, "SOSI_navn", datatypeElementSOSInavn)
									if datatypeElementSOSInavn <> "" then
										Call TVSetElementTaggedValue("["+currentElement.Name+"] attribute", currentAttribute, "SOSI_navn", datatypeElementSOSInavn)
									else
										'SOSI_name not found in datatype, just generate new UPPERCASE name from attribute name
										Call TVSetElementTaggedValue("["+currentElement.Name+"] attribute", currentAttribute, "SOSI_navn", currentAttribute.Name)
									end if
								else
									'datatype not connected, just generate new UPPERCASE name from attribute name
									Call TVSetElementTaggedValue("["+currentElement.Name+"] attribute", currentAttribute, "SOSI_navn", currentAttribute.Name)
								end if
							end if
							end if
							end if
							end if
							end if
							end if
						end if
						' Følgende er ikke påkrevet!
					Next
				'retrieve all associations for this element and traverse all roles:
				dim connectors as EA.Collection
				set connectors = currentElement.Connectors

				'iterate the connectors
				'Session.Output("Found " & connectors.Count & " connectors for featureType " & currentElement.Name)
				dim connectorsCounter
				for connectorsCounter = 0 to connectors.Count - 1
					dim currentConnector as EA.Connector
					set currentConnector = connectors.GetAt( connectorsCounter )

					dim sourceElementID
					sourceElementID = currentConnector.ClientID
					dim sourceEndName
					sourceEndName = currentConnector.ClientEnd.Role
					dim targetElementID
					targetElementID = currentConnector.SupplierID
					dim targetEndName
					targetEndName = currentConnector.SupplierEnd.Role
					'Session.Output("Source: " & sourceEndName & " target: " & targetEndName & " type: " & currentConnector.Type)
					dim elementOnOppositeSide as EA.Element
					if currentElement.ElementID = sourceElementID and not currentConnector.Type = "Realization" and not currentConnector.Type = "Generalization" then
						set elementOnOppositeSide = Repository.GetElementByID(targetElementID)
						'Session.Output("1 elementOnOppositeSide.Name: " & elementOnOppositeSide.Name)
						' if element has SOSI_navn then get it TODO

						'else set rolename in capitals
						    Call TVSetElementTaggedValueRole("role", currentConnector.SupplierEnd, "SOSI_navn", targetEndName)

					end if
					if currentElement.ElementID = targetElementID and not currentConnector.Type = "Realization" and not currentConnector.Type = "Generalization" then
						set elementOnOppositeSide = Repository.GetElementByID(sourceElementID)
						'Session.Output("2 elementOnOppositeSide.Name: " & elementOnOppositeSide.Name)
						' if element has SOSI_navn then get it TODO

						'else set rolename in capitals
						    Call TVSetElementTaggedValueRole("role", currentConnector.ClientEnd, "SOSI_navn", sourceEndName)

					end if

				next
				' reset sequenceNumber to the sequence the attributes currently have in the model
				' resequenceAttributes()
				' reset sequenceNumber to a sequence after all the attributes, keep the old internal role sequence
				' resequenceRoles()


				end if

				'Session.Output( "Done with element ["& currentElement.Name &"]")
			next
	if debug then Session.Output( "Done with package ["& package.Name &"]")

end sub


' Sets the specified TaggedValue on the provided element. If the provided element does not already
' contain a TaggedValue with the specified name, a new TaggedValue is created with the requested
' name and value. If a TaggedValue already exists with the specified name then nothing will be changed.
'
'
sub TVGetElementTaggedValue( ownerElementName, theElement, taggedValueName, theValue)
	'Session.Output( "  Checking if tagged value [" & taggedValueName & "] exists")
	if not theElement is nothing and Len(taggedValueName) > 0 then
		dim newTaggedValue as EA.TaggedValue
		set newTaggedValue = nothing
		dim taggedValueExists, taggedValueValue
		taggedValueExists = False
		taggedValueValue = ""
		theValue = ""

		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			if currentExistingTaggedValue.Name = taggedValueName then
				taggedValueValue = currentExistingTaggedValue.Value
				taggedValueExists = True
			end if
		next

		'if the element does not contain a tagged value with the provided name, create a new one
		if not taggedValueExists = True then
			'set newTaggedValue = theElement.TaggedValues.AddNew( taggedValueName, theValue )
			'newTaggedValue.Update()
			'Session.Output( "    Added to " & ownerElementName & " [" & theElement.Name & "] tagged value [" & taggedValueName & " = " & theValue & "]")
			'Session.Output( "    On " & ownerElementName & " [" & theElement.Name & "] tagged value [" & taggedValueName &  "] is not found.")
		else
			if debug then Session.Output( "    Found on " & ownerElementName & " [" & theElement.Name & "] tagged value [" & taggedValueName & " = " & taggedValueValue & "]")
			theValue = taggedValueValue
		end if
	end if
end sub

sub TVSetElementTaggedValue( ownerElementName, theElement, taggedValueName, theValue)
	'Session.Output( "  Checking if tagged value [" & taggedValueName & "] exists")
	if not theElement is nothing and Len(taggedValueName) > 0 then
		dim newTaggedValue as EA.TaggedValue
		set newTaggedValue = nothing
		dim taggedValueExists, taggedValueValue
		taggedValueExists = False
		taggedValueValue = ""

		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			if currentExistingTaggedValue.Name = taggedValueName then
				taggedValueValue = currentExistingTaggedValue.Value
				taggedValueExists = True
			end if
		next

		'if the element does not contain a tagged value with the provided name, create a new one
		if not taggedValueExists = True then
			set newTaggedValue = theElement.TaggedValues.AddNew( taggedValueName, theValue )
			newTaggedValue.Update()
			Session.Output( "    Added to " & ownerElementName & " [" & theElement.Name & "] tagged value [" & taggedValueName & " = " & theValue & "]")
		else
			if debug then Session.Output( "    Found on " & ownerElementName & " [" & theElement.Name & "] tagged value [" & taggedValueName & " = " & taggedValueValue & "]")
		end if
	end if
end sub

sub TVSetElementTaggedValueRole( ownerElementName, theElement, taggedValueName, theValue)
	'Session.Output( "  Checking if tagged value [" & taggedValueName & "] exists")
	if not theElement is nothing and Len(taggedValueName) > 0 then
		dim newTaggedValue as EA.TaggedValue
		set newTaggedValue = nothing
		dim taggedValueExists, taggedValueValue
		taggedValueExists = False
		taggedValueValue = ""
		dim theElementName

		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			if currentExistingTaggedValue.Tag = taggedValueName then
				taggedValueValue = currentExistingTaggedValue.Value
				taggedValueExists = True
			end if
		next

		'if the element does not contain a tagged value with the provided name, create a new one
		if not taggedValueExists = True then
			set newTaggedValue = theElement.TaggedValues.AddNew( taggedValueName, theValue )
			newTaggedValue.Update()
			Session.Output( "    Added to " & ownerElementName & " [" & theElement.Role & "] tagged value [" & taggedValueName & " = " & theValue & "]")
		else
			if debug then Session.Output( "    Found on " & ownerElementName & " [" & theElement.Role & "] tagged value [" & taggedValueName & " = " & taggedValueValue & "]")
		end if
	end if
end sub

function toNCName(namestring, blankbeforenumber)
		' make name legal NCName
    Dim txt, res, tegn, i, u
    u=0
		txt = Trim(namestring)
		res = UCase( Mid(txt,1,1) )
			'Repository.WriteOutput "Script", "New NCName: " & txt & " " & res,0

		' loop gjennom alle tegn
		For i = 2 To Len(txt)
		  ' blank, komma, !, ", #, $, %, &, ', (, ), *, +, /, :, ;, <, =, >, ?, @, [, \, ], ^, `, {, |, }, ~
		  ' (tatt med flere fnuttetyper, men hva med "."?)
		  tegn = Mid(txt,i,1)
		  if tegn = " " or tegn = "," or tegn = """" or tegn = "#" or tegn = "$" or tegn = "%" or tegn = "&" or tegn = "(" or tegn = ")" or tegn = "*" Then
			  'Repository.WriteOutput "Script", "Bad1: " & tegn,0
			  u=1
		  Else
		    if tegn = "+" or tegn = "/" or tegn = ":" or tegn = ";" or tegn = "<" or tegn = ">" or tegn = "?" or tegn = "@" or tegn = "[" or tegn = "\" Then
			    'Repository.WriteOutput "Script", "Bad2: " & tegn,0
			    u=1
		    Else
		      If tegn = "]" or tegn = "^" or tegn = "`" or tegn = "{" or tegn = "|" or tegn = "}" or tegn = "~" or tegn = "'" or tegn = "�" or tegn = "�" Then
			      'Repository.WriteOutput "Script", "Bad3: " & tegn,0
			      u=1
		      else
			      'Repository.WriteOutput "Script", "Good: " & tegn,0
			      If u = 1 Then
		          If tegn = "1" or tegn = "2" or tegn = "3" or tegn = "4" or tegn = "5" or tegn = "6" or tegn = "7" or tegn = "8" or tegn = "9" or tegn = "0" Then
		            res = res + blankbeforenumber + tegn
  			      else
		            res = res + UCase(tegn)
		          End If
		          u=0
			      else
		          res = res + tegn
		        End If
		      End If
		    End If
		  End If
		Next
		'Repository.WriteOutput "Script", "New NCName: " & res,0
    toNCName = res
		Exit function
end function

'start the main function
OnProjectBrowserScript
