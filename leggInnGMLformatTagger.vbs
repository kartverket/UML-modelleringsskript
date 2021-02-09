option explicit

!INC Local Scripts.EAConstants-VBScript

' scriptnavn: leggInnGMLTagger
' formål:	  legger inn nødvendige tagged values for generering av GML-Applikasjonsskjema
' Script Name: 	leggInnGMLformatTagger (AddMissingTags)
' Author: 		Magnus Karge / Kent Jonsrud
' Purpose: 		To add missing tags on model elements 
'               (application schemas, feature types & attributes, data types & attributes,code lists, enumerations)
'				for genetrating GML-ApplicationSchema as defined in the Norwegian standard "SOSI regler for UML-modellering"
' Date: 		2016-08-24     Original:11.09.2015   + Moddet av Kent 2016-03-09/08-24: Legger nå inn forslag til verdi i alle taggene!

' Date: 		2021-02-09  Går nå gjennom alle underpakker
'							setter inn sequenceNumber på alle navigerbare assosiasjonsender (med roller)
'							setter inn inLineOrByReference = byReference på navigerbare assosiasjonsender mot FeatureTypes
'
'
'	TBD: ta bort unødvendig utskrift av tagger som allerede finnes

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
			mess = 	"Creates tags needed for generating GML format from model elements. Script version 2021-02-09." & vbCrLf
			mess = mess + "NOTE! This script may add content to all model elements in package: "& vbCrLf & "[«" & thePackage.element.Stereotype & "» " & thePackage.Name & "]."

			box = Msgbox (mess, vbOKCancel)
			select case box
			case vbOK
				if LCase(thePackage.element.Stereotype) = "applicationschema" then
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
			dim conn as EA.Collection
			dim connEnd as EA.ConnectorEnd
			Session.Output("The current package is: " & package.Name)
			'if the current package has stereotype applicationSchema then check tagged values
			if package.element.stereotype = "applicationSchema" or package.element.stereotype = "ApplicationSchema" then
				' Kapittel13kravSOSI: language, version, targetNamespace, SOSI_kortnavn, SOSI_modellstatus, SOSI_versjon
				' Kapittel13kravGML: language, version, targetNamespace, xmlns, xsdDocument, SOSI_modellstatus
				' Kapittel13kravAlle: designation og definition for engelsk
				ASpackage = "http://skjema.geonorge.no/SOSI/produktspesifikasjon/" + toNCName(package.Name,"/-DraftName")
				'Call TVSetElementTaggedValue("ApplicationSchema",package.element, "SOSI_kortnavn",toNCName(package.Name,"-"))
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "SOSI_modellstatus","utkastOgSkjult")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "language","no")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "version","0.1")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "targetNamespace",ASpackage)
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "xmlns","app")
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "xsdDocument",toNCName(package.Name, "-") & "-DraftName.xsd")
				'Call TVSetElementTaggedValue("ApplicationSchema",package.element, "definition","""""@en")
				' TODO: Klipp inn engelsk fra notefeltet dersom du finner --Definition --
				'Call TVSetElementTaggedValue("ApplicationSchema",package.element, "designation","""""@en")
				' TODO: Klipp inn det engelske navnet fra Alias-feltet dersom dette finnes
				' Er denne også praktisk å ha med her nå?
				Call TVSetElementTaggedValue("ApplicationSchema",package.element, "xsdEncodingRule","sosi")
				'
				'TODO: sette korrekt case på pakkestereotypen?
				'TODO: package.element.stereotype = "ApplicationSchema"
				'TODO: package.element.stereotype.Update()
				'
			end if

			dim elements as EA.Collection
			'collection of elements that belong to this package (classes, notes... BUT NO packages)
			set elements = package.Elements

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

			'navigate the elements collection
			dim elementsCounter
			for elementsCounter = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( elementsCounter )

				Session.Output("The current element is: " & currentElement.Name & " [Stereotype: " & currentElement.Stereotype & "]")

				'check if the currentElement has stereotype FeatureType.
				if ((currentElement.Stereotype = "FeatureType") or (currentElement.Stereotype = "featureType")) then
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
					'Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
					'Call TVSetElementTaggedValue(currentElement, "isCollection")
					' Følgende er ikke påkrevet!
					'Call TVSetElementTaggedValue("FeatureType", currentElement, "byValuePropertyType", "false")
					'Call TVSetElementTaggedValue("FeatureType", currentElement, "noPropertyType", "true")
				end if

				'check if the currentElement has stereotype CodeList.
				if ((currentElement.Stereotype = "CodeList") or (currentElement.Stereotype = "codeList")) then
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
					'Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
					'Call TVSetElementTaggedValue(currentElement, "SOSI_datatype")
					'Call TVSetElementTaggedValue(currentElement, "SOSI_lengde")
					' Følgende er ikke påkrevet!
					Call TVSetElementTaggedValue("CodeList", currentElement, "asDictionary", "false")
					if ASpackage <> "" then
						Call TVSetElementTaggedValue("CodeList", currentElement, "codeList", ASpackage + "/" + currentElement.Name)
					end if
				end if

				'check if the currentElement has stereotype dataType.
				if ((currentElement.Stereotype = "DataType") or (currentElement.Stereotype = "dataType")) then
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
					'Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
				end if

				'check if the currentElement has stereotype enumeration.
				if ((currentElement.Stereotype = "Enumeration") or (currentElement.Stereotype = "enumeration")) then
					'Call TVSetElementTaggedValue(currentElement, "SOSI_navn")
					'call sub function TVSetElementTaggedValue
					'one function call for each of the required tags
				end if

				'if the currentElement has stereotype dataType or FeatureType then
				'navigate the attributes and check for missing tags
				if ((currentElement.Stereotype = "DataType") or (currentElement.Stereotype = "dataType") or (currentElement.Stereotype = "FeatureType") or (currentElement.Stereotype = "featureType")) then
					dim attributesCounter
					for attributesCounter = 0 to currentElement.Attributes.Count - 1
						dim currentAttribute as EA.Attribute
						set currentAttribute = currentElement.Attributes.GetAt ( attributesCounter )
						'Session.Output( "  The current attribute is ["& currentAttribute.Name &"]")
						'call sub function TVSetElementTaggedValue
						'one function call for each of the required tags
						'Call TVSetElementTaggedValue(currentAttribute, "SOSI_navn")
						'Call TVSetElementTaggedValue(currentAttribute, "SOSI_datatype")
						'Call TVSetElementTaggedValue(currentAttribute, "SOSI_lengde")
						' Følgende er ikke påkrevet!
						'Call TVSetElementTaggedValue(currentElement.Name, currentAttribute, "inLineOrByReference", "inline")
						'Call TVSetElementTaggedValue(currentElement.Name, currentAttribute, "isMetadata", "false")

						'Call TVSetElementTaggedValue(currentElement.Name, currentAttribute, "sequenceNumber", "1")
						'Call TVSetElementTaggedValue(currentElement.Name, currentAttribute, "sequenceNumber", "")
					Next
					' traverse all roles: tbd

					for each conn in currentElement.Connectors
						if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then

						else
							if conn.SupplierEnd.Navigable = "Navigable" then
								Call TVSetEndTaggedValue(currentElement.Name, conn.SupplierEnd, "sequenceNumber", "")
								if LCase(Repository.GetElementByID(conn.SupplierID).Stereotype) = "featuretype" then 
									Call TVSetEndTaggedValue(currentElement.Name, conn.SupplierEnd, "inLineOrByReference", "byReference")
								end if
								conn.SupplierEnd.Update()
							end if
							if conn.ClientEnd.Navigable = "Navigable" then
								Call TVSetEndTaggedValue(currentElement.Name, conn.ClientEnd, "sequenceNumber", "")
								if LCase(Repository.GetElementByID(conn.ClientID).Stereotype) = "featuretype" then 
									Call TVSetEndTaggedValue(currentElement.Name, conn.ClientEnd, "inLineOrByReference", "byReference")
								end if
								conn.ClientEnd.Update()
							end if
						end if
					next
									
									
								' reset sequenceNumber to the sequence the attributes currently have in the model (?)
								' resequenceAttributes()
								' reset sequenceNumber to a sequence after all the attributes, keep the old internal role sequence
								' resequenceRoles()


				end if

				Session.Output( "Done with element ["& currentElement.Name &"]")
			next
	Session.Output( "Done with package ["& package.Name &"]")

	dim subP as EA.Package
	for each subP in package.packages
	    call FindElementsWithMissingTagsInPackage(subP)
	next
	
end sub


' Sets the specified TaggedValue on the provided element. If the provided element does not already
' contain a TaggedValue with the specified name, a new TaggedValue is created with the requested
' name and value. If a TaggedValue already exists with the specified name then nothing will be changed.
'
'
sub TVSetEndTaggedValue( ownerElementName, connectorEnd, taggedValueName, theValue)
	'Session.Output( "  Checking if tagged value [" & taggedValueName & "] exists")
	if not connectorEnd is nothing and Len(taggedValueName) > 0 then
		dim newTaggedValue as EA.RoleTag
		set newTaggedValue = nothing
		dim taggedValueExists, taggedValueValue
		taggedValueExists = False
		taggedValueValue = ""

		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.RoleTag
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to connectorEnd.TaggedValues.Count - 1
			set currentExistingTaggedValue = connectorEnd.TaggedValues.GetAt(taggedValuesCounter)
			if currentExistingTaggedValue.Tag = taggedValueName then
				taggedValueValue = currentExistingTaggedValue.Value
				taggedValueExists = True
			end if
		next

		'if the element does not contain a tagged value with the provided name, create a new one
		if not taggedValueExists = True then
			set newTaggedValue = connectorEnd.TaggedValues.AddNew( taggedValueName, theValue )
			newTaggedValue.Update()
			Session.Output( "    ADDED To " & ownerElementName & " " & connectorEnd.Role & " tagged value [" & taggedValueName & " = " & theValue & "]")
		else
			Session.Output( "    FOUND On " & ownerElementName & " " & connectorEnd.Role & " tagged value [" & taggedValueName & " = " & taggedValueValue & "]")
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
			Session.Output( "    ADDED To " & ownerElementName & " " & theElement.Name & " tagged value [" & taggedValueName & " = " & theValue & "]")
		else
			Session.Output( "    FOUND On " & ownerElementName & " " & theElement.Name & " tagged value [" & taggedValueName & " = " & taggedValueValue & "]")
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
		      If tegn = "]" or tegn = "^" or tegn = "`" or tegn = "{" or tegn = "|" or tegn = "}" or tegn = "~" or tegn = "'" or tegn = "´" or tegn = "¨" Then
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
