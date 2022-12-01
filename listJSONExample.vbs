option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		listJsonExample
' purpose:		Generates example objects in a combination (?) of GeoJSON and JSON-FG from all non-abstract feature types in a model
' note:         Important vbs-snag: decimal point character for numbers need to be set to "." in the regional settings
'
' author: 		Kent Jonsrud
'
' version:		2022-11-11/30 geometry types (WGS84 TBD)
' version:		2022-11-09 rydda i skilletegn, lagt inn en links
' version:		2022-11-04 timestamp
' version:		2022-10-17 Based on a script listGMLExample from 2018 to generate GML-example in system output from classe in a ApplicationSchema package
'
' 				candidates for improvement:
' TBD:			variants of GM_Solid geometry (Polyhedron, Prism, ...)
' TBD:			add the prefix on elements from external ApplicationSchema packages, sort (attributes?+roles!) by tagged value sequenceNumber
' TBD:          utf8 (for all codes with any sami characters) does not survive the clipboard, should concider to write directly to file
' TBD:			loopdetection on inheritance and datatype, association is ok-ish
' TBD:			standard (getFeatureById-) container for the single object selection
' TBD:			documentation and cleanup


	DIM epsg, debug, namespace, kortnavn, pnteller, cuteller, suteller, soteller, obteller, eastoffset, eastdelta, ftcount, ftnum, gnum

	'roof - starting on 1
	dim p1(18,2)
	'walls - starting on 1
	dim b1(29,2)
	'surface - starting on 0
	dim s1(7,2)
	'point
	dim r1, r2, r3
	'curve - starting on 1
	dim q1(5), q2(5), q3(5)
		
	debug = false

sub listJsonExample()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	dim ftname
	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		if Repository.GetTreeSelectedItemType() = otPackage then
			if UCASE(theElement.Element.Stereotype) = "APPLICATIONSCHEMA" then
				dim message, indent
			'	dim box
			'	box = Msgbox ("Script listJsonExample" & vbCrLf & vbCrLf & "Scriptversion 2022-10-17" & vbCrLf & "Listing to JSON example for package : [" & theElement.Name & "].",1)
			'	select case box
			'	case vbOK
					dim xsdfile
					'tømmer System Output 
					Repository.ClearOutput "Script"
					Repository.CreateOutputTab "Error"
					Repository.ClearOutput "Error"
					kortnavn = getPackageTaggedValue(theElement,"SOSI_kortnavn")
					if kortnavn = "" then
						kortnavn = theElement.Name
					end if

					namespace = getPackageTaggedValue(theElement,"targetNamespace")
					if namespace = "" then
						namespace = "http://schemas.someserver.org/" & kortnavn
					end if
					
					xsdfile = getPackageTaggedValue(theElement,"xsdDocument")
					if xsdfile = "" then
						xsdfile = kortnavn & ".xsd"
					end if
					SessionOutput("{")
					SessionOutput("  ""type"": ""FeatureCollection"",")
					SessionOutput("  ""time"": { ""timestamp"": """ & nao() & """ },")
					SessionOutput("  ""describedby"": { ""href"": """ & namespace & "/" & kortnavn & """ },")

					SessionOutput("  ""features"": [")
					ftnum = 0
					ftcount = 0
					call getFeatureTypeCount(theElement)
					call initCoord()
					call listFeatureTypes(theElement)

					SessionOutput("  ]")
					SessionOutput("}")

					

			'	case VBcancel

			'	end select
			else			'No «ApplicationSchema» Package or a «FeatureType» Class selected in the tree
				MsgBox( "This script requires a «ApplicationSchema» Package or a «FeatureType» Class to be selected in the Project Browser." & vbCrLf & _
				"Please select a «ApplicationSchema» Package or a «FeatureType» Class  in the Project Browser and try again." )
		
			end if
		Else
			if Repository.GetTreeSelectedItemType() = otElement then
				if theElement.Type="Class" and UCASE(theElement.Stereotype) = "FEATURETYPE" then
					if debug then Repository.WriteOutput "Script", "Debug: theElement.Name [«" & theElement.Stereotype & "» " & theElement.Name & "] currentElement.Type [" & theElement.Type & "] currentElement.Abstract [" & theElement.Abstract & "].",0

					Repository.ClearOutput "Script"
					Repository.CreateOutputTab "Error"
					Repository.ClearOutput "Error"
					namespace = "http://schemas.someserver.org/someproduct"
					kortnavn = "shortNamespace"
					indent = "    "
					ftname = theElement.Name

					indent = "      "
					SessionOutput("{")
					SessionOutput("  ""type"": ""Feature"",")
					SessionOutput("  ""time"": { ""timestamp"": """ & nao() & """ },")
					SessionOutput("  ""describedby"": { ""href"": """ & namespace & "/" & kortnavn & """ },")
					
					call initCoord()
					SessionOutput("   ""id"": """ & "http://data.geonorge.no/sosi/" & Kortnavn & "/" & ftname & ".1"",")

					SessionOutput("  ""featureType"": """ & ftname & """,")
					SessionOutput("  ""time"": { ""timestamp"": """ & nao() & """ },")
					SessionOutput("  ""coordRefSys"": ""http://www.opengis.net/def/crs/EPSG/0/5972"",")
					gnum = 0
					SessionOutput("  ""place"": {")
					call listJsonFgGeometry(ftname,theElement,indent)
					SessionOutput("  },")	
					
					gnum = 0
					SessionOutput("  ""geometry"": {")
					call listGeoJsonGeometry(ftname,theElement,indent)
					SessionOutput("  },")

					SessionOutput("  ""links"": [")
					call listLinks(indent)
					SessionOutput("  ],")
					
					SessionOutput("  ""properties"": {")
					call listDatatypes(ftname,theElement,indent)
					SessionOutput("  }")


					SessionOutput("}")					
'					SessionOutput("    </" & utf8(theElement.Name) & ">")
'					SessionOutput("  </wfs:member>")
				else
					'Other than «ApplicationSchema» Package or a «FeatureType» Class selected in the tree
					MsgBox( "This script requires a Package or a «FeatureType» Class to be selected in the Project Browser." & vbCrLf & _
					"Please select a Package or a «FeatureType» Class in the Project Browser and try again." )
				end if
			else
				'Other than «ApplicationSchema» Package or a «FeatureType» Class selected in the tree
				MsgBox( "Element type selected: " & theElement.Type & vbCrLf & _
				"This script requires a Package or a «FeatureType» Class to be selected in the Project Browser." & vbCrLf & _
				"Please select a Package or a «FeatureType» Class in the Project Browser and try again." )
			end If
		end if
		'Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
	end if
end sub


sub listFeatureTypes(pkg)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim super as EA.Element
	dim datatype as EA.Element
	dim conn as EA.Collection
 	set elements = pkg.Elements 
	dim i, sosinavn, sositype, sosilengde, sosimin, sosimax, koder, prikkniv, sosierlik, superlist
	dim indent, ftname
	if debug then Repository.WriteOutput "Script", "<!-- Debug: pkg.Name [" & pkg.Name & "]. -->",0

	dim currentElement as EA.Element 
 
	for i = 0 to elements.Count - 1 
		set currentElement = elements.GetAt( i ) 
				
		if debug then Repository.WriteOutput "Script", "<!-- Debug: currentElement.Name [«" & currentElement.Stereotype & "» " & currentElement.Name & "] currentElement.Type [" & currentElement.Type & "] currentElement.Abstract [" & currentElement.Abstract & "]. -->",0
		if currentElement.Type = "Class" and LCase(currentElement.Stereotype) = "featuretype" and currentElement.Abstract = 0 then
			ftnum = ftnum + 1
			SessionOutput("      {")
			SessionOutput("        ""type"": ""Feature"",")
			
			ftname = currentElement.Name
			superlist = ""
			indent = "           "

			SessionOutput("        ""id"": """ & "http://data.geonorge.no/sosi/" & Kortnavn & "/" & ftname & ".1"",")

			SessionOutput("        ""featureType"": """ & ftname & """,")
			SessionOutput("        ""time"": { ""timestamp"": """ & nao() & """ },")
			SessionOutput("        ""coordRefSys"": ""http://www.opengis.net/def/crs/EPSG/0/5972"",")
			gnum = 0
			SessionOutput("        ""place"": {")
			call listJsonFgGeometry(ftname,currentElement,indent)
			SessionOutput("        },")	
			
			gnum = 0
			SessionOutput("        ""geometry"": {")
			call listGeoJsonGeometry(ftname,currentElement,indent)
			SessionOutput("        },")

			SessionOutput("        ""links"": [")
			call listLinks(indent)
			SessionOutput("        ],")
			
			SessionOutput("        ""properties"": {")
			call listDatatypes(ftname,currentElement,indent)
			SessionOutput("        }")

			if ftnum >= ftcount then
				SessionOutput("      }")
			else
				SessionOutput("      },")
			end if
'			SessionOutput("    </" & utf8(currentElement.Name) & ">")
'			SessionOutput("  </wfs:member>")
			eastoffset = eastoffset + eastdelta	

		end if
	
	next

	dim subP as EA.Package
	for each subP in pkg.packages
	    call listFeatureTypes(subP)
	next


end sub

'
'	List association roles as links, attributes and subattributes except attributes with geometry type
'
sub listDatatypes(ftname,element,indent)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim element0 as EA.Element
	dim super as EA.Element
	dim datatype as EA.Element
	dim subbtype as EA.Element
	dim conn as EA.Collection
	dim connEnd as EA.ConnectorEnd
	dim i, umlnavn, sosinavn, sositype, sosilengde, sosimin, sosimax, sosierlik, koder, prikkniv1, roleEndElementID, sosidef, selfref, subID
	dim indent0, indent1, superlist, ccount, acount, cnum, anum, cont
	dim code1, code2, codeListUrl
	
'	eastoffset = eastoffset + eastdelta		
	if debug then Repository.WriteOutput "Script", "<!-- Debug: --------listDatatypes element.Name [" & element.Name & "] element.ElementID [" & element.ElementID & "]. -->",0
'ZZZ	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "" or LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
		' sjekk om siste klasse i arverekken er tom
		acount = element.Attributes.Count
		ccount = element.Connectors.Count
		for each conn in element.Connectors
			if debug then Repository.WriteOutput "Script", "<!-- Debug: conn.Type [" & conn.Type & "] conn.ClientID [" & conn.ClientID & "] conn.SupplierID [" & conn.SupplierID & "]. -->",0
			if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then
				ccount = ccount - 1
			end if
			if conn.Type = "Generalization" then
				if element.ElementID = conn.ClientID then
					if debug then Repository.WriteOutput "Script", "<!-- Debug: supertype [" & Repository.GetElementByID(conn.SupplierID).Name & "]. -->",0
'					superlist = getSupertypes(ftname, conn.SupplierID, indent)
					set super = Repository.GetElementByID(conn.SupplierID)
					call listDatatypes(ftname,super,indent)
					if acount > 0 or ccount > 0 then
						SessionOutput(indent & "," )
						' TBD empty supertypes ?
					end if
				end if
			end if
		next


		dim attr as EA.Attribute
		anum = 0
		cont = ","
		for each attr in element.Attributes
			anum = anum + 1
	'		if anum >= acount or LCase(element.Stereotype) = "union" then
			if anum >= acount and ccount = 0 then
				cont = ""
				'TBD if isLastNonGeometryAttribute(element,attr.Name...)
			end if
			if debug then Repository.WriteOutput "Script", "<!-- Debug: anum,acount,ccont,cont [" & anum & " " & acount & " "  & ccount & " " & cont & "] -->",0
			if getSosiGeometritype(attr) = "" then
				' normal attribute, not geometry
				if debug then Repository.WriteOutput "Script", "<!-- Debug: attr.Name [" & attr.Name & "] not geometry. -->",0
				if attr.ClassifierID <> 0 and getBasicSOSIType(attr.Type) = "*" then
					set datatype = Repository.GetElementByID(attr.ClassifierID)
					'see if the datatype has a supertype, if so then write all its elements first - TBD
					
					if datatype.Name = element.Name and datatype.ParentID = element.ParentID then
					'if datatype.ClassifierID = element.ClassifierID then
						Repository.WriteOutput "Script", "Error - circular self reference: datatype.Name [" & datatype.Name & "] from attribute name [" & element.Name & "." & attr.Name & "].",0
						exit sub
					else
						if datatype.Type = "Enumeration" or LCase(datatype.Stereotype) = "codelist" or LCase(datatype.Stereotype) = "enumeration" then
							'list first code in the list
							if getTaggedValue(attr,"inlineOrByReference") = "byReference" then
								'variant gml:ReferenceType
								'if debug then 
								if attr.UpperBound <> "1" then
								'	SessionOutput(indent & "<" & attr.Name & ">" & listCodeType(datatype) & "</" & attr.Name & ">")
					'				SessionOutput(indent & "<" & attr.Name & " xlink:href=""" & namespace & "/" & attr.Type & "/" & listCodeType(datatype) & """/>")
									SessionOutput(indent & """" & attr.Name & """: """ & namespace & "/" & attr.Type & "/" & listCodeType(datatype) & """")
								end if
								SessionOutput(indent & "<" & attr.Name & " xlink:href=""" & namespace & "/" & attr.Type & "/" & listCodeType(datatype) & """/>")
								'SessionOutput(indent & "<" & attr.Name & " xlink:href=""" & listReferenceType(attr.Type) & """/>")

							else
								if getTaggedValue(datatype,"asDictionary") = "true" then
								'	if getTaggedValue(datatype,"codeList") <> "" and getTaggedValue(datatype,"codeList") = getTaggedValue(attr,"defaultCodeSpace") then
									if getTaggedValue(datatype,"codeList") <> "" then
									'TBD read tag defaultCodeSpace on the attribute instead
										codeListUrl = getTaggedValue(datatype,"codeList")
										if getTaggedValue(attr,"defaultCodeSpace") <> "" then
											if getTaggedValue(attr,"defaultCodeSpace") <> codeListUrl then
												Repository.WriteOutput "Script", "<!-- Info: attr.Name [" & attr.Name & "] has a tagged value defaultCodeSpace [" & getTaggedValue(attr,"defaultCodeSpace") & "] that is different from the codelist class tagged value codeList [" & codeListUrl & "]. -->",0
											end if
											codeListUrl = getTaggedValue(attr,"defaultCodeSpace")
										end if
										if debug then Repository.WriteOutput "Script", "<!-- Debug: codeListUrl [" & codeListUrl & "]. -->",0
	
										call get2firstCodes(codeListUrl,code1,code2)
										
										if code1 = "" then
											code1 = "someExternalCodelistCode"
										end if
										if code2 = "" then
											code2 = "someOtherExternalCodelistCode"
										end if
									else
										code1 = "someExternalCodelistCode"
										code2 = "someOtherExternalCodelistCode"
									end if
									if attr.UpperBound <> "1" then
										SessionOutput(indent & """" & attr.Name & """: """ & code1 & """,")
										SessionOutput(indent & """" & attr.Name & """: """ & code2 & """" & cont)
									else
										SessionOutput(indent & """" & attr.Name & """: """ & code1 & """" & cont)
									end if	
							
								else
									'TBD asDictionary= false, variant gml:CodeType
									if code1 = "" then
										code1 = "someInternalCodelistCode"
									end if
									if code2 = "" then
										code2 = "someOtherInternalCodelistCode"
									end if
									if attr.UpperBound <> "1" then
										SessionOutput(indent & """" & attr.Name & """: """ & code1 & """,")
										SessionOutput(indent & """" & attr.Name & """: """ & code2 & """")
									else			
										SessionOutput(indent & """" & attr.Name & """: """ & code1 & """")
									end if
								end if
							end if
							'listCodeType(attr)
						else
							if attr.UpperBound <> "1" then
								SessionOutput(indent & """" & utf8(attr.Name) & """: {")
								indent0 = indent & "  "
								call listDatatypes(ftname, datatype,indent0)
								SessionOutput(indent & "},")
							end if
							SessionOutput(indent & """" & utf8(attr.Name) & """: {")
							indent0 = indent & "  "
							call listDatatypes(ftname, datatype,indent0)
							SessionOutput(indent & "}" & cont)
							
						end if
					end if
				else
					'base type
					if attr.UpperBound <> "1" then
						SessionOutput(indent & """" & utf8(attr.Name) & """: """ & listBaseType(ftname, attr.Name,attr.Type) & """,")
					end if
					SessionOutput(indent & """" & utf8(attr.Name) & """: """ & listBaseType(ftname, attr.Name,attr.Type) & """" & cont)

				end if
			else
				'geometry type
			end if

			'TBD if Union then jump out of the loop after first(!) variant, this does not support well Unions having several different datatypes 
			if LCase(element.Stereotype) = "union" then
				Exit For
			end if
			
		next

		cnum = 0
		cont = ","
		for each conn in element.Connectors
			if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then
			
			else
				cnum = cnum + 1
				if cnum >= ccount then
					cont = ""
				end if
				if debug then Repository.WriteOutput "Script", "<!-- Debug: cnum,acount,ccont,cont [" & cnum & " " & acount & " "  & ccount & " " & cont & "] -->",0

				if debug then Repository.WriteOutput "Script", "<!-- Debug: Supplier Role.Name [" & conn.SupplierEnd.Role & "] datatypens navn [" & Repository.GetElementByID(conn.SupplierID).Name & "], conn.SupplierID [" & conn.SupplierID & "]. -->",0
				if debug then Repository.WriteOutput "Script", "<!-- Debug: Client Role.Name [" & conn.ClientEnd.Role & "] datatypens navn [" & Repository.GetElementByID(conn.ClientID).Name & "], conn.ClientID [" & conn.ClientID & "]. -->",0

				if conn.ClientID = element.ElementID then
				'	if getConnectorEndTaggedValue(conn.SupplierEnd,"xsdEncodingRule") <> "notEncoded" then
						set datatype = Repository.GetElementByID(conn.SupplierID)
						umlnavn = conn.SupplierEnd.Role
						if conn.ClientEnd.Aggregation = 2 and conn.SupplierID <> conn.ClientID then
'						if conn.ClientEnd.Aggregation = 2 and conn.SupplierID <> conn.ClientID or getConnectorEndTaggedValue(conn.SupplierEnd,"inlineOrByReference") = "inline"  then
							'composition+mandatory->nest as datatype inline?
							'inLineOrByReference?
							call listComposition(ftname,umlnavn,datatype,indent,conn.SupplierEnd.Cardinality,cont)
						else
							call listAssociation(ftname,umlnavn,datatype,indent,conn.SupplierEnd.Cardinality,conn.SupplierEnd.Navigable,element.Name,cont)
						end if
				'	end if

				else
				'	if getConnectorEndTaggedValue(conn.ClientEnd,"xsdEncodingRule") <> "notEncoded" then
						set datatype = Repository.GetElementByID(conn.ClientID)
						umlnavn = conn.ClientEnd.Role
						if conn.SupplierEnd.Aggregation = 2 then
							'composition+mandatory->nest as datatype inline?
							'inLineOrByReference?
							call listComposition(ftname,umlnavn,datatype,indent,conn.ClientEnd.Cardinality,cont)
						else
							call listAssociation(ftname,umlnavn,datatype,indent,conn.ClientEnd.Cardinality,conn.ClientEnd.Navigable,element.Name,cont)
						end if
				'	end if
				end if
			end if

		next

	end if

end sub



sub listLinks(indent)

	SessionOutput(indent & "{")
	SessionOutput(indent & "  ""href"": ""https://example.org/data/v1/collections/buildings/items/DENW19AL0000giv5BL?f=jsonfg"",")
	SessionOutput(indent & "  ""rel"": ""self"",")
	SessionOutput(indent & "  ""type"": ""application/vnd.ogc.fg+json"",")
	SessionOutput(indent & "  ""title"": ""This document""")
	SessionOutput(indent & "}")

end sub

sub listComposition(ftname,umlnavn,datatype,indent,cardinality,cont)
	dim indent0, indent1, subID
	dim subbtype as EA.Element

	SessionOutput(indent & """" & utf8(umlnavn) & """ {")
	indent0 = indent & "  "
'	'	SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
	indent1 = indent0 & "  "
	if datatype.Abstract = 1 then
		'must move down to make an example of a instanciable subtype of the class pointed to TODO, NB needed on mandatory attributes!
		call getFirstConcreteSubtypeName(datatype,subID)
		set subbtype = Repository.GetElementByID(subID)
	'	SessionOutput(indent0 & """" & utf8(subbtype.Name) & """: ")
		call listDatatypes(ftname, subbtype,indent1)
	'	SessionOutput(indent0 & "}")
	else
	'	SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
		call listDatatypes(ftname, datatype,indent1)
	'	SessionOutput(indent0 & "}")
	end if
	SessionOutput(indent & "}")
	if cardinality <> "0..1" and cardinality <> "1..1" and cardinality <> "1" then
		SessionOutput(indent & "<" & utf8(umlnavn) & ">")
		indent0 = indent & "  "
		SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
		indent1 = indent0 & "  "
		if datatype.Abstract = 1 then
			'must move down to make an example of a instanciable subtype of the class pointed to TODO, NB needed on mandatory attributes!
			call getFirstConcreteSubtypeName(datatype,subID)
			set subbtype = Repository.GetElementByID(subID)
			SessionOutput(indent0 & "<" & utf8(subbtype.Name) & ">")
			call listDatatypes(ftname, subbtype,indent1)
			SessionOutput(indent0 & "</" & utf8(subbtype.Name) & ">")
		else
			SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
			call listDatatypes(ftname, datatype,indent1)
			SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
		end if
		SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
		SessionOutput(indent & "</" & utf8(umlnavn) & ">")
	end if
end sub


sub listAssociation(ftname,umlnavn,datatype,indent,cardinality,navigable,elementname,cont)
'sub listComposition(ftname,umlnavn,datatype,indent,cardinality)
	dim indent0, indent1, subID, selfref
	dim subbtype as EA.Element
	
	if navigable = "Navigable" then
		'self assoc? if so make xlinks to other (imaginary) instances of the same class
		selfref = 1
			if datatype.Name = elementname then
				selfref = 2
			end if 
			if cardinality <> "0..1" and cardinality <> "1..1" and cardinality <> "1" then
			if datatype.Abstract = 1 then
				'must move down to make an example of a instanciable subtype of the class pointed to TODO, NB needed on mandatory attributes!
				SessionOutput(indent & """" & utf8(umlnavn) & """: ""#" & utf8(getFirstConcreteSubtypeName(datatype,subID)) & "." & selfref & """,")
			else
				SessionOutput(indent & """" & utf8(umlnavn) & """: ""#" & utf8(datatype.Name) & "." & selfref & """,")
			end if
			if datatype.Name = elementname then
				selfref = 3
			end if 
		end if
		if datatype.Abstract = 1 then
			SessionOutput(indent & """" & utf8(umlnavn) & """: ""#" & utf8(getFirstConcreteSubtypeName(datatype,subID)) & "." & selfref & """" & cont)
		else
			SessionOutput(indent & """" & utf8(umlnavn) & """: ""#" & utf8(datatype.Name) & "." & selfref & """" & cont)
		end if
		if debug then Repository.WriteOutput "Script", "<!-- Debug: .Cardinality [" & cardinality & "]. -->",0

	end if
							
end sub

function listBaseType(ftname,umlname, umltype)
	listBaseType = "FIX"
	if umltype = "CharacterString" then
		if umlname = "navnerom" or umlname = "namespace" then
			listBaseType = "http://data.geonorge.no/sosi/" & Kortnavn 
		else
			if umlname = "lokalId" or umlname = "localId" then
				listBaseType = ftname & ".1"
			else
				if umlname = "versjonId" or umlname = "versionId" then
					listBaseType = "version_1_of_this_object"
				else
					listBaseType = "Some meaningful text"
				end if
			end if
		end if
	end if
	if umltype = "Boolean" then
		listBaseType = "true"
	end if
	if umltype = "Date" then
		listBaseType = "2022-11-04"
	end if
	if umltype = "DateTime" then
		listBaseType = "2022-11-04T21:08:00Z"
	end if
	if umltype = "Integer" then
		listBaseType = "42"
	end if
	if umltype = "Real" then
		listBaseType = "21.22"
	end if
	if UCase(umltype) = "URI" then
		listBaseType = "http://"
	end if
end function


function listCodeType(element)
	listCodeType = "*"
	dim attr as EA.Attribute
	if element.Attributes.Count = 0 then
		listCodeType = "someExternalRegistryCode"
	else
		for each attr in element.Attributes
			listCodeType = attr.Name
			if attr.Default <> "" then listCodeType = attr.Default
			exit for
		next
	end if
end function



sub get2firstCodes(codeListUrl,code1,code2)
	Dim codelist
	codelist = ""
	code1 = ""
	code2 = ""
	' testing http get
	if codeListUrl <> "" then
	'	Session.Output("<!-- DEBUG codeListUrl: " & codeListUrl & " -->")
		Dim httpObject
		Dim parseText, line, linepart, part, kodenavn, kodedef, ualias, kodelistenavn
		Set httpObject = CreateObject("MSXML2.XMLHTTP")
	'	httpObject.open "GET", "http://skjema.geonorge.no/SOSI/basistype/Integer.html", false
		httpObject.open "GET", codeListUrl & ".gml", false
		httpObject.send
		if httpObject.status = 200 then
	'		Session.Output("DEBUG gml:Dictionary: "&httpObject.responseText&"")
	''		parseText = split(split(split(ResponseXML,SearchTag)(1),"</")(0),">")(1)
			parseText = split(httpObject.responseText,"<",-1,1)
			

			kodelistenavn = ""
			for each line in parseText
	'			Session.Output("DEBUG line: "&line&"")
				if mid(line,1,25) = "gml:identifier codeSpace=" then
					linepart = split(line,">",-1,1)
					for each part in linepart
						ualias = part
					next
				end if
				if mid(line,1,16) = "gml:description>" then
				linepart = split(line,">",-1,1)
					for each part in linepart
						kodedef = part
					next
				end if		
				if mid(line,1,9) = "gml:name>" then
				linepart = split(line,">",-1,1)
					for each part in linepart
						kodenavn = part
					next
				end if					
				

				if codelist <> "" and code1 <> "" and code2 = "" and ualias <> "" then code2 = ualias
		'		if codelist <> "" and code1 <> "" and code2 = "" then code2 = kodenavn
				if codelist <> "" and code1 = "" and ualias <> "" then code1 = ualias
		'		if codelist <> "" and code1 = "" then code1 = kodenavn
				if codelist = "" and ualias <> "" then codelist = ualias
				ualias = ""
				'if code2 <> "" then exit for					
			next
	'		Session.Output("|===")
		else
	'		Session.Output("Kodeliste kunne ikke hentes fra register: "&codeListUrl&"")	
	'		Session.Output(" ")		
			if debug then Session.Output("<!-- DEBUG feil ved lesing av kodeliste: ["&codeListUrl&"] status:["&httpObject.status&"]-->")
		end if
	end if
end sub


sub listGeometryType(elementName, geomtype, indent)

		if geomtype = "Punkt" or geomtype = "GM_Point" then
				pnteller = pnteller + 1
				if epsg = 5972 then
'					SessionOutput(indent & "<gml:Point gml:id=""" & elementName & ".pn." & pnteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "<gml:Point srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
'					SessionOutput(indent & "  <gml:pos>568497.68 6662024.15 90.67</gml:pos>")
					SessionOutput(indent & "  <gml:pos>"&r1+eastoffset&" "&r2&" "&r3&"</gml:pos>")
				else
					SessionOutput(indent & "<gml:Point gml:id=""" & elementName & ".pn." & pnteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
					SessionOutput(indent & "  <gml:pos>60.02 10.1</gml:pos>")
				end if
				SessionOutput(indent & "</gml:Point>")
		end if
		if geomtype = "Sverm" or geomtype = "GM_MultiPoint" then
			'getSosiGeometritype = "SVERM"
		end if


		
		if geomtype = "Kurve" or geomtype = "GM_Curve" then
				cuteller = cuteller + 1
				if epsg = 5972 then
'					SessionOutput(indent & "<gml:LineString gml:id=""" & elementName & ".cu." & cuteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "<gml:LineString srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
'					SessionOutput(indent & "  <gml:posList>568597.68 6662024.15 90.67 568497.68 6662024.15 90.67</gml:posList>")
					SessionOutput(indent & "  <gml:posList>"&q1(1)+eastoffset&" "&q2(1)&" "&q3(1)&" "&q1(2)+eastoffset&" "&q2(2)&" "&q3(2)&" "&q1(3)+eastoffset&" "&q2(3)&" "&q3(3)&" "&q1(4)+eastoffset&" "&q2(4)&" "&q3(4)&"</gml:posList>")
				else
					SessionOutput(indent & "<gml:LineString gml:id=""" & elementName & ".cu." & cuteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
					SessionOutput(indent & "  <gml:posList>60.02 10.1 60.02 10.3 60.03 10.2</gml:posList>")
				end if
				SessionOutput(indent & "</gml:LineString>")
		end if
		if geomtype = "GM_CompositeCurve" then
				cuteller = cuteller + 1
'				SessionOutput(indent & "<gml:Curve gml:id = """ & elementName & ".cc." & cuteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
				SessionOutput(indent & "<gml:CompositeCurve>")
				SessionOutput(indent & "  <gml:curveMember>")


				if epsg = 5972 then
					SessionOutput(indent & "    <gml:LineString gml:id=""" & elementName & ".cu." & cuteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
'					SessionOutput(indent & "     <gml:posList>568597.68 6662024.15 90.67 568497.68 6662024.15 90.67</gml:posList>")
					SessionOutput(indent & "    <gml:posList>"&q1(1)+eastoffset&" "&q2(1)&" "&q3(1)&" "&q1(2)+eastoffset&" "&q2(2)&" "&q3(2)&" "&q1(3)+eastoffset&" "&q2(3)&" "&q3(3)&" "&q1(4)+eastoffset&" "&q2(4)&" "&q3(4)&"</gml:posList>")
				else
					SessionOutput(indent & "    <gml:LineString gml:id=""" & elementName & ".cu." & cuteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
					SessionOutput(indent & "      <gml:posList>60.02 10.1 60.02 10.3 60.03 10.2</gml:posList>")
				end if
				SessionOutput(indent & "    </gml:LineString>")

				SessionOutput(indent & "  </gml:curveMember>")
				SessionOutput(indent & "</gml:CompositeCurve>")
'eller   <gml:CompositeCurve>
'          <gml:curveMember>
'            <gml:Curve gml:id = "PblTiltak.cc.1" srsName="http://www.opengis.net/def/crs/epsg/0/5972" srsDimension="3">
'              <gml:segments>
'                <gml:LineStringSegment>
'                 <gml:posList>568597.68 6662024.15 90.67 568497.68 6662024.15 90.67</gml:posList>
'                </gml:LineStringSegment>
'              </gml:segments>
'            </gml:Curve>
'          </gml:curveMember>
'        </gml:CompositeCurve>
		
		end if
		if geomtype = "Flate" or geomtype = "GM_Surface" then
'				SessionOutput(indent & "<gml:Surface gml:id = """ & elementName & ".su.1"" srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
				suteller = suteller + 1
				if epsg = 5972 then
'					SessionOutput(indent & "<gml:Polygon gml:id=""" & elementName & ".su." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
		'			call generateSurfaceExample(elementName, indent, height)
					call generatePolygonExample(elementName, indent, height)
'					SessionOutput(indent & "<gml:Polygon srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
				else
					SessionOutput(indent & "<gml:Polygon gml:id=""" & elementName & ".su." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
					SessionOutput(indent & "  <gml:exterior>")
					SessionOutput(indent & "    <gml:LinearRing>")
					SessionOutput(indent & "      <gml:posList>60.02 10.1 60.02 10.3 60.03 10.2 60.02 10.1</gml:posList>")
					SessionOutput(indent & "    </gml:LinearRing>")
					SessionOutput(indent & "  </gml:exterior>")
					SessionOutput(indent & "</gml:Polygon>")
				end if
'				if epsg = 5972 then
'					SessionOutput(indent & "      <gml:posList>568444.03 6661981.48 89.20 568506.41 6662009.49 91.20 568525.84 6661998.97 90.80 568529.64 6662001.85 91.00 568535.02 6662054.94 91.50 568476.33 6662067.85 90.50 568466.50 6662054.49 90.50 568444.03 6661981.48 89.20</gml:posList>")
'				end if

'				SessionOutput(indent & "</gml:Surface>")

'
'eller	<gml:CompositeSurface srsName="http://www.opengis.net/def/crs/epsg/0/4258" gml:id="Havneavsnitt.CS111111">
'	 		<gml:surfaceMember>
' 				<gml:Surface srsName="http://www.opengis.net/def/crs/epsg/0/4258" gml:id="Havneavsnitt.S111111">
' 					<gml:patches>
' 						<gml:PolygonPatch>
' 							<gml:exterior>
' 								<gml:Ring>
' 									<gml:curveMember xlink:href="Havneavsnittgrense.C444444"/>
' 								</gml:Ring>
' 							</gml:exterior>
' 						</gml:PolygonPatch>
' 					</gml:patches>
' 				</gml:Surface>
' 			</gml: surfaceMember >
' 		</gml:CompositeSurface>


		end if
		if geomtype = "GM_CompositeSurface" then
		
			call generateCompositeSurfaceExample(elementName, indent)
			


'eller delt mellom flere objekter
'        <gml:CompositeSurface gml:id = "PblTiltak.cs.0" srsName="http://www.opengis.net/def/crs/epsg/0/5972" srsDimension="3">
'          <gml:surfaceMember>
'            <gml:Surface gml:id = "PblTiltak.su.0" srsName="http://www.opengis.net/def/crs/epsg/0/5972" srsDimension="3">
'              <gml:patches>
'                <gml:PolygonPatch>
'                  <gml:exterior>
'                    <gml:Ring>
'                      <gml:curveMember xlink:href="#Tiltaksgrense.cc.2"/>
'                    </gml:Ring>
'                  </gml:exterior>
'                </gml:PolygonPatch>
'              </gml:patches>
'            </gml:Surface>
'          </gml:surfaceMember>
'        </gml:CompositeSurface>
				
				
				
		end if
		if geomtype = "GM_Solid" or geomtype = "GM_CompositeSolid" then
			'getSosiGeometritype = "NO GO"
			dim height
			height = 6.0
			call generateSolidExample(elementName, indent, height)
'''			call generateSOSISolidExample(elementName, indent, height)
		end if
		if geomtype = "GM_Object" or geomtype = "GM_Primitive" then
				obteller = obteller + 1
				SessionOutput(indent & "<gml:Point gml:id=""" & elementName & ".ob." & obteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
				SessionOutput(indent & "  <gml:pos>60.02 10.1</gml:pos>")
				SessionOutput(indent & "</gml:Point>")
		end if
end sub



function getSosiGeometritype(element)

		getSosiGeometritype = ""
		if element.Type = "Punkt" or element.Type = "GM_Point" then
			getSosiGeometritype = "PUNKT"
		end if
		if element.Type = "Sverm" or element.Type = "GM_MultiPoint" then
			getSosiGeometritype = "SVERM"
		end if
		if element.Type = "Kurve" or element.Type = "GM_Curve" or element.Type = "GM_CompositeCurve" then
			getSosiGeometritype = "KURVE,BUEP,KLOTOIDE"
		end if
		if element.Type = "Flate" or element.Type = "GM_Surface" or element.Type = "GM_CompositeSurface" then
			getSosiGeometritype = "FLATE"
		end if
		if element.Type = "GM_Solid" or element.Type = "GM_CompositeSolid" then
			getSosiGeometritype = "NO GO"
		end if
		if element.Type = "GM_Object" or element.Type = "GM_Primitive" then
			getSosiGeometritype = "PUNKT,SVERM,KURVE,BUEP,KLOTOIDE,FLATE"
		end if
end function


function getTaggedValue(element,taggedValueName)
		dim i, existingTaggedValue
		getTaggedValue = ""
		for i = 0 to element.TaggedValues.Count - 1
			set existingTaggedValue = element.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = taggedValueName then
				getTaggedValue = existingTaggedValue.Value
			end if
		next
end function

function getPackageTaggedValue(package,taggedValueName)
		dim i, existingTaggedValue
		getPackageTaggedValue = ""
		for i = 0 to package.element.TaggedValues.Count - 1
			set existingTaggedValue = package.element.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = taggedValueName then
				getPackageTaggedValue = existingTaggedValue.Value
			end if
		next
end function

function getConnectorEndTaggedValue(connectorEnd,taggedValueName)
	getConnectorEndTaggedValue = ""
	if not connectorEnd is nothing and Len(taggedValueName) > 0 then
		dim existingTaggedValue as EA.RoleTag 
		dim i
		for i = 0 to connectorEnd.TaggedValues.Count - 1
			set existingTaggedValue = connectorEnd.TaggedValues.GetAt(i)
			if existingTaggedValue.Tag = taggedValueName then
				getConnectorEndTaggedValue = existingTaggedValue.Value
			end if 
		next
	end if 
end function 

function getBasicSOSIType(umltype)
	getBasicSOSIType = "*"
	if umltype = "CharacterString" then
		getBasicSOSIType = "T"
	end if
	if umltype = "Boolean" then
		getBasicSOSIType = "BOOLSK"
	end if
	if umltype = "Date" then
		getBasicSOSIType = "DATO"
	end if
	if umltype = "DateTime" then
		getBasicSOSIType = "DATOTID"
	end if
	if umltype = "Integer" then
		getBasicSOSIType = "H"
	end if
	if umltype = "Real" then
		getBasicSOSIType = "D"
	end if
end function

function utf8(str)
	' make string utf-8
	Dim txt, res, tegn, utegn, vtegn, wtegn, xtegn, i
	
	utf8 = str
	exit function
	
    res = ""
	txt = Trim(str)
	' loop gjennom alle tegn
	For i = 1 To Len(txt)
		tegn = Mid(txt,i,1)

		'if      (c <    0x80) {  *out++=  c;                bits= -6; }
        'else if (c <   0x800) {  *out++= ((c >>  6) & 0x1F) | 0xC0;  bits=  0; }
        'else if (c < 0x10000) {  *out++= ((c >> 12) & 0x0F) | 0xE0;  bits=  6; }
        'else                  {  *out++= ((c >> 18) & 0x07) | 0xF0;  bits= 12; }

		if AscW(tegn) < 128 then
			res = res + tegn
		else if AscW(tegn) < 2048 then
			'u = AscW(tegn)
			'Repository.WriteOutput "Script", "tegn: " & AscW(tegn) & " " & Chr(AscW(tegn) / 64) & " " & int(u / 64),0
			'            c   229=E5/1110 0101
			'            c   192=C0/1100 0000  64=40/0100 0000
			utegn = Chr((int(AscW(tegn) / 64) or 192) )
			res = res + utegn
			'               c          63=3F/0011 1111
			vtegn = Chr((AscW(tegn) and 63) or 128)
			res = res + vtegn
			'            C3A5=å   195/1100 0011   165/1010 0101
			'Repository.WriteOutput "Script", "utf8: " & tegn & " -> " & utegn & " + " & vtegn,0
			'Repository.WriteOutput "Script", "int : " & AscW(tegn) & " -> " & Asc(utegn) & " + " & Asc(vtegn),0
		else if AscW(tegn) < 65536 then
			utegn = Chr((int(AscW(tegn) / 4096) or 224) )
			res = res + utegn
			vtegn = Chr((int(AscW(tegn) / 64) or 128) )
			res = res + vtegn
			wtegn = Chr((AscW(tegn) and 63) or 128)
			res = res + wtegn
			'putchar (0xE0 | c>>12);  E0=224, 2^12=4096
			'putchar (0x80 | c>>6 & 0x3F);  80=128, 2^6=64
			'putchar (0x80 | c & 0x3F);  80=128
		else if AscW(tegn) < 2097152 then	'/* 2^21 */
			utegn = Chr((int(AscW(tegn) / 262144) or 240) )
			res = res + utegn
			vtegn = Chr((int(AscW(tegn) / 4096) or 128) )
			res = res + vtegn
			wtegn = Chr((int(AscW(tegn) / 64) or 128) )
			res = res + wtegn
			xtegn = Chr((AscW(tegn) and 63) or 128)
			res = res + xtegn
			'putchar (0xF0 | c>>18);  F0=240, 2^18=262144
			'putchar (0x80 | c>>12 & 0x3F); 80=128, 2^12=4096
			'putchar (0x80 | c>>6 & 0x3F);  80=128, 2^6=64
			'putchar (0x80 | c & 0x3F);  80=128, 3F=63
		end if
		end if
		end if
		end if

	Next
	' return res
	utf8 = res

End function



sub listJsonFgGeometry(ftname,element,indent)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim element0 as EA.Element
	dim super as EA.Element
	dim datatype as EA.Element
	dim subbtype as EA.Element
	dim conn as EA.Collection
	dim connEnd as EA.ConnectorEnd
	dim i, umlnavn, sosinavn, sositype, sosilengde, sosimin, sosimax, sosierlik, koder, prikkniv1, roleEndElementID, sosidef, selfref, subID
	dim indent0, indent1, superlist
	dim geomtype
	dim height
	height = 6.0
'	eastoffset = eastoffset + eastdelta		
	if debug then Repository.WriteOutput "Script", "<!-- Debug: --------listDatatypes element.Name [" & element.Name & "] element.ElementID [" & element.ElementID & "]. -->",0
'ZZZ	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "" or LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
		
		for each conn in element.Connectors
			if debug then Repository.WriteOutput "Script", "<!-- Debug: conn.Type [" & conn.Type & "] conn.ClientID [" & conn.ClientID & "] conn.SupplierID [" & conn.SupplierID & "]. -->",0
			if conn.Type = "Generalization" then
				if element.ElementID = conn.ClientID then
					if debug then Repository.WriteOutput "Script", "<!-- Debug: supertype [" & Repository.GetElementByID(conn.SupplierID).Name & "]. -->",0
'					superlist = getSupertypes(ftname, conn.SupplierID, indent)
					set super = Repository.GetElementByID(conn.SupplierID)
					call listJsonFgGeometry(ftname,super,indent)
				end if
			end if
		next
'		if debug then Repository.WriteOutput "Script", "Debug: superlist [" & superlist & "].",0

'' TBD first find first geometry
		dim attr as EA.Attribute
		for each attr in element.Attributes
			if getSosiGeometritype(attr) <> "" then
'' TBD
				if gnum > 0 then
					SessionOutput("             ,")
				end if
				gnum = gnum + 1

				geomtype = getSosiGeometritype(attr)
				geomtype = attr.Type
		'		SessionOutput("geomtype "&geomtype) 
				if geomtype = "Punkt" or geomtype = "GM_Point" then
					pnteller = pnteller + 1
					SessionOutput("             ""type"": ""Point"",")
					SessionOutput("             ""coordinates"": [" &r1+eastoffset&", "&r2&", "&r3&"]")			
				end if
				if geomtype = "Sverm" or geomtype = "GM_MultiPoint" then
					pnteller = pnteller + 1
					SessionOutput("             ""type"": ""MultiPoint"",")
					SessionOutput("             ""coordinates"": [ [" &r1+eastoffset&", "&r2&", "&r3&"]")
					pnteller = pnteller + 1
					SessionOutput("             ],")
					SessionOutput("             [")
					SessionOutput("               [" &r1+eastoffset&", "&r2+12&", "&r3&"]")			
					SessionOutput("             ]")
				end if
				
				if geomtype = "Kurve" or geomtype = "GM_Curve" or geomtype = "GM_CompositeCurve" then
					cuteller = cuteller + 1
					SessionOutput("             ""type"": ""LineString"",")
					SessionOutput("             ""coordinates"": [")
					SessionOutput("               ["&q1(1)+eastoffset&", "&q2(1)&", "&q3(1)&"],") 
					SessionOutput("               ["&q1(2)+eastoffset&", "&q2(2)&", "&q3(2)&"],")
					SessionOutput("               ["&q1(3)+eastoffset&", "&q2(3)&", "&q3(3)&"],")
					SessionOutput("               ["&q1(4)+eastoffset&", "&q2(4)&", "&q3(4)&"]")
					SessionOutput("             ]")
				end if
				if geomtype = "Multikurve" or geomtype = "GM_MultiCurve" then
					cuteller = cuteller + 1
					SessionOutput("             ""type"": ""MultiLineString"",")
					SessionOutput("             ""coordinates"": [")
					SessionOutput("               [")
					SessionOutput("                 ["&q1(1)+eastoffset&", "&q2(1)&", "&q3(1)&"],") 
					SessionOutput("                 ["&q1(2)+eastoffset&", "&q2(2)&", "&q3(2)&"],")
					SessionOutput("                 ["&q1(3)+eastoffset&", "&q2(3)&", "&q3(3)&"],")
					SessionOutput("                 ["&q1(4)+eastoffset&", "&q2(4)&", "&q3(4)&"]")
					SessionOutput("               ],")				
					SessionOutput("               [")
					SessionOutput("                 ["&q1(1)+eastoffset&", "&q2(1)+12&", "&q3(1)&"],") 
					SessionOutput("                 ["&q1(2)+eastoffset&", "&q2(2)+12&", "&q3(2)&"],")
					SessionOutput("                 ["&q1(3)+eastoffset&", "&q2(3)+12&", "&q3(3)&"],")
					SessionOutput("                 ["&q1(4)+eastoffset&", "&q2(4)+12&", "&q3(4)&"]")
					SessionOutput("               ]")				
					SessionOutput("             ]")				
				end if
				
				if geomtype = "Flate" or geomtype = "GM_Surface" or geomtype = "GM_CompositeSurface" then
					suteller = suteller + 1
					SessionOutput("             ""type"": ""Polygon"",")
					SessionOutput("             ""coordinates"": [")
					call generatePolygonExample(element, indent, height)
					SessionOutput("             ]")
				end if
				if geomtype = "Multiflate" or geomtype = "GM_MultiSurface" then
					suteller = suteller + 1
					SessionOutput("             ""type"": ""MultiPolygon"",")
					SessionOutput("             ""coordinates"": [")
					call generatePolygonExample(element, indent + "  ", height)
					SessionOutput("             ],")
					suteller = suteller + 1
					SessionOutput("             [")
					call generatePolygonExample(element, indent + "  ", height + 12)
					SessionOutput("             ]")				
				end if
				
				if geomtype = "GM_Solid" then
					SessionOutput("             ""type"": ""Polyhedron""")
					SessionOutput("             ""coordinates"": [")
					SessionOutput("               [")
					call generatePolyhedronExample(element, indent, height)
					SessionOutput("               ]")
					SessionOutput("             ]")
  				end if
				if geomtype = "GM_MultiSolid" then
					SessionOutput("             ""type"": ""MultiPolyhedron""")
					SessionOutput("             ""coordinates"": [")
					SessionOutput("               [")
					call generatePolyhedronExample(element, indent, height)
					SessionOutput("               ],")
					suteller = suteller + 1
					SessionOutput("               [")
					call generatePolyhedronExample(element, indent, height + 12)
					SessionOutput("               ]")
					SessionOutput("             ]")
  				end if
				'Multi
				if geomtype = "GM_Object" then
					SessionOutput("             ""type"": ""GeometryCollectionTBD""")
					SessionOutput("             ""geometries"": [10.0, 60.5]")	
				end if

			end if
		next

	end if
	
end sub



sub listGeoJsonGeometry(ftname,element,indent)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim element0 as EA.Element
	dim super as EA.Element
	dim datatype as EA.Element
	dim subbtype as EA.Element
	dim conn as EA.Collection
	dim connEnd as EA.ConnectorEnd
	dim i, umlnavn, sosinavn, sositype, sosilengde, sosimin, sosimax, sosierlik, koder, prikkniv1, roleEndElementID, sosidef, selfref, subID
	dim indent0, indent1, superlist
	dim geomtype
	dim height
	height = 6.0
	
'	eastoffset = eastoffset + eastdelta		
	if debug then Repository.WriteOutput "Script", "<!-- Debug: --------listDatatypes element.Name [" & element.Name & "] element.ElementID [" & element.ElementID & "]. -->",0
'ZZZ	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "" or LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
		
		for each conn in element.Connectors
			if debug then Repository.WriteOutput "Script", "<!-- Debug: conn.Type [" & conn.Type & "] conn.ClientID [" & conn.ClientID & "] conn.SupplierID [" & conn.SupplierID & "]. -->",0
			if conn.Type = "Generalization" then
				if element.ElementID = conn.ClientID then
					if debug then Repository.WriteOutput "Script", "<!-- Debug: supertype [" & Repository.GetElementByID(conn.SupplierID).Name & "]. -->",0
'					superlist = getSupertypes(ftname, conn.SupplierID, indent)
					set super = Repository.GetElementByID(conn.SupplierID)
					call listGeoJsonGeometry(ftname,super,indent)
				end if
			end if
		next
'		if debug then Repository.WriteOutput "Script", "Debug: superlist [" & superlist & "].",0

'' TBD first find first geometry
		dim attr as EA.Attribute
		for each attr in element.Attributes
			if getSosiGeometritype(attr) <> "" then
'' TBD
				gnum = gnum + 1
				if gnum > 1 then
					exit sub
				end if
				geomtype = getSosiGeometritype(attr)
				geomtype = attr.Type
		'		SessionOutput("geomtype "&geomtype) 
				if geomtype = "Punkt" or geomtype = "GM_Point" then
					pnteller = pnteller + 1
					SessionOutput("             ""type"": ""Point"",")
					SessionOutput("             ""coordinates"": [" &r1+eastoffset&", "&r2&", "&r3&"]")			
				end if
				if geomtype = "Sverm" or geomtype = "GM_MultiPoint" then
					pnteller = pnteller + 1
					SessionOutput("             ""type"": ""MultiPoint"",")
					SessionOutput("             ""coordinates"": [ [" &r1+eastoffset&", "&r2&", "&r3&"]")
					pnteller = pnteller + 1
					SessionOutput("             ],")
					SessionOutput("             [")
					SessionOutput("               [" &r1+eastoffset&", "&r2+12&", "&r3&"]")			
					SessionOutput("             ]")
				end if
				
				if geomtype = "Kurve" or geomtype = "GM_Curve" or geomtype = "GM_CompositeCurve" then
					cuteller = cuteller + 1
					SessionOutput("             ""type"": ""LineString"",")
					SessionOutput("             ""coordinates"": [")
					SessionOutput("               ["&q1(1)+eastoffset&", "&q2(1)&", "&q3(1)&"],") 
					SessionOutput("               ["&q1(2)+eastoffset&", "&q2(2)&", "&q3(2)&"],")
					SessionOutput("               ["&q1(3)+eastoffset&", "&q2(3)&", "&q3(3)&"],")
					SessionOutput("               ["&q1(4)+eastoffset&", "&q2(4)&", "&q3(4)&"]")
					SessionOutput("             ]")
				end if
				if geomtype = "Multikurve" or geomtype = "GM_MultiCurve" then
					cuteller = cuteller + 1
					SessionOutput("             ""type"": ""MultiLineString"",")
					SessionOutput("             ""coordinates"": [")
					SessionOutput("               [")
					SessionOutput("                 ["&q1(1)+eastoffset&", "&q2(1)&", "&q3(1)&"],") 
					SessionOutput("                 ["&q1(2)+eastoffset&", "&q2(2)&", "&q3(2)&"],")
					SessionOutput("                 ["&q1(3)+eastoffset&", "&q2(3)&", "&q3(3)&"],")
					SessionOutput("                 ["&q1(4)+eastoffset&", "&q2(4)&", "&q3(4)&"]")
					SessionOutput("               ],")				
					SessionOutput("               [")
					SessionOutput("                 ["&q1(1)+eastoffset&", "&q2(1)+12&", "&q3(1)&"],") 
					SessionOutput("                 ["&q1(2)+eastoffset&", "&q2(2)+12&", "&q3(2)&"],")
					SessionOutput("                 ["&q1(3)+eastoffset&", "&q2(3)+12&", "&q3(3)&"],")
					SessionOutput("                 ["&q1(4)+eastoffset&", "&q2(4)+12&", "&q3(4)&"]")
					SessionOutput("               ]")				
					SessionOutput("             ]")				
				end if
				
				if geomtype = "Flate" or geomtype = "GM_Surface" or geomtype = "GM_CompositeSurface" then
					suteller = suteller + 1
					SessionOutput("             ""type"": ""Polygon"",")
					SessionOutput("             ""coordinates"": [")
					call generateGeoPolygonExample(element, indent, height)
					SessionOutput("             ]")
				end if
				if geomtype = "Multiflate" or geomtype = "GM_MultiSurface" then
					suteller = suteller + 1
					SessionOutput("             ""type"": ""MultiPolygon"",")
					SessionOutput("             ""coordinates"": [")
					call generateGeoPolygonExample(element, indent + "  ", height)
					SessionOutput("             ],")
					suteller = suteller + 1
					SessionOutput("             [")
					call generateGeoPolygonExample(element, indent + "  ", height + 12)
					SessionOutput("             ]")				
				end if
				
				if geomtype = "GM_Solid" or geomtype = "GM_MultiSolid" then
					gnum = gnum - 1
			'		SessionOutput("             ""type"": ""Polyhedron""")
			'		SessionOutput("             ""coordinates"": [")
			'		SessionOutput("               [")
			'		call generatePolyhedronExample(element, indent, height)
			'		SessionOutput("               ]")
			'		SessionOutput("             ],")
  				end if
				'Multi
				if geomtype = "GM_Object" then
					SessionOutput("             ""type"": ""GeometryCollection""")
					SessionOutput("             ""geometries"": [10.0, 60.5]")	
				end if

			end if
		next

	end if
	
end sub



sub generateCompositeSurfaceExample(elementName, indent)


	'			if epsg = 5972 then
				SessionOutput(indent & "<gml:CompositeSurface>")
 				SessionOutput(indent & "  <gml:surfaceMember>")
				if epsg = 5972 then
					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".northroof." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(1,0)+eastoffset&" "&p1(1,1)&" "&p1(1,2)&" "&p1(2,0)+eastoffset&" "&p1(2,1)&" "&p1(2,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(3,0)+eastoffset&" "&p1(3,1)&" "&p1(3,2)&" "&p1(4,0)+eastoffset&" "&p1(4,1)&" "&p1(4,2)&" "&p1(1,0)+eastoffset&" "&p1(1,1)&" "&p1(1,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")

					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".southroof1." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(2,0)+eastoffset&" "&p1(2,1)&" "&p1(2,2)&" "&p1(5,0)+eastoffset&" "&p1(5,1)&" "&p1(5,2)&" "&p1(6,0)+eastoffset&" "&p1(6,1)&" "&p1(6,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(2,0)+eastoffset&" "&p1(2,1)&" "&p1(2,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")

					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".southroof2." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(8,0)+eastoffset&" "&p1(8,1)&" "&p1(8,2)&" "&p1(3,0)+eastoffset&" "&p1(3,1)&" "&p1(3,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")

					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".westroof." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(6,0)+eastoffset&" "&p1(6,1)&" "&p1(6,2)&" "&p1(9,0)+eastoffset&" "&p1(9,1)&" "&p1(9,2)&" "&p1(10,0)+eastoffset&" "&p1(10,1)&" "&p1(10,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(6,0)+eastoffset&" "&p1(6,1)&" "&p1(6,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")

					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".eastroof." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(10,0)+eastoffset&" "&p1(10,1)&" "&p1(10,2)&" "&p1(11,0)+eastoffset&" "&p1(11,1)&" "&p1(11,2)&" "&p1(8,0)+eastoffset&" "&p1(8,1)&" "&p1(8,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(10,0)+eastoffset&" "&p1(10,1)&" "&p1(10,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")
		
				else
					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".su." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>60.02 10.1 60.02 10.3 60.03 10.2 60.02 10.1</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")
				end if
 				SessionOutput(indent & "  </gml:surfaceMember>")
				
				
				
				
 				SessionOutput(indent & "</gml:CompositeSurface>")

end sub

sub generateHouseCompositeSurfaceExample(elementName, indent)


	'			if epsg = 5972 then
				SessionOutput(indent & "<gml:CompositeSurface>")
 				SessionOutput(indent & "  <gml:surfaceMember>")
				if epsg = 5972 then
					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".northroof." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&b1(1,0)+eastoffset&" "&b1(1,1)&" "&b1(1,2)&" "&b1(2,0)+eastoffset&" "&b1(2,1)&" "&b1(2,2)&" "&b1(7,0)+eastoffset&" "&b1(7,1)&" "&b1(7,2)&" "&b1(3,0)+eastoffset&" "&b1(3,1)&" "&b1(3,2)&" "&b1(4,0)+eastoffset&" "&b1(4,1)&" "&b1(4,2)&" "&b1(1,0)+eastoffset&" "&b1(1,1)&" "&b1(1,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")
					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".northwall." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&b1(1,0)+eastoffset&" "&b1(1,1)&" "&b1(1,2)&" "&b1(4,0)+eastoffset&" "&b1(4,1)&" "&b1(4,2)&" "&b1(4,0)+eastoffset&" "&b1(4,1)&" "&b1(4,2)-5.00&" "&b1(1,0)+eastoffset&" "&b1(1,1)&" "&b1(1,2)-5.00&" "&b1(1,0)+eastoffset&" "&b1(1,1)&" "&b1(1,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")
					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".northwestwall." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&b1(1,0)+eastoffset&" "&b1(1,1)&" "&b1(1,2)&" "&b1(1,0)+eastoffset&" "&b1(1,1)&" "&b1(1,2)-5.00&" "&b1(5,0)+eastoffset&" "&b1(5,1)&" "&b1(5,2)-5.00&" "&b1(5,0)+eastoffset&" "&b1(5,1)&" "&b1(5,2)&" "&b1(2,0)+eastoffset&" "&b1(2,1)&" "&b1(2,2)&" "&b1(1,0)+eastoffset&" "&b1(1,1)&" "&b1(1,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")

					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".southroof1." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(2,0)+eastoffset&" "&p1(2,1)&" "&p1(2,2)&" "&p1(5,0)+eastoffset&" "&p1(5,1)&" "&p1(5,2)&" "&p1(6,0)+eastoffset&" "&p1(6,1)&" "&p1(6,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(2,0)+eastoffset&" "&p1(2,1)&" "&p1(2,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")

					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".southroof2." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(8,0)+eastoffset&" "&p1(8,1)&" "&p1(8,2)&" "&p1(3,0)+eastoffset&" "&p1(3,1)&" "&p1(3,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")

					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".westroof." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(6,0)+eastoffset&" "&p1(6,1)&" "&p1(6,2)&" "&p1(9,0)+eastoffset&" "&p1(9,1)&" "&p1(9,2)&" "&p1(10,0)+eastoffset&" "&p1(10,1)&" "&p1(10,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(6,0)+eastoffset&" "&p1(6,1)&" "&p1(6,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")

					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".eastroof." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>"&p1(10,0)+eastoffset&" "&p1(10,1)&" "&p1(10,2)&" "&p1(11,0)+eastoffset&" "&p1(11,1)&" "&p1(11,2)&" "&p1(8,0)+eastoffset&" "&p1(8,1)&" "&p1(8,2)&" "&p1(7,0)+eastoffset&" "&p1(7,1)&" "&p1(7,2)&" "&p1(10,0)+eastoffset&" "&p1(10,1)&" "&p1(10,2)&"</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")
		
				else
					suteller = suteller + 1
					SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".su." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
					SessionOutput(indent & "      <gml:exterior>")
					SessionOutput(indent & "        <gml:LinearRing>")
					SessionOutput(indent & "          <gml:posList>60.02 10.1 60.02 10.3 60.03 10.2 60.02 10.1</gml:posList>")
					SessionOutput(indent & "        </gml:LinearRing>")
					SessionOutput(indent & "      </gml:exterior>")
					SessionOutput(indent & "    </gml:Polygon>")
				end if
 				SessionOutput(indent & "  </gml:surfaceMember>")
				
				
				
				
 				SessionOutput(indent & "</gml:CompositeSurface>")

end sub

sub generateSurfaceExample(elementName, indent, height)

	dim c1(2), z1(2), h1, posNum, i
	h1 = height

	posNum = 8
	z1(0) =0.0
	z1(1) =0.0
	z1(2) =0.0
	
'	calculate the central point and mean height

	for i = 0 to posNum - 2
		s1(i,0) = s1(i,0) + eastoffset
		z1(0) = z1(0) + s1(i,0)
		z1(1) = z1(1) + s1(i,1)
		z1(2) = z1(2) + s1(i,2)
	next
	
	c1(0) = Round( z1(0) / (posNum - 1),2)
	c1(1) = Round( z1(1) / (posNum - 1),2)
	c1(2) = Round( z1(2) / (posNum - 1),2)
	
	soteller = soteller + 1	
	
	'				correct use of non-planar surface
		SessionOutput(indent & "<gml:Surface gml:id=""" & elementName & ".0612.202.27" & ".so." & soteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
		for i = 0 to posNum - 2
		
			suteller = suteller + 1	
			SessionOutput(indent & "  <gml:surfaceMember>")
			SessionOutput(indent & "    <gml:Polygon gml:id=""" & elementName & ".0612.202.27" & ".so."  & soteller & ".sh.1.su." & suteller & """>")
			SessionOutput(indent & "      <gml:exterior>")
			SessionOutput(indent & "        <gml:LinearRing>")
			
			SessionOutput(indent & "          <gml:posList>" & c1(0) & " " & c1(1) & " " & c1(2) & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2) & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & " " & c1(0) & " " & c1(1) & " " & c1(2) & "</gml:posList>")
			
			SessionOutput(indent & "        </gml:LinearRing>")
			SessionOutput(indent & "      </gml:exterior>")
			SessionOutput(indent & "    </gml:Polygon>")
			SessionOutput(indent & "  </gml:surfaceMember>")

		next
		SessionOutput(indent & "</gml:Surface>")

end sub

sub generatePolygonExample(element, indent, height)

	soteller = soteller + 1	
	SessionOutput("                 ["&s1(0,0)+eastoffset&", "&s1(0,1)&", "&s1(0,2)&"],")
	SessionOutput("                 ["&s1(1,0)+eastoffset&", "&s1(1,1)&", "&s1(1,2)&"],")
	SessionOutput("                 ["&s1(2,0)+eastoffset&", "&s1(2,1)&", "&s1(2,2)&"],")
	SessionOutput("                 ["&s1(3,0)+eastoffset&", "&s1(3,1)&", "&s1(3,2)&"],")
	SessionOutput("                 ["&s1(4,0)+eastoffset&", "&s1(4,1)&", "&s1(4,2)&"],") 
	SessionOutput("                 ["&s1(5,0)+eastoffset&", "&s1(5,1)&", "&s1(5,2)&"],") 
	SessionOutput("                 ["&s1(6,0)+eastoffset&", "&s1(6,1)&", "&s1(6,2)&"],") 
	SessionOutput("                 ["&s1(0,0)+eastoffset&", "&s1(0,1)&", "&s1(0,2)&"]")
	
end sub


sub generateGeoPolygonExample(element, indent, height)

	soteller = soteller + 1	
	SessionOutput("                 ["&s1(0,0)+eastoffset&", "&s1(0,1)&", "&s1(0,2)&"],")
	SessionOutput("                 ["&s1(1,0)+eastoffset&", "&s1(1,1)&", "&s1(1,2)&"],")
	SessionOutput("                 ["&s1(2,0)+eastoffset&", "&s1(2,1)&", "&s1(2,2)&"],")
	SessionOutput("                 ["&s1(3,0)+eastoffset&", "&s1(3,1)&", "&s1(3,2)&"],")
	SessionOutput("                 ["&s1(4,0)+eastoffset&", "&s1(4,1)&", "&s1(4,2)&"],") 
	SessionOutput("                 ["&s1(5,0)+eastoffset&", "&s1(5,1)&", "&s1(5,2)&"],") 
	SessionOutput("                 ["&s1(6,0)+eastoffset&", "&s1(6,1)&", "&s1(6,2)&"],") 
	SessionOutput("                 ["&s1(0,0)+eastoffset&", "&s1(0,1)&", "&s1(0,2)&"]")
	
end sub

sub generatePolyhedronExample(element, indent, height)

	dim c1(2), z1(2), h1, posNum, i
	h1 = height

	posNum = 8
	z1(0) =0.0
	z1(1) =0.0
	z1(2) =0.0
	
'	calculate the central point and mean height

'	for i = 0 to posNum - 2
	for i = 0 to posNum - 1
		s1(i,0) = s1(i,0) + eastoffset
		z1(0) = z1(0) + s1(i,0)
		z1(1) = z1(1) + s1(i,1)
		z1(2) = z1(2) + s1(i,2)
	next
	
	c1(0) = Round( z1(0) / (posNum - 1),2)
	c1(1) = Round( z1(1) / (posNum - 1),2)
	c1(2) = Round( z1(2) / (posNum - 1),2)
	soteller = soteller + 1	
		suteller = suteller + 1	
		
'	generate the floor tiles

	for i = 0 to posNum - 2	
		SessionOutput("                 [")
		SessionOutput("                   [")
		SessionOutput("                     [" & c1(0) & ", " & c1(1) & ", " & c1(2) & "],")
		SessionOutput("                     [" & s1(i+1,0) & ", " & s1(i+1,1) & ", " & s1(i+1,2) & "],") 
		SessionOutput("                     [" & s1(i,0) & ", " & s1(i,1) & ", " & s1(i,2) & "],") 
		SessionOutput("                     [" & c1(0) & ", " & c1(1) & ", " & c1(2) & "]")	
		SessionOutput("                   ]")
		SessionOutput("                 ],")
	next	
	
'	erect the sheet piles
	for i = 0 to posNum - 2	
		SessionOutput("                 [")
		SessionOutput("                   [")
		SessionOutput("                    [" & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & "],")
		SessionOutput("                     [" & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2) & "],")
		SessionOutput("                     [" & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2)+h1 & "],")
		SessionOutput("                     [" & s1(i,0) & " " & s1(i,1) & " " & s1(i,2)+h1 & "],")
		SessionOutput("                     [" & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & "]")
		SessionOutput("                   ]")
		SessionOutput("                 ],")
	next

'	generate the roof
	for i = 0 to posNum - 2	
		SessionOutput("                 [")
		SessionOutput("                   [")
		SessionOutput("                     [" & c1(0) & ", " & c1(1) & ", " & c1(2)+h1 & "],")
		SessionOutput("                     [" & s1(i+1,0) & ", " & s1(i+1,1) & ", " & s1(i+1,2)+h1 & "],") 
		SessionOutput("                     [" & s1(i,0) & ", " & s1(i,1) & ", " & s1(i,2)+h1 & "],") 
		SessionOutput("                     [" & c1(0) & ", " & c1(1) & ", " & c1(2)+h1 & "]")	
		SessionOutput("                   ]")
		if i < posNum - 2 then
		SessionOutput("                 ],")
		else
		SessionOutput("                 ]")
		end if
	next	

	
	
end sub

sub generateSolidExample(elementName, indent, height)

'
'	start with a small surface with different elevations in each coordinate position, and with no interiors

'	test whether the whole surface is in a single plane, and if so consider skipping the center point part(?)

'	split the surface in subsurfaces where it is possible to generate a central point thet has direct vision to all its perimeter points(?)

'	find the central point and the mean height

'	construct the set of floor surface slices from the central point to every two consecutive points on the perimeter 

'	erect a set of sheet piles from two and two perimeter points up the given height above the floor

'	copy the reverse of the floor as a roof and add the given height to it

'	.


'	hardcode a totally random surface to start with
'   srsName="urn:ogc:def:crs:EPSG::5972" srsDimension="3">
'	568444.03 6661981.48 89.20
'	568506.41 6662009.49 91.20
'	568525.84 6661998.97 90.80
'	568529.64 6662001.85 91.00
'	568535.02 6662054.94 91.50
'	568476.33 6662067.85 90.50
'	568466.50 6662054.49 90.50
'	568444.03 6661981.48 89.20
	dim c1(2), z1(2), h1, posNum, i
	h1 = height

	posNum = 8
	z1(0) =0.0
	z1(1) =0.0
	z1(2) =0.0
	
'	calculate the central point and mean height

'	for i = 0 to posNum - 2
	for i = 0 to posNum - 1
		s1(i,0) = s1(i,0) + eastoffset
		z1(0) = z1(0) + s1(i,0)
		z1(1) = z1(1) + s1(i,1)
		z1(2) = z1(2) + s1(i,2)
	next
	
	c1(0) = Round( z1(0) / (posNum - 1),2)
	c1(1) = Round( z1(1) / (posNum - 1),2)
	c1(2) = Round( z1(2) / (posNum - 1),2)
	
'	start the xml structure of the gml:Solid
	soteller = soteller + 1	
    SessionOutput(indent & "<gml:Solid gml:id=""" & elementName & ".0612.202.27" & ".so." & soteller & """")
    SessionOutput(indent & "  srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
    SessionOutput(indent & "  <gml:exterior>")
    SessionOutput(indent & "    <gml:Shell gml:id=""" & elementName & ".0612.202.27" & ".so." & soteller & ".sh.1""")
    SessionOutput(indent & "      srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")

'	generate the floor tiles
	
	for i = 0 to posNum - 2
	
		suteller = suteller + 1	
		SessionOutput(indent & "      <gml:surfaceMember>")
		SessionOutput(indent & "        <gml:Polygon gml:id=""" & elementName & ".0612.202.27" & ".so."  & soteller & ".sh.1.su." & suteller & """>")
		SessionOutput(indent & "          <gml:exterior>")
		SessionOutput(indent & "            <gml:LinearRing>")
		
		SessionOutput(indent & "              <gml:posList>" & c1(0) & " " & c1(1) & " " & c1(2) & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2) & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & " " & c1(0) & " " & c1(1) & " " & c1(2) & "</gml:posList>")
		
		SessionOutput(indent & "            </gml:LinearRing>")
		SessionOutput(indent & "          </gml:exterior>")
		SessionOutput(indent & "        </gml:Polygon>")
		SessionOutput(indent & "      </gml:surfaceMember>")

	next

'	erect the sheet piles

	for i = 0 to posNum - 2
	
		suteller = suteller + 1	
		SessionOutput(indent & "      <gml:surfaceMember>")
		SessionOutput(indent & "        <gml:Polygon gml:id=""" & elementName & ".0612.202.27" & ".so."  & soteller & ".sh.1.su." & suteller & """>")
		SessionOutput(indent & "          <gml:exterior>")
		SessionOutput(indent & "            <gml:LinearRing>")
		
		SessionOutput(indent & "              <gml:posList>" & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2) & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2)+h1 & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2)+h1 & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & "</gml:posList>")
		
		SessionOutput(indent & "            </gml:LinearRing>")
		SessionOutput(indent & "          </gml:exterior>")
		SessionOutput(indent & "        </gml:Polygon>")
		SessionOutput(indent & "      </gml:surfaceMember>")

	next
	
'	generate the roof

	for i = 0 to posNum - 2
	
		suteller = suteller + 1	
		SessionOutput(indent & "      <gml:surfaceMember>")
		SessionOutput(indent & "        <gml:Polygon gml:id=""" & elementName & ".0612.202.27" & ".so."  & soteller & ".sh.1.su." & suteller & """>")
		SessionOutput(indent & "          <gml:exterior>")
		SessionOutput(indent & "            <gml:LinearRing>")
		
		SessionOutput(indent & "              <gml:posList>" & c1(0) & " " & c1(1) & " " & c1(2)+h1 & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2)+h1 & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2)+h1 & " " & c1(0) & " " & c1(1) & " " & c1(2)+h1 & "</gml:posList>")
		
		SessionOutput(indent & "            </gml:LinearRing>")
		SessionOutput(indent & "          </gml:exterior>")
		SessionOutput(indent & "        </gml:Polygon>")
		SessionOutput(indent & "      </gml:surfaceMember>")

	next

'	end the xml structure of the gml:Solid
    SessionOutput(indent & "    </gml:Shell>")
    SessionOutput(indent & "  </gml:exterior>")
    SessionOutput(indent & "</gml:Solid>")

end sub


sub generateSOSISolidExample(elementName, indent, height)

'	start with a small surface with different elevations in each coordinate position, and with no interiors

'	test whether the whole surface is in a single plane, and if so consider skipping the center point part(?)

'	split the surface in subsurfaces where it is possible to generate a central point thet has direct vision to all its perimeter points(?)

'	find the central point and the mean height

'	construct the set of floor surface slices from the central point to every two consecutive points on the perimeter 

'	erect a set of sheet piles from two and two perimeter points up the given height above the floor

'	copy the reverse of the floor as a roof and add the given height to it

'	.


'	hardcode a totally random surface to start with
'   srsName="urn:ogc:def:crs:EPSG::5972" srsDimension="3">
'	568444.03 6661981.48 89.20
'	568506.41 6662009.49 91.20
'	568525.84 6661998.97 90.80
'	568529.64 6662001.85 91.00
'	568535.02 6662054.94 91.50
'	568476.33 6662067.85 90.50
'	568466.50 6662054.49 90.50
'	568444.03 6661981.48 89.20
	dim h1, posNum, i, c1(2), z1(2)
	h1 = height*100

	
	posNum = 8
	z1(0) =0.0
	z1(1) =0.0
	z1(2) =0.0
	
'	calculate the central point and mean height

	for i = 0 to posNum - 2
		z1(0) = z1(0) + s1(i,0)
		z1(1) = z1(1) + s1(i,1)
		z1(2) = z1(2) + s1(i,2)
	next
	
'	c1(0) = Round( z1(0) / (posNum - 1),2)
'	c1(1) = Round( z1(1) / (posNum - 1),2)
'	c1(2) = Round( z1(2) / (posNum - 1),2)
	c1(0) = Round( z1(0) / (posNum - 1),0)
	c1(1) = Round( z1(1) / (posNum - 1),0)
	c1(2) = Round( z1(2) / (posNum - 1),0)
	
'	start the xml structure of the gml:Solid
	soteller = soteller + 1	
		SessionOutput(".HODE")
		SessionOutput("..TRANSPAR")
		SessionOutput("...KOORDSYS 22 EUREF89 UTM")
		SessionOutput("...ORIGO-NØ 0 0")
		SessionOutput("...ENHET 0.01")
		SessionOutput("..OMRÅDE")
		SessionOutput("...MIN-NØ 6660981 568044")
		SessionOutput("...MAX-NØ 6663024 568997")
		SessionOutput("!...ENHET-H 0.01")
		SessionOutput("!...VERT-DATUM NN2000")
'	generate the floor tiles
	
	for i = 0 to posNum - 2
		suteller = suteller + 1	
		SessionOutput(".FLATE " & suteller & ":")	
		SessionOutput("..OBJTYPE " & elementName)	
		suteller = suteller + 1	
		SessionOutput("..REF :" & suteller)
		SessionOutput(".KURVE " & suteller & ":")	
		SessionOutput("..OBJTYPE " & elementName & "grense")	
		SessionOutput("..NØH")
		SessionOutput(c1(0) & " " & c1(1) & " " & c1(2)) 
		SessionOutput(s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2))
		SessionOutput(s1(i,0) & " " & s1(i,1) & " " & s1(i,2))
		SessionOutput(c1(0) & " " & c1(1) & " " & c1(2))


	next

'	erect the sheet piles

	for i = 0 to posNum - 2
	
		suteller = suteller + 1	
		SessionOutput(".FLATE " & suteller & ":")	
		SessionOutput("..OBJTYPE " & elementName)	
		suteller = suteller + 1	
		SessionOutput("..REF :" & suteller)
		SessionOutput(".KURVE " & suteller & ":")	
		SessionOutput("..OBJTYPE " & elementName & "grense")	
		SessionOutput("..NØH")
		SessionOutput(s1(i,0) & " " & s1(i,1) & " " & s1(i,2))
		SessionOutput(s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2))
		SessionOutput(s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2)+h1)
		SessionOutput(s1(i,0) & " " & s1(i,1) & " " & s1(i,2)+h1)
		SessionOutput(s1(i,0) & " " & s1(i,1) & " " & s1(i,2))


	next
	
'	generate the roof

	for i = 0 to posNum - 2
	
		suteller = suteller + 1	
		SessionOutput(".FLATE " & suteller & ":")	
		SessionOutput("..OBJTYPE " & elementName)	
		suteller = suteller + 1	
		SessionOutput("..REF :" & suteller)
		SessionOutput(".KURVE " & suteller & ":")	
		SessionOutput("..OBJTYPE " & elementName & "grense")	
		SessionOutput("..NØH")
		SessionOutput(c1(0) & " " & c1(1) & " " & c1(2)+h1)
		SessionOutput(s1(i,0) & " " & s1(i,1) & " " & s1(i,2)+h1)
		SessionOutput(s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2)+h1)
		SessionOutput(c1(0) & " " & c1(1) & " " & c1(2)+h1)


	next
	SessionOutput(".SLUTT")

end sub

sub SessionOutput(text)
	Session.Output(text)
end sub

sub getFeatureTypeCount(pkg)
 	dim currentElement as EA.Element
 	dim elements as EA.Collection 
	set elements = pkg.Elements
	dim i

	for i = 0 to elements.Count - 1
		set currentElement = elements.GetAt( i ) 
		if currentElement.Type = "Class" and LCase(currentElement.Stereotype) = "featuretype" and currentElement.Abstract = 0 then
			ftcount = ftcount + 1
		end if
	next
	
	dim subP as EA.Package
	for each subP in pkg.Packages
	    call getFeatureTypeCount(subP)
	next
	
end sub

function getFirstConcreteSubtypeName(datatype,subID)
	dim subber as EA.Element
'	dim datatype as EA.Element
	dim conn as EA.Collection
'	dim connEnd as EA.ConnectorEnd

	'subID = datatype.ElementID
	getFirstConcreteSubtypeName = "datatype.Name"
				
'	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
	if datatype.Abstract = 1 then
		if debug then Repository.WriteOutput "Script", "<!-- Debug: --------datatype.Name [" & datatype.Name & "] datatype.ElementID [" & datatype.ElementID & "]. -->",0
		for each conn in datatype.Connectors
			if debug then Repository.WriteOutput "Script", "<!-- Debug: conn.Type [" & conn.Type & "] conn.ClientID [" & conn.ClientID & "] conn.SupplierID [" & conn.SupplierID & "]. -->",0
			if conn.Type = "Generalization" then
				if datatype.ElementID <> conn.ClientID then
					if debug then Repository.WriteOutput "Script", "<!-- Debug: subbtype [" & Repository.GetElementByID(conn.ClientID).Name & "]. -->",0
					set subber = Repository.GetElementByID(conn.ClientID)
					getFirstConcreteSubtypeName =  getFirstConcreteSubtypeName(subber,subID)
					exit function
				end if
			end if
		next
	else
		getFirstConcreteSubtypeName = datatype.Name
		subID = datatype.ElementID
	end if
	
end function

function getFirstConcreteSubtypeID(datatype)
	dim subber as EA.Element
'	dim datatype as EA.Element
	dim conn as EA.Collection
'	dim connEnd as EA.ConnectorEnd
	dim subID

	subID = datatype.ElementID
	getFirstConcreteSubtypeID = 0
				
'	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
	if datatype.Abstract = 1 then
		if debug then Repository.WriteOutput "Script", "<!-- Debug: --------datatype.Name [" & datatype.Name & "] datatype.ElementID [" & datatype.ElementID & "]. -->",0
		call getFirstConcreteSubtypeName(datatype,subID)
	end if
	getFirstConcreteSubtypeID = subID
	
end function

function nao()
					' I just wanted a correct xml timestamp to document when the script was run
					dim m,d,t,min,sek,tm,td,tt,tmin,tsek
					m = Month(Date)
					if m < 10 then
						tm = "0" & FormatNumber(m,0,0,0,0)
					else
						tm = FormatNumber(m,0,0,0,0)
					end if
					d = Day(Date)
					if d < 10 then
						td = "0" & FormatNumber(d,0,0,0,0)
					else
						td = FormatNumber(d,0,0,0,0)
					end if
					t = Hour(Time)
					if t < 10 then
						tt = "0" & FormatNumber(t,0,0,0,0)
					else
						tt = FormatNumber(t,0,0,0,0)
					end if
					if t = 0 then tt = "00"
					min = Minute(Time)
					if min < 10 then
						tmin = "0" & FormatNumber(min,0,0,0,0)
					else
						tmin = FormatNumber(min,0,0,0,0)
					end if
					if min = 0 then tmin = "00"
					sek = Second(Time)
					if sek < 10 then
						tsek = "0" & FormatNumber(sek,0,0,0,0)
					else
						tsek = FormatNumber(sek,0,0,0,0)
					end if
					if sek = 0 then tsek = "00"
					'SessionOutput("  timeStamp=""" & Year(Date) & "-" & tm & "-" & td & "T" & tt & ":" & tmin & ":" & tsek & "Z""")
					nao = Year(Date) & "-" & tm & "-" & td & "T" & tt & ":" & tmin & ":" & tsek & "Z"
end function

sub initCoord()
	pnteller=0
	cuteller=0
	suteller=0
	soteller=0
	obteller=0
'	debug = false
	epsg = 5972
	eastdelta = 100.00
	eastoffset = 0.00
	'epsg = 4259
	'eastdelta = 0.0003 ?TBD
	
	p1(0,0) = 568000.00
	p1(0,1) = 6660000.00 
	p1(0,2) = 80.00

	p1(1,0) = 568491.54
	p1(1,1) = 6662044.12
	p1(1,2) = 94.48
	p1(2,0) = 568489.13
	p1(2,1) = 6662039.14 
	p1(2,2) = 97.70
	p1(3,0) = 568503.65
	p1(3,1) = 6662032.15
	p1(3,2) = 97.77
	p1(4,0) = 568506.07
	p1(4,1) = 6662037.43
	p1(4,2) = 94.50
	p1(5,0) = 568487.11
	p1(5,1) = 6662034.97
	p1(5,2) = 95.11
	p1(6,0) = 568493.32
	p1(6,1) = 6662031.98
	p1(6,2) = 95.10
	p1(7,0) = 568499.55
	p1(7,1) = 6662034.12
	p1(7,2) = 97.77
	p1(8,0) = 568501.71
	p1(8,1) = 6662028.04
	p1(8,2) = 95.07
	p1(9,0) = 568489.40
	p1(9,1) = 6662023.88
	p1(9,2) = 95.18
	p1(10,0) = 568493.58
	p1(10,1) = 6662021.85
	p1(10,2) = 97.77
	p1(11,0) = 568499.39
	p1(11,1) = 6662023.37
	p1(11,2) = 95.07
	p1(12,0) = 568491.32
	p1(12,1) = 6662023.64
	p1(12,2) = 95.48
	p1(13,0) = 568487.15
	p1(13,1) = 6662015.14
	p1(13,2) = 95.48
	p1(14,0) = 568490.11
	p1(14,1) = 6662013.64 
	p1(14,2) = 93.53
	p1(15,0) = 568494.22
	p1(15,1) = 6662022.27 
	p1(15,2) = 93.53
	p1(16,0) = 568488.52
	p1(16,1) = 6662024.94
	p1(16,2) = 93.53
	p1(17,0) = 568484.40
	p1(17,1) = 6662016.53
	p1(17,2) = 93.68

	q1(1) = 568525.84
	q2(1) = 6661998.97
	q3(1) = 90.80
	q1(2) = 568520.52
	q2(2) = 6662017.57
	q3(2) = 90.80
	q1(3) = 568517.55
	q2(3) = 6662029.69
	q3(3) = 90.80
	q1(4) = 568511.79
	q2(4) = 6662034.72
	q3(4) = 90.80
		
'	r1 = 568413.83
	r1 = 568513.83
	r2 = 6662030.36
	r3 = 90.67	
	
	s1(0,0) = 568444.03
	s1(0,1) = 6661981.48 
	s1(0,2) = 89.20
	s1(1,0) = 568506.41 
	s1(1,1) = 6662009.49 
	s1(1,2) = 91.20
	s1(2,0) = 568525.84 
	s1(2,1) = 6661998.97
	s1(2,2) = 90.80
	s1(3,0) = 568529.64 
	s1(3,1) = 6662001.85 
	s1(3,2) = 91.00
	s1(4,0) = 568535.02 
	s1(4,1) = 6662054.94 
	s1(4,2) = 91.50
	s1(5,0) = 568476.33 
	s1(5,1) = 6662067.85 
	s1(5,2) = 90.50
	s1(6,0) = 568466.50 
	s1(6,1) = 6662054.49 
	s1(6,2) = 90.50
	s1(7,0) = 568444.03 
	s1(7,1) = 6661981.48 
	s1(7,2) = 89.20
	
	b1(1,0) = 568492.07
	b1(1,1) = 6662041.11
	b1(1,2) = 96.00
	b1(2,0) = 568490.32
	b1(2,1) = 6662038.50
	b1(2,2) = 97.50
	b1(3,0) = 568502.81
	b1(3,1) = 6662032.53
	b1(3,2) = 97.50
	b1(4,0) = 568504.15
	b1(4,1) = 6662035.87
	b1(4,2) = 96.0
	b1(5,0) = 568489.11
	b1(5,1) = 6662035.60
	b1(5,2) = 96.00
	b1(6,0) = 568494.91
	b1(6,1) = 6662032.53
	b1(6,2) = 96.00
	b1(7,0) = 568499.55
	b1(7,1) = 6662034.12
	b1(7,2) = 97.50
	b1(8,0) = 568501.19
	b1(8,1) = 6662029.51
	b1(8,2) = 96.00
	b1(9,0) = 568492.12
	b1(9,1) = 6662026.50
	b1(9,2) = 96.00
	b1(10,0) = 568495.18
	b1(10,1) = 6662025.07
	b1(10,2) = 97.50
	b1(11,0) = 568498.19
	b1(11,1) = 6662023.81
	b1(11,2) = 96.00
	

	
	
	
	
end sub

listJsonExample
