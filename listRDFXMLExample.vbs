option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		listRDFXMLExample
' purpose:		Generates OWL/RDF/XML example objects from types in the source model according to a common national metamodel
' version:		2020-02-28
' version:		2020-03-18/19/20 Refaktorering med katalogdelen med informasjonsmodellen inline
' author:		Kent Jonsrud
' TODO:			list feature types 
' TODO:			mappe ut assosiasjonsroller
' TODO:			opprydding

		DIM debug, namespace, kortnavn, pnteller, cuteller, suteller, soteller, obteller, pversion, katalognavn
		debug = false

sub listRDFXMLExample()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"

	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		if Repository.GetTreeSelectedItemType() = otPackage then
			if UCASE(theElement.Element.Stereotype) = "APPLICATIONSCHEMA" then
				'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
						dim message, indent
			'	dim box
			'	box = Msgbox ("Script listGMLExample" & vbCrLf & vbCrLf & "Scriptversion 2018-09-16" & vbCrLf & "Listing to GML example for package : [" & theElement.Name & "].",1)
			'	select case box
			'	case vbOK
					dim xsdfile
					'tømmer System Output for lettere å fange opp hele RDF/XML-fila
					Repository.ClearOutput "Script"
					Repository.CreateOutputTab "Error"
					Repository.ClearOutput "Error"
					kortnavn = getPackageTaggedValue(theElement,"SOSI_kortnavn")
					if kortnavn = "" then
						kortnavn = theElement.Name
					'	Repository.WriteOutput "Script", "Pakken mangler tagged value SOSI_kortnavn! Kjører midlertidig videre med pakkenavnet som forslag til kortnavn: " & vbCrLf & kortnavn, 0
					end if

					pversion = getPackageTaggedValue(theElement,"version")
					namespace = getPackageTaggedValue(theElement,"targetNamespace")
					if namespace = "" then
						namespace = kortnavn
					end if
					
					namespace = InputBox("Please select the namespace name for the ontology.", "namespace", namespace)
					
					xsdfile = getPackageTaggedValue(theElement,"xsdDocument")
					if xsdfile = "" then
						xsdfile = kortnavn & ".xsd"
					end if
					'SessionOutput("test3: SOSI_kortnavn: " & vbCrLf & utf8(kortnavn) )
	'				SessionOutput("<?xml version=""1.0"" encoding=""utf-8""?>")
	'				SessionOutput("<wfs:FeatureCollection")
	'''				SessionOutput("  xmlns=""" & utf8(namespace) & """")
	'				SessionOutput("  xmlns:wfs=""http://www.opengis.net/wfs/2.0""")
	'				SessionOutput("  xmlns:gml=""http://www.opengis.net/gml/3.2""")
	'				SessionOutput("  xmlns:xlink=""http://www.w3.org/1999/xlink""")
	'				SessionOutput("  xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""")
	'				SessionOutput("  xsi:schemaLocation=""" & utf8(namespace))
					'SessionOutput("                     """ & namespace & "." & kortnavn & ".xsd""")
	'				SessionOutput("                     " & utf8(namespace) & "/" & utf8(xsdfile))
	'				SessionOutput("                     http://www.opengis.net/wfs/2.0")
	'				SessionOutput("                     http://schemas.opengis.net/wfs/2.0/wfs.xsd""")
					'SessionOutput("  timeStamp=""" & now & """")
					'SessionOutput("  timeStamp=""" & Year(Date) & "-" & FormatNumber(Month(Date),0,-1,0,0) & "-" & Day(Date) & "T" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "Z""")
					'SessionOutput("  timeStamp=""" & Year(Date) & "-" & LPad(Month(Date),"0",2) & "-" & Day(Date) & "T" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "Z""")


					' I will have a correct xml timestamp to document when the script was run
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
	'				SessionOutput("  timeStamp=""" & Year(Date) & "-" & tm & "-" & td & "T" & tt & ":" & tmin & ":" & tsek & "Z""")
	'				SessionOutput("  numberMatched=""unknown""")
	'				SessionOutput("  numberReturned=""0"">")
					'SessionOutput("  >")
					'SessionOutput("  <wfs:member>")
					pnteller=0
					cuteller=0
					suteller=0
					soteller=0
					obteller=0

  
					SessionOutput("<?xml version=""1.0"" encoding=""utf-8"" ?>")
					SessionOutput("<rdf:RDF xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""")
					SessionOutput("         xmlns:owl=""http://www.w3.org/2002/07/owl#""")
					SessionOutput("         xmlns:rdfs=""http://www.w3.org/2000/01/rdf-schema#""")
					SessionOutput("         xmlns:skos=""http://www.w3.org/2004/02/skos/core#""")
					SessionOutput("         xmlns:xsd=""http://www.w3.org/2001/XMLSchema#""")
					SessionOutput("         xmlns:dc=""http://purl.org/dc/terms/""")
					SessionOutput("         xmlns:dcat=""http://www.w3.org/ns/dcat#""")
					SessionOutput("         xmlns:foaf=""http://xmlns.com/foaf/0.1/""")
					SessionOutput("         xmlns:dcatno=""http://data.norge.no/informationmodel/""")
					SessionOutput("         xmlns:app=""" & utf8(namespace) & """>")
'?					SessionOutput("  xml:base=""" & utf8(namespace) & "/"">")


'					Katalogdelen 2020-03-18
					katalognavn = "http://sosi.geonorge.no/sosi-modellregister/katalog-over-objekttypebegreper"
					SessionOutput("  ")
					SessionOutput("  <!-- Kataloginformasjonen -->")
					
					SessionOutput("  <owl:NamedIndividual rdf:about=""" + katalognavn + """>")
					SessionOutput("    <rdf:type rdf:resource=""http://www.w3.org/ns/dcat#Catalog""/>")
					SessionOutput("    <dc:description xml:lang=""nb"">Katalog over elementer fra geomatikknorges UML-modeller</dc:description>")
'					SessionOutput("    <dc:publisher rdf:resource="https://raw.githubusercontent.com/Informasjonsforvaltning/model-publisher/master/src/model/model-catalog.ttl#Utgiver"/>
					SessionOutput("    <dc:publisher rdf:resource=""" + katalognavn + """/>")
					SessionOutput("    <dc:title xml:lang=""nb"">SOSI-modellregister</dc:title>")
					SessionOutput("    <dcat:record>")

'					SessionOutput("      <owl:NamedIndividual rdf:about="https://raw.githubusercontent.com/Informasjonsforvaltning/model-publisher/master/src/model/model-catalog.ttl#AdresseKatalogpost">
					SessionOutput("      <owl:NamedIndividual rdf:about=""" + katalognavn + "/" + utf8(kortnavn) + ".katalogpost""/>")
					SessionOutput("      <rdf:type rdf:resource=""http://www.w3.org/ns/dcat#CatalogRecord""/>")
'					SessionOutput("      <dc:identifier rdf:datatype=""http://www.w3.org/2001/XMLSchema#string"">abc123</dc:identifier>")
'					SessionOutput("      <dc:identifier rdf:datatype=""http://www.w3.org/2001/XMLSchema#string"">" + utf8(theElement.Element.FQName) + "</dc:identifier>")
					SessionOutput("      <dc:identifier rdf:datatype=""http://www.w3.org/2001/XMLSchema#string"">http://sosi.geonorge.no/informasjonsmodeller/" + utf8(theElement.Name) + ".owl.rdf.xml</dc:identifier>")
					SessionOutput("      <dc:title xml:lang=""nb"">" + "SOSI-modellregister" + "</dc:title>")
					SessionOutput("      <foaf:primaryTopic>")

'					SessionOutput("        <rdf:Description rdf:about=""https://raw.githubusercontent.com/Informasjonsforvaltning/model-publisher/master/src/model/model-catalog.ttl#AdresseModell">
'					SessionOutput("        <rdf:Description rdf:about=""" + katalognavn + "/" + kortnavn + "">
					SessionOutput("          <owl:NamedIndividual rdf:about=""" & utf8(namespace) & "/" & utf8(theElement.Name) & """>")
					SessionOutput("          <rdf:type rdf:resource=""http://data.norge.no/informationmodel/InformationModel""/>")
'					SessionOutput("          <rdf:type rdf:resource=""http://www.w3.org/2002/07/owl#NamedIndividual""/>")
					
					
'		?			call listFeatureTypeNames(theElement)
'					SessionOutput("          <ns0:containsModelElement rdf:resource="https://raw.githubusercontent.com/Informasjonsforvaltning/model-publisher/master/src/model/model-catalog.ttl#Adresse""/>")
'					SessionOutput("          <ns0:containsModelElement rdf:resource="https://raw.githubusercontent.com/Informasjonsforvaltning/model-publisher/master/src/model/model-catalog.ttl#GeografiskAdresse"/>
'					SessionOutput("          <ns0:containsModelElement>

'					SessionOutput("           <ns0:ModelElement rdf:about="https://raw.githubusercontent.com/Informasjonsforvaltning/model-publisher/master/src/model/model-catalog.ttl#Matrikkeladresse">
'					SessionOutput("             <rdf:type rdf:resource="http://www.w3.org/2002/07/owl#NamedIndividual"/>


					SessionOutput("  ")
					SessionOutput("        <!-- Pakken med informasjonsmodellen -->")
					
'?					
'					SessionOutput("  <owl:Ontology rdf:about=""" & utf8(namespace) & "/" & utf8(theElement.Name) & """>")
'					SessionOutput("      <rdfs:label>" & utf8(trimDefinitionText(theElement.Notes)) & "</rdfs:label>")
'					SessionOutput("      <dc:source>" & theElement.Element.FQName & "</dc:source>")
'					SessionOutput("      <owl:versionInfo>" & utf8(pversion) & "</owl:versionInfo>")
'					SessionOutput("      <owl:imports rdf:resource=""http://data.norge.no/informationmodel""/>")
'					SessionOutput("  </owl:Ontology>")
					
					
'					SessionOutput("  <owl:NamedIndividual rdf:about=""" & utf8(namespace) & "/" & utf8(theElement.Name) & """>")
'					SessionOutput("    <rdf:type rdf:resource=""http://data.norge.no/informationmodel/InformationModel""/>")

'					namespace = namespace & "/" & theElement.Name

'					SessionOutput("    <dcatno:containsModelElement rdf:resource=" TODO liste over stien til alle FT, DT, CL, Enum. (Union?)
'					call listFeatureTypeNames(theElement)

'					SessionOutput("    <dcatno:description xml:lang=""nb"">" & utf8(trimDefinitionText(theElement.Notes)) & "</dcatno:description>")
'					SessionOutput("    <dcatno:name xml:lang=""nb"">" & utf8(kortnavn) & "</dcatno:name>")
'					SessionOutput("    <dcatno:versjon rdf:datatype=""http://www.w3.org/2001/XMLSchema#string"">" & utf8(pversion) & "</dcatno:versjon>")
'					SessionOutput("  </owl:NamedIndividual>")
					

					call listFeatureTypes(theElement)

					SessionOutput("        <!-- Slutt pakken med informasjonsmodellen -->")


					SessionOutput("        </owl:NamedIndividual>")
					SessionOutput("      </foaf:primaryTopic>")
					SessionOutput("    </dcat:record>")
					SessionOutput("  </owl:NamedIndividual>")

					SessionOutput("</rdf:RDF>")

					

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
					namespace = "http://some.server.no/namespace"
					kortnavn = "shortNamespace"
					pnteller=0
					cuteller=0
					suteller=0
					soteller=0
					obteller=0
					SessionOutput("  <wfs:member>")
					SessionOutput("    <" & utf8(theElement.Name) & " gml:id="""& utf8(theElement.Name) & ".1"">")
					indent = "      "
					
					call listDatatypes(theElement.Name,theElement,indent)
					
					SessionOutput("    </" & utf8(theElement.Name) & ">")
					SessionOutput("  </wfs:member>")
				else
					'Other than «ApplicationSchema» Package or a «FeatureType» Class selected in the tree
					MsgBox( "This script requires a «ApplicationSchema» Package or a «FeatureType» Class to be selected in the Project Browser." & vbCrLf & _
					"Please select a «ApplicationSchema» Package or a «FeatureType» Class in the Project Browser and try again." )
				end if
			else
				'Other than «ApplicationSchema» Package or a «FeatureType» Class selected in the tree
				MsgBox( "Element type selected: " & theElement.Type & vbCrLf & _
				"This script requires a «ApplicationSchema» Package or a «FeatureType» Class to be selected in the Project Browser." & vbCrLf & _
				"Please select a «ApplicationSchema» Package or a «FeatureType» Class in the Project Browser and try again." )
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
	indent = "        "
	if debug then Repository.WriteOutput "Script", "Debug: pkg.Name [" & pkg.Name & "].",0
	SessionOutput(" ")
	SessionOutput(indent + "  <!-- Pakkenavn " & utf8(pkg.Name) & "  -->")
	for i = 0 to elements.Count - 1 
		dim currentElement as EA.Element 
		set currentElement = elements.GetAt( i ) 
				
		if debug then Repository.WriteOutput "Script", "Debug: currentElement.Name [«" & currentElement.Stereotype & "» " & currentElement.Name & "] currentElement.Type [" & currentElement.Type & "] currentElement.Abstract [" & currentElement.Abstract & "].",0
	'	if currentElement.Type = "Class" and LCase(currentElement.Stereotype) = "featuretype" and currentElement.Abstract = 0 then
		if currentElement.Type = "Class" and LCase(currentElement.Stereotype) = "featuretype" or LCase(currentElement.Stereotype) = "datatype" or LCase(currentElement.Stereotype) = "union" then

			SessionOutput(" ")

			if  LCase(currentElement.Stereotype) = "featuretype"  then
				SessionOutput(indent + "  <!-- Objekttype -->")
			end if
			if  LCase(currentElement.Stereotype) = "datatype"  then
				SessionOutput(indent + "  <!-- Datatype -->")
			end if
			if  LCase(currentElement.Stereotype) = "union"  then
				SessionOutput(indent + "  <!-- Union (ikke handtert i metamodellen) -->")
			end if
			SessionOutput(indent + "  <owl:NamedIndividual rdf:about=""" & utf8(namespace) & "/" & utf8(currentElement.Name) & """>")
			SessionOutput(indent + "    <rdf:type rdf:resource=""http://data.norge.no/informationmodel/ModelElement""/>")
			
			ftname = currentElement.Name
			superlist = ""
			indent = "      "
			for each conn in currentElement.Connectors
				if conn.Type = "Generalization" then
					if currentElement.ElementID = conn.ClientID then
						superlist = getSupertypes(conn.SupplierID)
					end if
				end if
			next
			if superlist <> "" then
				SessionOutput(indent + "    <dcatno:isSubtypeOf rdf:resource=""" & utf8(namespace) & "/" & utf8(superlist) & """/>")
			end if
			
			call listDatatypes(ftname,currentElement,indent)

			SessionOutput(" ")
			SessionOutput(indent + "    <dcatno:description xml:lang=""nb"">" & utf8(trimDefinitionText(currentElement.Notes)) & "</dcatno:description>")
			if  LCase(currentElement.Stereotype) = "featuretype"  then
				SessionOutput(indent + "    <dcatno:modelElementType rdf:datatype=""http://www.w3.org/2001/XMLSchema#string"">objekttype</dcatno:modelElementType>")
			end if
			if  LCase(currentElement.Stereotype) = "datatype"  then
				SessionOutput(indent + "    <dcatno:modelElementType rdf:datatype=""http://www.w3.org/2001/XMLSchema#string"">datatype</dcatno:modelElementType>")
			end if
			if  LCase(currentElement.Stereotype) = "union"  then
				SessionOutput(indent + "    <dcatno:modelElementType rdf:datatype=""http://www.w3.org/2001/XMLSchema#string"">datatype</dcatno:modelElementType>")
			end if
'    <dcatno:isDescribedBy rdf:datatype="http://www.w3.org/2001/XMLSchema#anyURI">http://begrepskatalog/begrep/20b2e302-9fe1-11e5-a9f8-e4115b280940</dcatno:isDescribedBy>
			SessionOutput(indent + "    <dcatno:name xml:lang=""nb"">" & utf8(currentElement.Name) & "</dcatno:name>")
			SessionOutput(indent + "  </owl:NamedIndividual>")

		end if
		
		if currentElement.Type = "Enumeration" or currentElement.Type = "Class" and LCase(currentElement.Stereotype) = "codelist" or LCase(currentElement.Stereotype) = "enumeration" then
			


 			SessionOutput(" ")
 
			SessionOutput(indent + "  <!-- Kodeliste -->")

			SessionOutput(indent + "  <owl:NamedIndividual rdf:about=""" & utf8(namespace) & "/" & utf8(currentElement.Name) & """>")
			SessionOutput(indent + "    <rdf:type rdf:resource=""http://data.norge.no/informationmodel/CodeList""/>")
			if getTaggedValue(currentElement,"asDictionary") = "true" and getTaggedValue(currentElement,"codeList") <> "" then
				SessionOutput(indent + "    <dcatno:isDescribedBy rdf:datatype=""http://www.w3.org/2001/XMLSchema#anyURI"">" & getTaggedValue(currentElement,"codeList") & "</dcatno:isDescribedBy>")			
			end if
			
			ftname = currentElement.Name
			superlist = ""
			indent = "      "
			for each conn in currentElement.Connectors
				if conn.Type = "Generalization" then
					if currentElement.ElementID = conn.ClientID then
						superlist = getSupertypes(conn.SupplierID)
					end if
				end if
			next

			call listKodeliste(ftname,currentElement,indent)
						
			SessionOutput(" ")
			SessionOutput(indent + "    <dcatno:description xml:lang=""nb"">" & utf8(trimDefinitionText(currentElement.Notes)) & "</dcatno:description>")
			SessionOutput(indent + "    <dcatno:modelElementType rdf:datatype=""http://www.w3.org/2001/XMLSchema#string"">kodeliste</dcatno:modelElementType>")
'    <dcatno:isDescribedBy rdf:datatype="http://www.w3.org/2001/XMLSchema#anyURI">http://begrepskatalog/begrep/20b2e302-9fe1-11e5-a9f8-e4115b280940</dcatno:isDescribedBy>
			SessionOutput(indent + "    <dcatno:name xml:lang=""nb"">" & utf8(currentElement.Name) & "</dcatno:name>")
			SessionOutput(indent + "  </owl:NamedIndividual>")

		end if
	
	next

	dim subP as EA.Package
	for each subP in pkg.packages
	    call listFeatureTypes(subP)
	next


end sub


sub listDatatypes(ftname,element,indent)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim super as EA.Element
	dim datatype as EA.Element
	dim conn as EA.Collection
	dim connEnd as EA.ConnectorEnd
	dim i, umlnavn, sosinavn, sositype, sosilengde, sosimin, sosimax, sosierlik, koder, prikkniv1, roleEndElementID, sosidef, selfref
	dim indent0, indent1, superlist
	
				
	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then

		if debug then Repository.WriteOutput "Script", "Debug: --------listDatatypes element.Name [" & element.Name & "] element.ElementID [" & element.ElementID & "].",0
	'	for each conn in element.Connectors
	'		if debug then Repository.WriteOutput "Script", "Debug: conn.Type [" & conn.Type & "] conn.ClientID [" & conn.ClientID & "] conn.SupplierID [" & conn.SupplierID & "].",0
	'		if conn.Type = "Generalization" then
	'			if element.ElementID = conn.ClientID then
	'				if debug then Repository.WriteOutput "Script", "Debug: supertype [" & Repository.GetElementByID(conn.SupplierID).Name & "].",0
'	'				superlist = getSupertypes(ftname, conn.SupplierID, indent)
	'				set super = Repository.GetElementByID(conn.SupplierID)
	'				call listDatatypes(ftname,super,indent)
	'			end if
	'		end if
	'	next
		'if debug then Repository.WriteOutput "Script", "Debug: superlist [" & superlist & "].",0

		i = 0
		dim attr as EA.Attribute
		for each attr in element.Attributes
			i = i + 1
			SessionOutput(" ")
			SessionOutput(indent + "    <dcatno:hasProperty>")			
			SessionOutput(indent + "      <owl:NamedIndividual rdf:about=""" & utf8(namespace) & "/" & utf8(ftname) & "/" & utf8(attr.Name) & """>")
			SessionOutput(indent + "      <rdf:type rdf:resource=""http://data.norge.no/informationmodel/Property""/>")
			SessionOutput(indent + "      <dcatno:type rdf:resource=""" & utf8(attrType(namespace,attr.Type)) & """/>")
			SessionOutput(indent + "      <dcatno:name xml:lang=""nb"">" & utf8(attr.Name) & "</dcatno:name>")
			SessionOutput(indent + "      <dcatno:propertyType>attributt</dcatno:propertyType>")
			if attr.UpperBound = "*" then
				SessionOutput(indent + "      <xsd:maxOccurs rdf:datatype=""http://www.w3.org/2000/01/rdf-schema#Literal"">*</xsd:maxOccurs>")
			else
				SessionOutput(indent + "      <xsd:maxOccurs rdf:datatype=""http://www.w3.org/2001/XMLSchema#nonNegativeInteger"">" & attr.UpperBound & "</xsd:maxOccurs>")
			end if
			SessionOutput(indent + "      <xsd:minOccurs rdf:datatype=""http://www.w3.org/2001/XMLSchema#nonNegativeInteger"">" & attr.LowerBound & "</xsd:minOccurs>")
			SessionOutput(indent + "      <dcatno:description xml:lang=""nb"">" & utf8(trimDefinitionText(attr.Notes)) & "</dcatno:description>")

			SessionOutput(indent + "      </owl:NamedIndividual>")	
			SessionOutput(indent + "    </dcatno:hasProperty>")			
			
			
			
			if false then
	'f		if getSosiGeometritype(attr) = "" then
				if debug then Repository.WriteOutput "Script", "Debug: attr.Name [" & attr.Name & "] not geometry.",0
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
								SessionOutput(indent & "<" & attr.Name & " xlink:href=""" & namespace & "/" & attr.Type & "/" & listCodeType(datatype) & """/>")
								'SessionOutput(indent & "<" & attr.Name & " xlink:href=""" & listReferenceType(attr.Type) & """/>")
								if attr.UpperBound <> "1" then
								'	SessionOutput(indent & "<" & attr.Name & ">" & listCodeType(datatype) & "</" & attr.Name & ">")
									SessionOutput(indent & "<" & attr.Name & " xlink:href=""" & namespace & "/" & attr.Type & "/" & listCodeType(datatype) & """/>")
								end if
							else
								'variant gml:CodeType
								SessionOutput(indent & "<" & attr.Name & ">" & listCodeType(datatype) & "</" & attr.Name & ">")
								if attr.UpperBound <> "1" then
									SessionOutput(indent & "<" & attr.Name & ">" & listCodeType(datatype) & "</" & attr.Name & ">")
								end if
							end if
							'listCodeType(attr)
						else
							SessionOutput(indent & "<" & utf8(attr.Name) & ">")
							indent0 = indent & "  "
							SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
							indent1 = indent0 & "  "
							call listDatatypes(ftname, datatype,indent1)
							SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
							SessionOutput(indent & "</" & utf8(attr.Name) & ">")
							if attr.UpperBound <> "1" then
								' write a second instance of the attribute, currently with exactly same content
								' but should be made to pick a different value or the second code (TBD)
								SessionOutput(indent & "<" & utf8(attr.Name) & ">")
								indent0 = indent & "  "
								SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
								indent1 = indent0 & "  "
								call listDatatypes(ftname, datatype,indent1)
								SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
								SessionOutput(indent & "</" & utf8(attr.Name) & ">")
							end if

						end if
					end if
				else
					'base type
					SessionOutput(indent & "<" & utf8(attr.Name) & ">" & listBaseType(ftname, attr.Name,attr.Type) & "</" & utf8(attr.Name) & ">")
					if attr.UpperBound <> "1" then
						SessionOutput(indent & "<" & utf8(attr.Name) & ">" & listBaseType(ftname, attr.Name,attr.Type) & "</" & utf8(attr.Name) & ">")
					end if


				end if
	'f		else
				'geometry type 
				if debug then Repository.WriteOutput "Script", "Debug: attr.Name [" & attr.Name & "] is geometry: " & getSosiGeometritype(attr) & ".",0
				SessionOutput(indent & "<" & utf8(attr.Name) & ">")
				call listGeometryType(ftname, attr.Type, indent & "  ")			
				SessionOutput(indent & "</" & utf8(attr.Name) & ">")
				if attr.UpperBound <> "1" then
					SessionOutput(indent & "<" & utf8(attr.Name) & ">")
					call listGeometryType(ftname, attr.Type, indent & "  ")			
					SessionOutput(indent & "</" & utf8(attr.Name) & ">")
				end if
			end if

			'if Union then jump out of the loop after first(!) variant, this does not support well Unions having several different datatypes 
			if LCase(element.Stereotype) = "union" then
				Exit For
			end if
			'SessionOutput(indent & "</" & attr.Name & ">")
		next
			
		for each conn in element.Connectors
			if false then
'			if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then

'			else
				'Repository.WriteOutput "Script", "Debug: Supplier Role.Name [" & conn.SupplierEnd.Role & "] datatypens SOSI_navn [" & getTaggedValue(Repository.GetElementByID(conn.ClientID).Name,"SOSI_navn") & "].",0
				'Repository.WriteOutput "Script", "Debug: Client Role.Name [" & conn.ClientEnd.Role & "] datatypens SOSI_navn [" & getTaggedValue(Repository.GetElementByID(conn.ClientID).Name,"SOSI_navn") & "].",0
				if debug then Repository.WriteOutput "Script", "Debug: Supplier Role.Name [" & conn.SupplierEnd.Role & "] datatypens SOSI_navn [" & Repository.GetElementByID(conn.SupplierID).Name & "].",0
				if debug then Repository.WriteOutput "Script", "Debug: Client Role.Name [" & conn.ClientEnd.Role & "] datatypens SOSI_navn [" & Repository.GetElementByID(conn.ClientID).Name & "].",0

				if conn.ClientID = element.ElementID then
					if getConnectorEndTaggedValue(conn.SupplierEnd,"xsdEncodingRule") <> "notEncoded" then
						set datatype = Repository.GetElementByID(conn.SupplierID)
						umlnavn = conn.SupplierEnd.Role
						if conn.ClientEnd.Aggregation = 2 then
							'composition+mandatory->nest as datatype inline?
							SessionOutput(indent & "<" & utf8(umlnavn) & ">")
							indent0 = indent & "  "
							SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
							indent1 = indent0 & "  "
							call listDatatypes(ftname, datatype,indent1)
							SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
							SessionOutput(indent & "</" & utf8(umlnavn) & ">")
							if conn.SupplierEnd.Cardinality <> "0..1" and conn.SupplierEnd.Cardinality <> "1..1" and conn.SupplierEnd.Cardinality <> "1" then
								SessionOutput(indent & "<" & utf8(umlnavn) & ">")
								indent0 = indent & "  "
								SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
								indent1 = indent0 & "  "
								call listDatatypes(ftname, datatype,indent1)
								SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
							SessionOutput(indent & "</" & utf8(umlnavn) & ">")
							end if
						else
							if conn.SupplierEnd.Navigable = "Navigable" then
								'self assoc? if so make xlinks to other (imaginary) instances of the same class
								selfref = 1
								if datatype.Name = element.Name and datatype.ElementID = element.ElementID then
									selfref = 2
								end if 
								'navigable->make xlink? 
								SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(datatype.Name) & "." & selfref & """/>")
								if debug then Repository.WriteOutput "Script", "Debug: SupplierEnd.Cardinality [" & conn.SupplierEnd.Cardinality & "].",0
								if conn.SupplierEnd.Cardinality <> "0..1" and conn.SupplierEnd.Cardinality <> "1..1" and conn.SupplierEnd.Cardinality <> "1" then
									SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(datatype.Name) & "." & selfref + 1 & """/>")
								end if
							end if
						end if
					end if
				else
					if getConnectorEndTaggedValue(conn.ClientEnd,"xsdEncodingRule") <> "notEncoded" then
						set datatype = Repository.GetElementByID(conn.ClientID)
						umlnavn = conn.ClientEnd.Role
						if conn.SupplierEnd.Aggregation = 2 then
							'composition+mandatory->nest as datatype inline?
							SessionOutput(indent & "<" & utf8(umlnavn) & ">")
							indent0 = indent & "  "
							SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
							indent1 = indent0 & "  "
							call listDatatypes(ftname, datatype,indent1)
							SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
							SessionOutput(indent & "</" & utf8(umlnavn) & ">")
							if conn.ClientEnd.Cardinality <> "0..1" and conn.ClientEnd.Cardinality <> "1..1" and conn.ClientEnd.Cardinality <> "1" then
								SessionOutput(indent & "<" & utf8(umlnavn) & ">")
								indent0 = indent & "  "
								SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
								indent1 = indent0 & "  "
								call listDatatypes(ftname, datatype,indent1)
								SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
								SessionOutput(indent & "</" & utf8(umlnavn) & ">")
							end if
						else
							if conn.ClientEnd.Navigable = "Navigable" then
								'self assoc? if so make xlinks to other (imaginary) instances of the same class
								selfref = 1
								if datatype.Name = element.Name and datatype.ElementID = element.ElementID then
									selfref = 2
								end if 
								'navigable->make xlink? 
								SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(datatype.Name) & "." & selfref & """/>")
								if debug then Repository.WriteOutput "Script", "Debug: ClientEnd.Cardinality [" & conn.ClientEnd.Cardinality & "].",0
								if conn.ClientEnd.Cardinality <> "0..1" and conn.ClientEnd.Cardinality <> "1..1" and conn.ClientEnd.Cardinality <> "1" then
									SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(datatype.Name) & "." & selfref + 1 & """/>")
								end if
							end if
						end if
					end if
				end if

			end if

		next

	end if

end sub


sub listKodeliste(ftname,element,indent)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim super as EA.Element
	dim datatype as EA.Element
	dim conn as EA.Collection
	dim connEnd as EA.ConnectorEnd
	dim i, umlnavn, sosinavn, sositype, sosilengde, sosimin, sosimax, sosierlik, koder, prikkniv1, roleEndElementID, sosidef, selfref
	dim indent0, indent1, superlist
	
		if debug then Repository.WriteOutput "Script", "Debug: --------listDatatypes element.Name [" & element.Name & "] element.ElementID [" & element.ElementID & "].",0

		i = 0
		dim attr as EA.Attribute
		for each attr in element.Attributes
			i = i + 1
			SessionOutput(" ")
			SessionOutput(indent + "    <dcatno:containsCodename>")
			SessionOutput(indent + "      <owl:NamedIndividual rdf:about=""" & utf8(namespace) & "/" & utf8(ftname) & "/" & utf8(attr.Name) & """>")
			SessionOutput(indent + "        <rdf:type rdf:resource=""http://data.norge.no/informationmodel/Codename""/>")
			SessionOutput(indent + "        <dcatno:name xml:lang=""nb"">" & utf8(attr.Name) & "</dcatno:name>")
			SessionOutput(indent + "        <dcatno:description xml:lang=""nb"">" & utf8(trimDefinitionText(attr.Notes)) & "</dcatno:description>")
			SessionOutput(indent + "      </owl:NamedIndividual>")
			SessionOutput(indent + "    </dcatno:containsCodename>")		
		next


end sub


function getSupertypes(elementID)
	dim super as EA.Element
	dim conn as EA.Collection
	dim supername, supernames
	set super = Repository.GetElementByID(elementID)
'	for each conn in super.Connectors
'		if conn.Type = "Generalization" then
'			if super.ElementID = conn.ClientID then
'				supername = getSupertypes(conn.SupplierID)
'			end if
'		end if
'	next
	if debug then Repository.WriteOutput "Script", "Debug: super.Name [" & super.Name & "]  supername [" & supername & "].",0
'	getSupertypes = super.Name & " "  & supername
	getSupertypes = super.Name
end function


function attrType(namespace, umltype)
	attrType = ""
	if umltype = "CharacterString" then
		attrType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Boolean" then
		attrType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Date" then
		attrType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "DateTime" then
		attrType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Integer" then
		attrType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Real" then
		attrType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Punkt" then
		attrType = "http://skjema.geonorge.no/SOSI/geometri/" & umltype 
	end if
	if umltype = "Kurve" then
		attrType = "http://skjema.geonorge.no/SOSI/geometri/" & umltype 
	end if
	if umltype = "Flate" then
		attrType = "http://skjema.geonorge.no/SOSI/geometri/" & umltype 
	end if
	if umltype = "Sverm" then
		attrType = "http://skjema.geonorge.no/SOSI/geometri/" & umltype 
	end if
	if umltype = "GM_Point" then
		attrType = "http://skjema.geonorge.no/SOSI/geometri/" & umltype 
	end if
	if umltype = "GM_Curve" then
		attrType = "http://skjema.geonorge.no/SOSI/geometri/" & umltype 
	end if
	if umltype = "GM_Surface" then
		attrType = "http://skjema.geonorge.no/SOSI/geometri/" & umltype 
	end if
	if umltype = "GM_Solid" then
		attrType = "http://skjema.geonorge.no/SOSI/geometri/" & umltype 
	end if

	if attrType = "" and umltype <> "" then
		attrType = namespace & "/" & umltype 
	end if
end function

function listBaseType(ftname,umlname, umltype)
	listBaseType = "*"
	if umltype = "CharacterString" then
		if umlname = "navnerom" or umlname = "namespace" then
			listBaseType = "http://data.geonorge.no/SOSI/" & Kortnavn 
		else
			if umlname = "lokalId" or umlname = "localId" then
				listBaseType = ftname & ".1"
			else
				listBaseType = "Some text"
			end if
		end if
	end if
	if umltype = "Boolean" then
		listBaseType = "true"
	end if
	if umltype = "Date" then
		listBaseType = "2019-05-04"
	end if
	if umltype = "DateTime" then
		listBaseType = "2019-05-04T21:08:00Z"
	end if
	if umltype = "Integer" then
		listBaseType = "42"
	end if
	if umltype = "Real" then
		listBaseType = "92.92"
	end if
end function

function mapBaseType(namespace, umltype)
	mapBaseType = ""
	if umltype = "CharacterString" then
		mapBaseType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Boolean" then
		mapBaseType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Date" then
		mapBaseType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "DateTime" then
		mapBaseType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Integer" then
		mapBaseType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if umltype = "Real" then
		mapBaseType = "http://skjema.geonorge.no/SOSI/basistype/" & umltype 
	end if
	if mapBaseType = "" and umltype <> "" then
		mapBaseType = namespace & "/" & umltype 
	end if
end function


function listCodeType(element)
	listCodeType = "*"
	dim attr as EA.Attribute
	for each attr in element.Attributes
		listCodeType = attr.Name
		if attr.Default <> "" then listCodeType = attr.Default
		exit for
	next
end function

sub listGeometryType(elementName, geomtype, indent)

		if geomtype = "Punkt" or geomtype = "GM_Point" then
				pnteller = pnteller + 1
				SessionOutput(indent & "<gml:Point gml:id=""" & elementName & ".pn." & pnteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
				SessionOutput(indent & "  <gml:pos>60.02 10.1</gml:pos>")
				SessionOutput(indent & "</gml:Point>")
		end if
		if geomtype = "Sverm" or geomtype = "GM_MultiPoint" then
			'getSosiGeometritype = "SVERM"
		end if
		if geomtype = "Kurve" or geomtype = "GM_Curve" or geomtype = "GM_CompositeCurve" then
				cuteller = cuteller + 1
'				SessionOutput(indent & "<gml:Curve gml:id = """ & elementName & ".cu." & cuteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
'				SessionOutput(indent & "  <gml:segments>
'				SessionOutput(indent & "    <gml:LineStringSegment>
'				SessionOutput(indent & "      <gml:posList>60.02 10.1 60.02 10.3 60.03 10.2</gml:posList>")
'				SessionOutput(indent & "    </gml:LineStringSegment>
'				SessionOutput(indent & "  </gml:segments>
'				SessionOutput(indent & "</gml:Curve>


				SessionOutput(indent & "<gml:LineString gml:id=""" & elementName & ".cu." & cuteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
				SessionOutput(indent & "  <gml:posList>60.02 10.1 60.02 10.3 60.03 10.2</gml:posList>")
				SessionOutput(indent & "</gml:LineString>")
		end if
		if geomtype = "Flate" or geomtype = "GM_Surface" or geomtype = "GM_CompositeSurface" then
'				SessionOutput(indent & "<gml:Surface gml:id = """ & elementName & ".su.1"" srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
				suteller = suteller + 1
				SessionOutput(indent & "<gml:Polygon gml:id=""" & elementName & ".su." & suteller & """ srsName=""http://www.opengis.net/def/crs/epsg/0/4258"">")
				SessionOutput(indent & "  <gml:exterior>")
				SessionOutput(indent & "    <gml:LinearRing>")
				SessionOutput(indent & "      <gml:posList>60.02 10.1 60.02 10.3 60.03 10.2 60.02 10.1</gml:posList>")
				SessionOutput(indent & "    </gml:LinearRing>")
				SessionOutput(indent & "  </gml:exterior>")
				SessionOutput(indent & "</gml:Polygon>")
'				SessionOutput(indent & "</gml:Surface>")
		end if
		if geomtype = "GM_Solid" or geomtype = "GM_CompositeSolid" then
			'getSosiGeometritype = "NO GO"
			dim height
			height = 6.0
			call generateSolidExample(elementName, indent, height)
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

function trimDefinitionText(tekst)
	'removes all formatting in notes fields
    Dim txt, res, tegn, i, u
    u=0
	trimDefinitionText = ""
		txt = Trim(tekst)
		res = ""
		' loop gjennom alle tegn
		For i = 1 To Len(txt)
		  tegn = Mid(txt,i,1)
		  If tegn = "<" Then
				u = 1
			   'res = res + " "
		  Else 
			If tegn = ">" Then
				u = 0
			   'res = res + " "
				'If tegn = """" Then
				'  res = res + "'"
			Else
				  If tegn < " " Then
					res = res + " "
				  Else
					if u = 0 then
						res = res + Mid(txt,i,1)
					end if
				  End If
				'End If
			End If
		  End If
		  
		Next
		
	trimDefinitionText = res

end function

function trim0DefinitionText(tekst)

    Dim txt, res, tegn, i, u
    u=0
	trim0DefinitionText = ""
		txt = Trim(tekst)
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
	trim0DefinitionText = res
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

sub generateSolidExample(elementName, indent, height)

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
	dim s1(7,2), c1(2), z1(2), h1, posNum, i
	h1 = height
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

sub SessionOutput(text)
	Session.Output(text)
end sub

listRDFXMLExample
