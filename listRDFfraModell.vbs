option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		listRDFfraModell
' purpose:		Generere filer på standard owl, html og skos-format for alle elementer i en UML-modell
' version:		2020-04-23, 04-26, 04-27, 04-29
' version:		2020-11-24/25 lager ei owl/rdf/turtle-fil
' version:		2021-06-12 tilpasset ModellDCAT versjon 1.1 
' author: 		Kent Jonsrud Kartverket
'
' ModellDCAT versjon 1.1
' TODO:			alle underpakker ut som Module
' TODO:			Stereotype som Module, tagged values med eller som egne Modules?
' TODO:			Union -> :Choice
' TODO:		v	supertyper blir FT_S_supertype
' TODO:			URI til eksterne kodelister - tas fra tV CodeList
' TODO:			ny license
' TODO:			namedIndividual tas ut
' TODO:		v	SimpleType skal ha Title - Kun en gang ut ()
' TODO:			aggregering - samling
' TODO:			begrepskobling: tV begrep = uri ? -> dct:Subject
' TODO:			Constraints -> Begrensningsregel
' TODO:			Note -> :Note (koblet til pakker og klasser, evt. også assosiasjoner)
' TODO:			Presentasjonsnavn fra toLabel() - ikke ekspandere kjente forkortelser (SOSI, NVDB, ++), andre? (-_?)
' TODO:			Begrepsrefeanse? (tV begrepsreferanse)
' TODO:			Hente Produsent fra en eierpakke? (tV Organisasjon)
' TODO:			Legge inn stien til endepunktet som fila er høstet fra?
' TODO:			kall datatime fra sub
'
' TODO:			fikse utskriftene til system output så de gir forståelig info om framdrift (nå kommer det en gml-fil ut der)
'
' 				Genererer ei fil på standard SKOS-format med alle begreper listet opp
' 				Generere ei SKOS-fil og ei html-fil pr. begrep, for direkte oppslag på hver http-URI.
' TBD:			/ i rolle og egenskapsnavn, feiler unødvendig på nøsting av første konkrete subtype
' TBD:			alle tagger, stereotyper og multiplisiteter o.l vises i html, med tagzzz = abc, stereotype = CodeList, multiplisitet = 1..2 etc.
' TBD:			rydde i koden, http://skjema.geonorge.no/basistype/Punkt + Flate + Kurve ++ selvom de er linket opp
' TBD:			liste ut navn på hvilke underpakke som elementet er definet i (referanse til Module, eller til IM)
' TBD:			hente fra tagged value definition og designation på alle mdoellelementer
' TBD			erstatte blanke i pakkenavn med - ?, nå fjernes blanke
' TBD			småfeil: doble , - app->pp - trailing space (?) - hasValueFrom på kodelisteegenskaper

	DIM kortnavnFSO
	DIM pkgFSO
	DIM skosFSO
	DIM htmlFSO
	DIM ttlFSO
	DIM skos2FSO
	DIM html2FSO
	DIM skosFile
	DIM htmlFile
	DIM ttlFile
	DIM skos2File
	DIM html2File
	DIM skosFileName
	DIM htmlFileName
	DIM ttlFileName
	DIM skos2FileName
	DIM html2FileName

	DIM debug, namespace, kortnavn, pnteller, cuteller, suteller, soteller, obteller, label, pkgname, pkgNCname, pkgNCFolder, nsprefix, supertype
	debug = false

sub listRDFfraModell()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	Dim tittel, klasseliste
	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		if Repository.GetTreeSelectedItemType() = otPackage then
			if UCASE(theElement.Element.Stereotype) = "APPLICATIONSCHEMA" then
				'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
						dim message, indent
						message = "Lister ut alle modellelementer til en OWL/RDF/Turtle-fil i henhold til den nasjonale spesifikasjonen ModellDCAT - versjon 1.1." & vbCrLf
						message = message & "Kjøres fra en UML-pakke der alt innhold er i henhold til SOSI-standarden Regler for UML-modellering - versjon 5.1." & vbCrLf
						message = message & "Lister også ut en filstruktur med html- og SKOS-filer for hvert modellelement. "
						message = message & "Denne filstrukturen kan lastes opp på en server og dette vil gi direkte tilgang til modellelementene via http-URI-ene." & vbCrLf
				dim box
				box = Msgbox ("Script listRDFfraModell" & vbCrLf & vbCrLf & "Scriptversion 2021-06-12" & vbCrLf & vbCrLf & message & vbCrLf & "  Package : [" & theElement.Name & "].",1)
				select case box
				case vbOK
					dim xsdfile
					'tømmer System Output for lettere å fange opp hele gml-fila
					Repository.ClearOutput "Script"
					Repository.CreateOutputTab "Error"
					Repository.ClearOutput "Error"
					kortnavn = getPackageTaggedValue(theElement,"SOSI_kortnavn")
					if kortnavn = "" then
						kortnavn = theElement.Name
					'	Repository.WriteOutput "Script", "Pakken mangler tagged value SOSI_kortnavn! Kjører midlertidig videre med pakkenavnet som forslag til kortnavn: " & vbCrLf & kortnavn, 0
					end if

					namespace = getPackageTaggedValue(theElement,"targetNamespace")
					if namespace = "" then
						namespace = "http://jonsrud.org/SOSI/ontologi/" & kortnavn
					end if
					
					xsdfile = getPackageTaggedValue(theElement,"xsdDocument")
					if xsdfile = "" then
						xsdfile = kortnavn & ".xsd"
					end if

					nsprefix = getPackageTaggedValue(theElement,"xmlns")
					if nsprefix = "" then
						nsprefix = "app"
					end if

					SessionOutput("<?xml version=""1.0"" encoding=""utf-8""?>")
					SessionOutput("<wfs:FeatureCollection")
					SessionOutput("  xmlns=""" & utf8(namespace) & """")
					SessionOutput("  xmlns:wfs=""http://www.opengis.net/wfs/2.0""")
					SessionOutput("  xmlns:gml=""http://www.opengis.net/gml/3.2""")
					SessionOutput("  xmlns:xlink=""http://www.w3.org/1999/xlink""")
					SessionOutput("  xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""")
					SessionOutput("  xsi:schemaLocation=""" & utf8(namespace))
					'SessionOutput("                     """ & namespace & "." & kortnavn & ".xsd""")
					SessionOutput("                     " & utf8(namespace) & "/" & utf8(xsdfile))
					SessionOutput("                     http://www.opengis.net/wfs/2.0")
					SessionOutput("                     http://schemas.opengis.net/wfs/2.0/wfs.xsd""")
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
					SessionOutput("  timeStamp=""" & Year(Date) & "-" & tm & "-" & td & "T" & tt & ":" & tmin & ":" & tsek & "Z""")
					SessionOutput("  numberMatched=""unknown""")
					SessionOutput("  numberReturned=""0"">")
					pnteller=0
					cuteller=0
					suteller=0
					soteller=0
					obteller=0
					
					'katalog for kortnavnet
			'		Set kortnavnFSO=CreateObject("Scripting.FileSystemObject")
			'		kortnavn = getNCNameX(kortnavn)
			'		if not kortnavnFSO.FolderExists(kortnavn) then
			'			kortnavnFSO.CreateFolder kortnavn
			'		end if

					'katalog for pakken
					pkgname = kortnavn & "/" & getNCNameX(theElement.Name)
					pkgNCname = getNCNameX(theElement.Name)
					Set pkgFSO=CreateObject("Scripting.FileSystemObject")
					SessionOutput(" pkgNCname: " & pkgNCname  & "...")
					SessionOutput(" pkgFSO.GetAbsolutePathName: " & pkgFSO.GetAbsolutePathName("./")  & "...")
					SessionOutput(" Repository.ConnectionString: " & Repository.ConnectionString() & "...")
					SessionOutput(" pkgFSO.GetBaseName: " & pkgFSO.GetBaseName(Repository.ConnectionString())  & "...")
					SessionOutput(" pkgFSO.GetParentFolderName: " & pkgFSO.GetParentFolderName(Repository.ConnectionString())  & "...")
					pkgNCFolder = pkgFSO.GetParentFolderName(Repository.ConnectionString())  & "\" & pkgNCname
					if not pkgFSO.FolderExists(pkgNCFolder) then
						pkgFSO.CreateFolder pkgNCFolder
					end if

					label = getPackageTaggedValue(theElement,"SOSI_presentasjonsnavn")
					if label = "" then label = theElement.Name
					' filer for pakken
					htmlFileName = pkgNCFolder & "/index.html"
					SessionOutput(" htmlFileName: " & htmlFileName )
					Set htmlFSO = CreateObject("Scripting.FileSystemObject")
					Set htmlFile = htmlFSO.CreateTextFile(htmlFileName,True,False)
					htmlFile.Write"<!DOCTYPE html>" & vbCrLf
					htmlFile.Write"<html lang=""no"">" & vbCrLf
					htmlFile.Write"	<head>" & vbCrLf
					htmlFile.Write"	  <meta charset=""utf-8""/>" & vbCrLf
					htmlFile.Write"	  <title>" & utf8(kortnavn) & "</title>" & vbCrLf
					htmlFile.Write"	</head>" & vbCrLf
					htmlFile.Write"	<body>" & vbCrLf
					'htmlFile.Write"    <p>xml:base=" & utf8(namespace) & "</p>" & vbCrLf
					htmlFile.Write"	   <h1>" & utf8(kortnavn) & "</h1>" & vbCrLf
					if getPackageTaggedValue(theElement,"SOSI_presentasjonsnavn") <> "" then
						htmlFile.Write"    <p>presentasjonsnavn = " & utf8(getPackageTaggedValue(theElement,"SOSI_presentasjonsnavn")) & "</p>" & vbCrLf
					end if
					htmlFile.Write"    <p>applikasjonsskjemaets definisjon = " & utf8(getCleanDefinitionText(theElement.Notes)) & "</p>" & vbCrLf
					htmlFile.Write"    <p>publiseringstidspunkt = " & Year(Date) & "-" & tm & "-" & td & "T" & tt & ":" & tmin & ":" & tsek & "Z</p>" & vbCrLf
					htmlFile.Write"    <p>http-URI = " & utf8(namespace) & "/" & utf8(kortnavn) & "</p>" & vbCrLf

					'htmlFile.Write"   <table border=""1"">" & vbCrLf
					htmlFile.Write"   <table>" & vbCrLf
					htmlFile.Write"   <tbody align=""left"">" & vbCrLf
					'htmlFile.Write"   <tbody>" & vbCrLf
					htmlFile.Write"   <tr>" & vbCrLf
					htmlFile.Write"   <th>_______Modellbegrep___________</th>	<th>Definisjon___________</th></tr><tr>" & vbCrLf
					
					skosFileName = pkgNCFolder & ".rdf"
					SessionOutput(" skosFileName: " & skosFileName )
					Set skosFSO = CreateObject("Scripting.FileSystemObject")
					Set skosFile = skosFSO.CreateTextFile(skosFileName,True,False)
					skosFile.Write"<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
					skosFile.Write"<rdf:RDF" & vbCrLf
					skosFile.Write"  xmlns:skos=""http://www.w3.org/2004/02/skos/core#""" & vbCrLf
					skosFile.Write"  xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""" & vbCrLf
					skosFile.Write"  xml:base=""" & utf8(namespace) & "/"">" & vbCrLf
					skosFile.Write"  <skos:Concept rdf:about=""" & utf8(namespace) & "/" & utf8(pkgNCname) & """>" & vbCrLf
					skosFile.Write"    <skos:inScheme rdf:resource=""" & utf8(namespace) & """/>" & vbCrLf
					skosFile.Write"    <skos:prefLabel xml:lang=""no"">" & utf8(label) & "</skos:prefLabel>" & vbCrLf
					skosFile.Write"    <skos:definition xml:lang=""no"">" & utf8(getCleanDefinitionText(theElement.Notes)) & "</skos:definition>" & vbCrLf
					'skosFile.Write"    <skos:broader rdf:resource="http://skjema.geonorge.no/SOSI/kodeliste/AdmEnheter/2020/Fylkesnummer/01"/>" & vbCrLf
					skosFile.Write"  </skos:Concept>" & vbCrLf
					skosFile.Write"  <skos:Collection rdf:about=""" & utf8(namespace) & "/" & utf8(pkgNCname) & "Collection"">" & vbCrLf

					ttlFileName = pkgNCFolder & ".ttl"
					SessionOutput(" ttlFileName: " & ttlFileName )
					Set ttlFSO = CreateObject("Scripting.FileSystemObject")
					Set ttlFile = skosFSO.CreateTextFile(ttlFileName,True,False)
					ttlFile.Write"@prefix " & nsprefix & ": <" & utf8(namespace) & "/" & utf8(pkgNCname) & "/> ." & vbCrLf
					ttlFile.Write"@prefix sosi:  <http://skjema.geonorge.no/SOSI/basistype/> ." & vbCrLf
					ttlFile.Write"@prefix adms:  <http://www.w3.org/ns/adms#> ." & vbCrLf
					ttlFile.Write"@prefix dct:   <http://purl.org/dc/terms/> ." & vbCrLf
					ttlFile.Write"@prefix owl:   <http://www.w3.org/2002/07/owl#> ." & vbCrLf
					ttlFile.Write"@prefix xsd:   <http://www.w3.org/2001/XMLSchema#> ." & vbCrLf
					ttlFile.Write"@prefix dcat:  <http://www.w3.org/ns/dcat#> ." & vbCrLf
					ttlFile.Write"@prefix foaf:  <http://xmlns.com/foaf/0.1/> ." & vbCrLf
					ttlFile.Write"@prefix vcard: <http://www.w3.org/2006/vcard/ns#> ." & vbCrLf
					ttlFile.Write"@prefix locn: <http://www.w3.org/ns/locn#> ." & vbCrLf
					ttlFile.Write"@prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#> ." & vbCrLf
					ttlFile.Write"@prefix skos: <http://www.w3.org/2004/02/skos/core#> ." & vbCrLf
					ttlFile.Write"@prefix xkos: <https://rdf-vocabulary.ddialliance.org/xkos/> ." & vbCrLf
					ttlFile.Write"@prefix xsd: <http://www.w3.org/2001/XMLSchema#> ." & vbCrLf
					ttlFile.Write"@prefix modelldcatno: <https://data.norge.no/vocabulary/modelldcatno#> ." & vbCrLf & vbCrLf

					ttlFile.Write nsprefix & ":" & utf8(pkgNCname) & " 		a           modelldcatno:InformationModel , owl:NamedIndividual ;" & vbCrLf
					ttlFile.Write"modelldcatno:containsModelElement" & vbCrLf
					klasseliste = getClassList(nsprefix, theElement)
					ttlFile.Write"        " & utf8(Mid(klasseliste,3,Len(klasseliste))) &  " ;" & vbCrLf
				'	ttlFile.Write"        app:Sted , app:Stedsnavn , app:SkrivemÃ¥te ;" & vbCrLf
					ttlFile.Write"        dct:identifier          """  & utf8(namespace) & "/" & utf8(pkgNCname) & """^^xsd:string ;" & vbCrLf
					ttlFile.Write"        dct:publisher          <https://organization-catalogue.fellesdatakatalog.digdir.no/organizations/971040238> ;" & vbCrLf
					ttlFile.Write"        dct:description          """ & utf8(getCleanDefinitionText(theElement.Notes)) & """@nb ;" & vbCrLf
					tittel = getPackageTaggedValue(theElement,"SOSI_langnavn")
					if tittel = "" then
						tittel = theElement.Name
					end if
					ttlFile.Write"        dct:title          		""" & utf8(tittel)  & """@nb ;" & vbCrLf
					ttlFile.Write"        dct:issued         	 	""" & Year(Date) & "-" & tm & "-" & td & "T" & tt & ":" & tmin & ":" & tsek & "+01:00""^^xsd:dateTime ;" & vbCrLf
					ttlFile.Write"        dct:language          <http://publications.europa.eu/resource/authority/language/NOB> ;" & vbCrLf
					ttlFile.Write"        owl:versionInfo          """ &  getPackageTaggedValue(theElement,"version")  & """^^xsd:string ;" & vbCrLf
					ttlFile.Write"        adms:status             <http://purl.org/adms/status/Completed> ;" & vbCrLf       ' =SOSI_modellstatus?
					ttlFile.Write"        dct:license             <http://creativecommons.org/licenses/by/4.0/deed.no> ;" & vbCrLf	   'TBD?
					ttlFile.Write"        dcat:contactPoint       " & nsprefix & ":standardiseringssekretariatet ;" & vbCrLf
					ttlFile.Write"        dcat:keyword            """ & utf8(label) & """@nb ;" & vbCrLf
					ttlFile.Write"        foaf:homepage           <http://sosi.geonorge.no> ;" & vbCrLf
					ttlFile.Write"        dct:spatial             <http://publications.europa.eu/resource/authority/country/NOR> ;" & vbCrLf
					ttlFile.Write"        dcat:theme              <https://psi.norge.no/los/tema/eiendom> ." & vbCrLf  & vbCrLf        'ikke eiendom!-???

					ttlFile.Write nsprefix & ":Katalog  		a        owl:NamedIndividual , dcat:Catalog ;" & vbCrLf
					ttlFile.Write"        dct:description         ""SOSI-modellregister med geografiske modeller""@nb ;" & vbCrLf
					ttlFile.Write"        dct:identifier          ""https://sosi.geonorge.no/svn""^^xsd:string ;" & vbCrLf
					ttlFile.Write"        dct:publisher          <https://organization-catalogue.fellesdatakatalog.digdir.no/organizations/971040238> ;" & vbCrLf
					ttlFile.Write"        dct:title          		""SOSI-modellregister""@nb ;" & vbCrLf
					ttlFile.Write"        dct:license             <http://creativecommons.org/licenses/by/4.0/deed.no> ;" & vbCrLf    'TBD?
					ttlFile.Write"        modelldcatno:model      " & nsprefix & ":" & utf8(pkgNCname) & " ." & vbCrLf & vbCrLf

					ttlFile.Write nsprefix & ":standardiseringssekretariatet 		 a          vcard:Kind , owl:NamedIndividual ;" & vbCrLf
					ttlFile.Write"        vcard:hasOrganizationName            ""Norwegian Mapping Authority""@en , ""Kartverket""@nn , ""Kartverket""@nb ;" & vbCrLf
					ttlFile.Write"        vcard:hasEmail			<mailto:standardiseringssekretariatet@kartverket.no> ;" & vbCrLf
					ttlFile.Write"        vcard:hasTelephone		<tel:04732118000> ." & vbCrLf & vbCrLf
					

' ----------------------
					call listFeatureTypes(theElement)
					call listSimpleTypes()
' ----------------------

					SessionOutput("</wfs:FeatureCollection>")

					skosFile.Write"  </skos:Collection>" & vbCrLf
					skosFile.Write"</rdf:RDF>"
					skosFile.Close
					Set skosFSO= Nothing
					
					htmlFile.Write"  </tr>" & vbCrLf
					htmlFile.Write"  </tbody>" & vbCrLf
					htmlFile.Write"  </table>" & vbCrLf
					htmlFile.Write"  </body>" & vbCrLf
					htmlFile.Write"</html>" & vbCrLf
					htmlFile.Close
					Set htmlFSO= Nothing
					
					ttlFile.Close
					Set ttlFSO= Nothing

				case VBcancel

				end select
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
					
' ----------------------
					call listDatatypes(theElement.Name,theElement,indent)
' ----------------------
					
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
	dim indent, ftname, tittel, propertylist
	if debug then Repository.WriteOutput "Script", "Debug: pkg.Name [" & pkg.Name & "].",0
	for i = 0 to elements.Count - 1 
		dim currentElement as EA.Element 
		set currentElement = elements.GetAt( i ) 
				
		if debug then Repository.WriteOutput "Script", "Debug: currentElement.Name [«" & currentElement.Stereotype & "» " & currentElement.Name & "] currentElement.Type [" & currentElement.Type & "] currentElement.Abstract [" & currentElement.Abstract & "].",0
		if currentElement.Type = "Class" and LCase(currentElement.Stereotype) = "featuretype" and currentElement.Abstract = 0 then
			
			SessionOutput("  <wfs:member>")
			SessionOutput("    <" & utf8(currentElement.Name) & " gml:id="""& utf8(currentElement.Name) & ".1"">")
			
			ftname = currentElement.Name
			superlist = ""
			indent = "      "

			call listDatatypes(ftname,currentElement,indent)
			
			SessionOutput("    </" & utf8(currentElement.Name) & ">")
			SessionOutput("  </wfs:member>")

		end if
	
		if currentElement.Type = "Class" or currentElement.Type = "Enumeration" then
		
			Call writeSkosElement(utf8(namespace),pkgNCFolder,currentElement.Name,utf8(getTaggedValue(currentElement,"SOSI_presentasjonsnavn")),utf8(getCleanDefinitionText(currentElement.Notes)))
			Call writeHtmlElement(utf8(namespace),pkgNCFolder & "/" & currentElement.Name,utf8(currentElement.Name),utf8(getTaggedValue(currentElement,"SOSI_presentasjonsnavn")),utf8(getCleanDefinitionText(currentElement.Notes)))
		
		
			if getTaggedValue(currentElement,"SOSI_presentasjonsnavn") <>  "" then
				tittel = """" & utf8(getTaggedValue(currentElement,"SOSI_presentasjonsnavn")) & """@nb"
			else
				tittel = """" & utf8(currentElement.Name) & """@nb"
			end if

			supertype = ""
			for each super in currentElement.BaseClasses
				supertype = super.name
			next
		'	SessionOutput("  -------------------------------supertype : " & supertype)
			if supertype <> "" then
				propertylist = nsprefix & ":" & supertype & "_s_supertype , " & getPropertyList(nsprefix,currentElement)

			else
				propertylist = getPropertyList(nsprefix,currentElement)
			end if

			'skosFile.Write"  <skos:Collection rdf:about=""" & utf8(namespace) & "/" & utf8(kortnavn) & "/Collection"">" & vbCrLf
			skosFile.Write"   <skos:member rdf:resource=""" & utf8(namespace) & "/" & utf8(pkgNCname) & "/" & utf8(currentElement.Name) & """/>" & vbCrLf

			'idxFile.Write"    <td>kode <a href=" & utf8(namespace) & "/" & utf8(codelist) & "/" & utf8(uricode) & ">	" & utf8(presentasjonsnavn) & "</a></td><td>" & utf8(getCleanDefinitionText(attr)) & "</td></tr><tr>" & vbCrLf
			htmlFile.Write"  <td><a href=" & utf8(currentElement.Name) & ">	" & utf8(toLabel(currentElement.Name)) & "</a></td><td>" & utf8(getCleanDefinitionText(currentElement.Notes)) & "</td></tr><tr>" & vbCrLf
			'htmlFile.Write"  <td><a href=" & utf8(pkgNCname) & "/" & utf8(currentElement.Name) & ">	" & utf8(toLabel(currentElement.Name)) & "</a></td><td>" & utf8(getCleanDefinitionText(currentElement.Notes)) & "</td></tr><tr>" & vbCrLf
			'htmlFile.Write"  <td><a href=" & utf8(pkgNCname) & "/" & utf8(currentElement.Name) & ">	" & utf8(toLabel(currentElement.Name)) & "</a></td><td>" & utf8(getCleanDefinitionText(currentElement.Notes)) & "</td></tr><tr>" & vbCrLf
			'htmlFile.Write"  <td>kode <a href=" & utf8(pkgname) & "/" & utf8(currentElement.Name) & ">	" & utf8(currentElement.Name) & "</a></td><td>" & utf8(currentElement.Notes) & "</td></tr><tr>" & vbCrLf

			if LCase(currentElement.Stereotype) = "featuretype" then
				ttlFile.Write nsprefix & ":" & utf8(currentElement.Name) & " 		a           owl:NamedIndividual , modelldcatno:ObjectType ;" & vbCrLf
'				dct:subject <https://data.skatteetaten.no/begreper/20b52aba-9fe1-11e5-a9f8-e4115b280940> ;
				ttlFile.Write"		modelldcatno:hasProperty" 
				ttlFile.Write" 		" & utf8(propertylist)  & " ;" & vbCrLf
			'	ttlFile.Write" 		app:stedsnavn , app:posisjon ;" & vbCrLf
			end if
			if LCase(currentElement.Stereotype) = "datatype" then
				ttlFile.Write nsprefix & ":" & utf8(currentElement.Name) & " 		a           owl:NamedIndividual , modelldcatno:DataType ;" & vbCrLf
				ttlFile.Write"		modelldcatno:hasProperty" 
				ttlFile.Write" 		" & utf8(propertylist)  & " ;" & vbCrLf
			'	ttlFile.Write" 		app:skrivemåte , app:kasuser ;" & vbCrLf 
			end if
			if LCase(currentElement.Stereotype) = "union" then
				ttlFile.Write nsprefix & ":" & utf8(currentElement.Name) & " 		a           owl:NamedIndividual , modelldcatno:Choice ;" & vbCrLf
				ttlFile.Write"		modelldcatno:hasProperty" 
				ttlFile.Write" 		" & utf8(propertylist)  & " ;" & vbCrLf
			'	ttlFile.Write" 		app:alternativ1 , app:alternativ2 ;" & vbCrLf
			end if
			if LCase(currentElement.Stereotype) = "codelist" or LCase(currentElement.Stereotype) = "enumeration" or currentElement.Type = "Enumeration" then
				ttlFile.Write nsprefix & ":" & utf8(currentElement.Name) & " 		a           owl:NamedIndividual , modelldcatno:CodeList ;" & vbCrLf
			end if
			ttlFile.Write"		dct:description """ & utf8(getCleanDefinitionText(currentElement.Notes)) & """@nb ;" & vbCrLf
			ttlFile.Write"		dct:identifier """ & utf8(namespace) & "/" & utf8(pkgNCname) & "/" & utf8(currentElement.Name) & """^^xsd:anyURI ;	" & vbCrLf		
			ttlFile.Write"		dct:title     " & tittel & " ;" & vbCrLf
		'	ttlFile.Write"		modelldcatno:typeDefinitionReference """ & utf8(namespace) & "/" & utf8(pkgNCname) & "/" & utf8(currentElement.Name) & """^^xsd:anyURI ;	" & vbCrLf		
			ttlFile.Write"		modelldcatno:belongsToModule   """ & utf8(pkg.Element.Name) & """@nb ." & vbCrLf & vbCrLf

			if supertype <> "" then
				ttlFile.Write nsprefix & ":" & utf8(currentElement.Name) & "_s_supertype 		a           owl:NamedIndividual , modelldcatno:Specialization ;" & vbCrLf
				ttlFile.Write"		modelldcatno:hasGeneralConcept  " & nsprefix & ":" & utf8(supertype) & " ;" & vbCrLf 
				ttlFile.Write"		dct:identifier """ & utf8(namespace) & "/" & utf8(pkgNCname) & "/" & utf8(currentElement.Name) & "_s_supertype""^^xsd:anyURI ;	" & vbCrLf		
				ttlFile.Write"		dct:title     ""Spesialisering av " & utf8(supertype) & """@nb ." & vbCrLf	& vbCrLf
			end if





' ----------------------
			call listClassProperties(namespace,pkgNCFolder,currentElement)
' ----------------------


		end if
	
	next

	dim subP as EA.Package
	for each subP in pkg.packages
	    call listFeatureTypes(subP)
	next

end sub


sub listClassProperties(ns,path,element)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim element0 as EA.Element
	dim super as EA.Element
	dim datatype as EA.Element
	dim subbtype as EA.Element
	dim conn as EA.Collection
	dim connEnd as EA.ConnectorEnd
	dim i, umlnavn, definisjon, sosinavn, sositype, sosilengde, sosimin, sosimax, sosierlik, koder, prikkniv1, roleEndElementID, sosidef, selfref, subID
	dim indent0, indent1, superlist, tittel, attnavn
	
				
	'if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then

		if debug then Repository.WriteOutput "Script", "Debug: --------listDatatypes element.Name [" & element.Name & "] element.ElementID [" & element.ElementID & "].",0

		dim attr as EA.Attribute
		for each attr in element.Attributes
			if getTaggedValue(attr,"SOSI_presentasjonsnavn") <>  "" then
				tittel = """" & utf8(getTaggedValue(attr,"SOSI_presentasjonsnavn")) & """@nb"
			else
				tittel = """" & utf8(attr.Name) & """@nb"
			end if
			if getTaggedValue(attr,"designation") <>  "" then
				tittel = tittel & " , " & utf8(getTaggedValue(attr,"designation"))
			end if
			skosFile.Write"   <skos:member rdf:resource=""" & utf8(ns) & "/" & utf8(pkgNCname) & "/" & utf8(element.Name) & "/" & utf8(attr.Name) & """/>" & vbCrLf

			'htmlFile.Write"  <td><a href=" & utf8(path) & "/" & utf8(element.Name) & "/" & utf8(attr.Name) & ">	" & utf8(toLabel(attr.Name)) & "</a></td><td>" & utf8(getCleanDefinitionText(attr.Notes)) & "</td></tr><tr>" & vbCrLf
			htmlFile.Write"  <td><a href=" & utf8(element.Name) & "/" & utf8(attr.Name) & ">	" & utf8(toLabel(attr.Name)) & "</a></td><td>" & utf8(getCleanDefinitionText(attr.Notes)) & "</td></tr><tr>" & vbCrLf

			attnavn = attr.Name
			if attr.default <> "" then
				attnavn = Trim(attr.Default)
			end if
			Call writeSkosElement(utf8(ns),path & "/" & element.Name,attnavn,utf8(getTaggedValue(attr,"SOSI_presentasjonsnavn")),utf8(getCleanDefinitionText(attr.Notes)))
			Call writeHtmlElement(utf8(ns),path & "/" & element.Name & "/" & attnavn,utf8(attr.Name),utf8(getTaggedValue(attr,"SOSI_presentasjonsnavn")),utf8(getCleanDefinitionText(attr.Notes)))

			if LCase(element.Stereotype) = "featuretype" or LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" then
				ttlFile.Write nsprefix & ":" & utf8(element.Name) & "_" & utf8(attr.Name) & " 		a           owl:NamedIndividual , modelldcatno:Attribute ;" & vbCrLf
				if attr.ClassifierID = 0 or isBasicType(attr.Type) or isGeometryType(attr.Type) then
'				if attr.ClassifierID = 0 then
'					ttlFile.Write"		modelldcatno:hasSimpleType   http://skjema.geonorge.no/basistype/" & attr.Type & " ;" & vbCrLf 
					ttlFile.Write"		modelldcatno:hasSimpleType   " & nsprefix & ":" & utf8(attr.Type) & " ;" & vbCrLf 
				else
					ttlFile.Write"		modelldcatno:hasDataType   " & nsprefix & ":" & utf8(attr.Type) & " ;" & vbCrLf 
				end if
				ttlFile.Write"		dct:description """ & utf8(getCleanDefinitionText(attr.Notes)) & """@nb ;" & vbCrLf
				ttlFile.Write"		dct:identifier """ & utf8(ns) & "/" & utf8(pkgNCname) & "/" & utf8(element.Name) & "/" & utf8(attr.Name) & """^^xsd:anyURI ;	" & vbCrLf		
				ttlFile.Write"		dct:title     " & tittel & " ;" & vbCrLf
		'		ttlFile.Write"		modelldcatno:typeDefinitionReference """ & utf8(ns) & "/" & utf8(pkgNCname) & "/" & utf8(element.Name) & "/" & utf8(attr.Name) & """^^xsd:anyURI ;	" & vbCrLf		
 				ttlFile.Write"		xsd:maxOccurs 	""" & attr.UpperBound & """^^xsd:string ;	" & vbCrLf	
				ttlFile.Write"		xsd:minOccurs 	""" & attr.LowerBound & """^^xsd:nonNegativeInteger .	" & vbCrLf & vbCrLf	
			end if		
			if LCase(element.Stereotype) = "codelist" or LCase(element.Stereotype) = "enumeration" or element.Type = "Enumeration" then
				ttlFile.Write nsprefix & ":" & utf8(element.Name) & "_" & utf8(attr.Name) & " 		a           owl:NamedIndividual , modelldcatno:CodeElement ;" & vbCrLf
				ttlFile.Write"		skos:definition """ & utf8(getCleanDefinitionText(attr.Notes)) & """@nb"
				if getTaggedValue(attr,"definition") <> "" then
					ttlFile.Write" , " & utf8(getTaggedValue(attr,"definition"))
				end if
				ttlFile.Write" ;" & vbCrLf
			'	ttlFile.Write"		dct:description """"" & utf8(getCleanDefinitionText(attr.Notes)) & """@nb ;" & vbCrLf
				ttlFile.Write"		dct:identifier """ & utf8(ns) & "/" & utf8(pkgNCname) & "/" & utf8(element.Name) & "/" & utf8(attr.Name) & """^^xsd:anyURI ;	" & vbCrLf		
			'	ttlFile.Write"		dct:title     """ & utf8(getTaggedValue(attr,"SOSI_presentasjonsnavn")) & """@nb ;" & vbCrLf
			'	ttlFile.Write"		dct:title     " & tittel & " ;" & vbCrLf
				ttlFile.Write"		skos:prefLabel     " & tittel & " ;" & vbCrLf	
				ttlFile.Write"		skos:inScheme     " & nsprefix & ":" & utf8(element.Name) & " ." & vbCrLf & vbCrLf		
			'	ttlFile.Write"		modelldcatno:typeDefinitionReference """ & utf8(ns) & "/" & utf8(pkgNCname) & "/" & utf8(element.Name) & "/" & utf8(attr.Name) & """^^xsd:anyURI .	" & vbCrLf	& vbCrLf		
			end if		
		next
		

		for each conn in element.Connectors
			if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then
			sosimin= "0"
			sosimax = "*"
			else
				umlnavn = ""
				definisjon = ""
				if conn.ClientID = element.ElementID then
					set datatype = Repository.GetElementByID(conn.SupplierID)
					umlnavn = getNCNameY(conn.SupplierEnd.Role)
					definisjon = conn.SupplierEnd.RoleNote
					presentasjonsnavn = getConnectorEndTaggedValue(conn.SupplierEnd,"SOSI_presentasjonsnavn")
					if presentasjonsnavn = "" then
						presentasjonsnavn = umlnavn
					end if
					if conn.SupplierEnd.Cardinality <> "" then
						if Mid(conn.SupplierEnd.Cardinality,1,1) <> "*" then
							sosimin = Mid(conn.SupplierEnd.Cardinality,1,1)
						end if
						if Mid(conn.SupplierEnd.Cardinality,Len(conn.SupplierEnd.Cardinality),1) <> "*" then
							sosimax = Mid(conn.SupplierEnd.Cardinality,Len(conn.SupplierEnd.Cardinality),1)
						end if
					end if
				else
					set datatype = Repository.GetElementByID(conn.ClientID)
					umlnavn = getNCNameY(conn.ClientEnd.Role)
					definisjon = conn.ClientEnd.RoleNote
					presentasjonsnavn = getConnectorEndTaggedValue(conn.ClientEnd,"SOSI_presentasjonsnavn")
					if presentasjonsnavn = "" then
						presentasjonsnavn = umlnavn
					end if
					if conn.ClientEnd.Cardinality <> "" then
						if Mid(conn.ClientEnd.Cardinality,1,1) <> "*" then
							sosimin = Mid(conn.ClientEnd.Cardinality,1,1)
						end if
						if Mid(conn.ClientEnd.Cardinality,Len(conn.ClientEnd.Cardinality),1) <> "*" then
							sosimax = Mid(conn.ClientEnd.Cardinality,Len(conn.ClientEnd.Cardinality),1)
						end if
					end if					
				end if
				if umlnavn <> "" then
					skosFile.Write"  <skos:member rdf:resource=""" & utf8(ns) & "/" & utf8(element.Name) &"/" & utf8(umlnavn) & """/>" & vbCrLf

					htmlFile.Write"  <td><a href=" & utf8(path) & "/" & utf8(element.Name) & ">	" & utf8(toLabel(umlnavn)) & "</a></td><td>" & utf8(getCleanDefinitionText(definisjon)) & "</td></tr><tr>" & vbCrLf

					Call writeSkosElement(utf8(ns),path & "/" & element.Name,utf8(umlnavn),utf8(presentasjonsnavn),utf8(getCleanDefinitionText(definisjon)))
					Call writeHtmlElement(utf8(ns),path & "/" & element.Name & "/" & umlnavn,utf8(umlnavn),utf8(presentasjonsnavn),utf8(getCleanDefinitionText(definisjon)))

'				if LCase(element.Stereotype) = "featuretype" or LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" then
			'		ttlFile.Write nsprefix & ":" & utf8(umlnavn) & " a           owl:NamedIndividual , modelldcatno:Role ;" & vbCrLf
					ttlFile.Write nsprefix & ":" & element.Name & "_" & utf8(umlnavn) & " 		a           owl:NamedIndividual , modelldcatno:Role ;" & vbCrLf
					ttlFile.Write"		modelldcatno:hasDataType   " & nsprefix & ":" & utf8(datatype.Name) & " ;" & vbCrLf 
					ttlFile.Write"		dct:description """ & utf8(getCleanDefinitionText(definisjon)) & """@nb ;" & vbCrLf
					ttlFile.Write"		dct:identifier """ & utf8(ns) & "/" & utf8(pkgNCname) & "/" & utf8(element.Name) & "/" & utf8(umlnavn) & """^^xsd:anyURI ;	" & vbCrLf		
					ttlFile.Write"		dct:title     """ & utf8(presentasjonsnavn) & """@nb ;" & vbCrLf
			'		ttlFile.Write"		modelldcatno:typeDefinitionReference " & utf8(ns) & "/" & utf8(pkgNCname) & "/" & utf8(element.Name) & "/" & utf8(umlnavn) & """^^xsd:anyURI ;	" & vbCrLf		

					ttlFile.Write"		xsd:maxOccurs 	""" & sosimax & """^^xsd:string ;	" & vbCrLf	
					ttlFile.Write"		xsd:minOccurs 	""" & sosimin & """^^xsd:nonNegativeInteger .	" & vbCrLf & vbCrLf	
'				end if		



				end if
			end if
		next

end sub

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
	dim indent0, indent1, superlist
	
				
	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then

		if debug then Repository.WriteOutput "Script", "Debug: --------listDatatypes element.Name [" & element.Name & "] element.ElementID [" & element.ElementID & "].",0

		dim attr as EA.Attribute
		for each attr in element.Attributes

			'SessionOutput(indent & "<" & attr.Name & ">")

			if getSosiGeometritype(attr) = "" then
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
			else
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
			if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then

			else
				'Repository.WriteOutput "Script", "Debug: Supplier Role.Name [" & conn.SupplierEnd.Role & "] datatypens SOSI_navn [" & getTaggedValue(Repository.GetElementByID(conn.ClientID).Name,"SOSI_navn") & "].",0
				'Repository.WriteOutput "Script", "Debug: Client Role.Name [" & conn.ClientEnd.Role & "] datatypens SOSI_navn [" & getTaggedValue(Repository.GetElementByID(conn.ClientID).Name,"SOSI_navn") & "].",0
				if debug then Repository.WriteOutput "Script", "Debug: Supplier Role.Name [" & conn.SupplierEnd.Role & "] datatypens navn [" & Repository.GetElementByID(conn.SupplierID).Name & "], conn.SupplierID [" & conn.SupplierID & "].",0
				if debug then Repository.WriteOutput "Script", "Debug: Client Role.Name [" & conn.ClientEnd.Role & "] datatypens navn [" & Repository.GetElementByID(conn.ClientID).Name & "], conn.ClientID [" & conn.ClientID & "].",0

				if conn.ClientID = element.ElementID then
					if getConnectorEndTaggedValue(conn.SupplierEnd,"xsdEncodingRule") <> "notEncoded" then
						set datatype = Repository.GetElementByID(conn.SupplierID)
						umlnavn = conn.SupplierEnd.Role
						if conn.ClientEnd.Aggregation = 2 and conn.SupplierID <> conn.ClientID then
							'composition+mandatory->nest as datatype inline?
							SessionOutput(indent & "<" & utf8(umlnavn) & ">")
							indent0 = indent & "  "
'						'	SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
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
'							call listDatatypes(ftname, datatype,indent1)
'						'	SessionOutput(indent0 & "</" & utf8(datatype.Name) & ">")
							SessionOutput(indent & "</" & utf8(umlnavn) & ">")
							if conn.SupplierEnd.Cardinality <> "0..1" and conn.SupplierEnd.Cardinality <> "1..1" and conn.SupplierEnd.Cardinality <> "1" then
								SessionOutput(indent & "<" & utf8(umlnavn) & ">")
								indent0 = indent & "  "
								SessionOutput(indent0 & "<" & utf8(datatype.Name) & ">")
								indent1 = indent0 & "  "
								if datatype.Abstract = 1 then
									'must move down to make an example of a instanciable subtype of the class pointed to TODO, NB needed on mandatory attributes!
									call getFirstConcreteSubtypeName(datatype,subID)
									set subbtype = Repository.GetElementByID(subID)
									call listDatatypes(ftname, subbtype,indent1)
								else
									call listDatatypes(ftname, datatype,indent1)
								end if
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
									if datatype.Abstract = 1 then
										'must move down to make an example of a instanciable subtype of the class pointed to TODO, NB needed on mandatory attributes!
										SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(getFirstConcreteSubtypeName(datatype,subID)) & "." & selfref & """/>")
									else
										SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(datatype.Name) & "." & selfref & """/>")
									end if
'								SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(datatype.Name) & "." & selfref & """/>")
								if debug then Repository.WriteOutput "Script", "Debug: SupplierEnd.Cardinality [" & conn.SupplierEnd.Cardinality & "].",0
								if conn.SupplierEnd.Cardinality <> "0..1" and conn.SupplierEnd.Cardinality <> "1..1" and conn.SupplierEnd.Cardinality <> "1" then
									if datatype.Abstract = 1 then
										'must move down to make an example of a instanciable subtype of the class pointed to TODO, NB needed on mandatory attributes!
										SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(getFirstConcreteSubtypeName(datatype,subID)) & "." & selfref + 1 & """/>")
									else
										SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(datatype.Name) & "." & selfref + 1 & """/>")
									end if
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
								if datatype.Abstract = 1 then
									'must move down to make an example of a instanciable subtype of the class pointed to TODO, NB needed on mandatory attributes!
									call getFirstConcreteSubtypeName(datatype,subID)
									set subbtype = Repository.GetElementByID(subID)
									call listDatatypes(ftname, subbtype,indent1)
								else
									call listDatatypes(ftname, datatype,indent1)
								end if
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
									if datatype.Abstract = 1 then
										'must move down to make an example of a instanciable subtype of the class pointed to TODO, NB needed on mandatory attributes!
										SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(getFirstConcreteSubtypeName(datatype,subID)) & "." & selfref + 1 & """/>")
									else
										SessionOutput(indent & "<" & utf8(umlnavn) & " xlink:href=""#" & utf8(datatype.Name) & "." & selfref + 1 & """/>")
									end if
							end if
							end if
						end if
					end if
				end if

			end if

		next

	end if

end sub


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

sub	listSimpleTypes()

'	må lage lokale kopier av ISO-basistypene for å peke på

	ttlFile.Write nsprefix & ":Integer 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Integer""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":Real 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Real""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":CharacterString 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/CharacterString""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":LanguageString 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/LanguageString""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":Date 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Date""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":DateTime 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/DateTime""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":Boolean 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Boolean""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":Uri 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Uri""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		


	ttlFile.Write nsprefix & ":Punkt 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Punkt""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":Sverm 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Sverm""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":Kurve 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Kurve""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":Flate 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/Flate""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":GM_Point 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/GM_Point""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":GM_Curve 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/GM_Curve""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":GM_Surface 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/GM_Surface""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":GM_Solid 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/GM_Solid""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":GM_MultiPoint 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/GM_MultiPoint""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		

	ttlFile.Write nsprefix & ":GM_Object 		a           owl:NamedIndividual , modelldcatno:SimpleType ;" & vbCrLf
	ttlFile.Write"		modelldcatno:typeDefinitionReference ""http://skjema.geonorge.no/basistype/GM_Object""^^xsd:anyURI .	" & vbCrLf	& vbCrLf		



end sub

function isBasicType(dtype)
		isBasicType = false
		if dtype = "Integer" or dtype = "Real" or dtype = "CharacterString" or dtype = "Boolean" or dtype = "Date" or dtype = "DateTime" or dtype = "Uri" or dtype = "LanguageString" then
			isBasicType = true
		end if
end function


function isGeometryType(dtype)
		isGeometryType = false
		if dtype = "Punkt" or dtype = "GM_Point" or dtype = "Sverm" or dtype = "GM_MultiPoint" or dtype = "Kurve" or dtype = "GM_Curve" or dtype = "GM_CompositeCurve" then
			isGeometryType = true
		end if
		if dtype = "Flate" or dtype = "GM_Surface" or dtype = "GM_CompositeSurface" or dtype = "GM_Solid" or dtype = "GM_CompositeSolid" then
			isGeometryType = true
		end if
		if dtype = "GM_Object" or dtype = "GM_Primitive" then
			isGeometryType = true
		end if
end function

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

function utf8X(str)
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

function getFirstConcreteSubtypeName(datatype,subID)
	dim subber as EA.Element
'	dim datatype as EA.Element
	dim conn as EA.Collection
'	dim connEnd as EA.ConnectorEnd

	subID = datatype.ElementID
	getFirstConcreteSubtypeName = "datatype.Name"
				
'	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then
	if datatype.Abstract = 1 then
		if debug then Repository.WriteOutput "Script", "Debug: --------datatype.Name [" & datatype.Name & "] datatype.ElementID [" & datatype.ElementID & "].",0
		for each conn in datatype.Connectors
			if debug then Repository.WriteOutput "Script", "Debug: conn.Type [" & conn.Type & "] conn.ClientID [" & conn.ClientID & "] conn.SupplierID [" & conn.SupplierID & "].",0
			if conn.Type = "Generalization" then
				if datatype.ElementID <> conn.ClientID then
					if debug then Repository.WriteOutput "Script", "Debug: subtype [" & Repository.GetElementByID(conn.ClientID).Name & "].",0
					set subber = Repository.GetElementByID(conn.ClientID)
					getFirstConcreteSubtypeName =  getFirstConcreteSubtypeName(subber,subID)
					exit function
				end if
			end if
		next
	end if
	
end function

function getClassList(prefix,pkg)
	dim element as EA.Element
	getClassList = ""
	for each element in pkg.Elements
		if element.Type = "Datatype" or element.Type = "Class" or element.Type = "Enumeration" then
			if getClassList <> "" then
				getClassList = getClassList & " ,"
			end if
			getClassList = getClassList & " " & prefix & ":" & element.Name
		end if
	next
	dim subP as EA.Package
	dim subelements
	for each subP in pkg.packages
	    subelements = getClassList(prefix,subP)
		if subelements <> "" then
			getClassList = getClassList & " ," & subelements
		end if
	next
end function


function getPropertyList(prefix,class1)
	dim attr as EA.Attribute
	getPropertyList = ""
	for each attr in class1.Attributes
		if getPropertyList <> "" then
			getPropertyList = getPropertyList & " ,"
		end if
'		getPropertyList = getPropertyList & " " & prefix & ":" & attr.Name
'		getPropertyList = getPropertyList & " " & prefix & ":" & class1.Name & "/" & attr.Name
		getPropertyList = getPropertyList & " " & prefix & ":" & class1.Name & "_" & attr.Name
	next
	dim conn as EA.Collection
	for each conn in class1.Connectors
		if conn.Type <> "Generalization" and conn.Type <> "Realisation" or conn.Type <> "NoteLink"  then
			if class1.ElementID = conn.ClientID then
				if conn.SupplierEnd.Role <> "" then
					getPropertyList = getPropertyList & " " & prefix & ":" & class1.Name & "_" & conn.SupplierEnd.Role
				end if
			else
				if conn.ClientEnd.Role <> "" then
					getPropertyList = getPropertyList & " " & prefix & ":" & class1.Name & "_" & conn.ClientEnd.Role
				end if
			end if
		end if
	next
end function

function getNCNameX(str)
	' make name legal NCName
	Dim txt, res, tegn, i, u
    u=0
		txt = Trim(str)
		'res = LCase( Mid(txt,1,1) )
		res = Mid(txt,1,1)
			'Repository.WriteOutput "Script", "New NCName: " & txt & " " & res,0

		' loop gjennom alle tegn
		For i = 2 To Len(txt)
		  ' blank, komma, !, ", #, $, %, &, ', (, ), *, +, /, :, ;, <, =, >, ?, @, [, \, ], ^, `, {, |, }, ~
		  ' (tatt med flere fnuttetyper, men hva med "."?) (‘'«»’)
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
		          res = res + UCase(tegn)
		          u=0
			      else
		          res = res + tegn
		        End If
		      End If
		    End If
		  End If
		Next
		' return res
		getNCNameX = res

End function

function getNCNameY(str)
	' make name legal NCName, give message if change is done
	Dim txt, res, tegn, i, u
    u=0
		txt = Trim(str)
		'res = LCase( Mid(txt,1,1) )
		res = Mid(txt,1,1)
			'Repository.WriteOutput "Script", "New NCName: " & txt & " " & res,0

		' loop gjennom alle tegn
		For i = 2 To Len(txt)
		  ' blank, komma, !, ", #, $, %, &, ', (, ), *, +, /, :, ;, <, =, >, ?, @, [, \, ], ^, `, {, |, }, ~
		  ' (tatt med flere fnuttetyper, men hva med "."?) (‘'«»’)
		  tegn = Mid(txt,i,1)
		  if tegn = " " or tegn = "," or tegn = """" or tegn = "#" or tegn = "$" or tegn = "%" or tegn = "&" or tegn = "(" or tegn = ")" or tegn = "*" Then
				SessionOutput(" tegn [" & tegn & "] er endret i streng ["  & str & "]")
			  'Repository.WriteOutput "Script", "Bad1: " & tegn,0
			  u=1
		  Else
		    if tegn = "+" or tegn = "/" or tegn = ":" or tegn = ";" or tegn = "<" or tegn = ">" or tegn = "?" or tegn = "@" or tegn = "[" or tegn = "\" Then
				SessionOutput(" tegn [" & tegn & "] er endret i streng ["  & str & "]")
			    'Repository.WriteOutput "Script", "Bad2: " & tegn,0
			    u=1
		    Else
		      If tegn = "]" or tegn = "^" or tegn = "`" or tegn = "{" or tegn = "|" or tegn = "}" or tegn = "~" or tegn = "'" or tegn = "´" or tegn = "¨" Then
					SessionOutput(" tegn [" & tegn & "] er endret i streng ["  & str & "]")
			      'Repository.WriteOutput "Script", "Bad3: " & tegn,0
			      u=1
		      else
			      'Repository.WriteOutput "Script", "Good: " & tegn,0
			      If u = 1 Then
		          res = res + UCase(tegn)
		          u=0
			      else
		          res = res + tegn
		        End If
		      End If
		    End If
		  End If
		Next
		' return res
		getNCNameY = res

End function


function getCleanDefinitionText(notes)
	'removes all formatting in notes fields, also all text after a < in the real definition text !
	' NB: men hva med &nbsp?
	' changes double quotes to singe quotes
    Dim txt, res, tegn, i, u
    u=0
	getCleanDefinitionText = ""
	txt = Trim(notes)
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
						if tegn = """" then
							res = res + "'"
						else
							res = res + Mid(txt,i,1)
						end if
					end if
				  End If
				'End If
			End If
		  End If
		  
		Next
		
	getCleanDefinitionText = res

end function

function toLabel(name)
	'expands tecnical NCNames to normal language names
    Dim txt, res, tegn, i, u
    u=0
	toLabel = ""
	txt = Trim(name)
		res = Mid(txt,1,1)
		' loop gjennom alle resterende tegn og sett inn blank og liten bokstav der det er stor bokstav
		' dersom det ikke er tre eller fire store etter hverandre - TBD
		For i = 2 To Len(txt)
			tegn = Mid(txt,i,1)
			If tegn <> LCase(tegn) Then
				res = res + " "
				res = res + LCase(tegn)
			Else 
				res = res + tegn
			End If
		Next
		
	toLabel = res

end function

Sub writeSkosElement(ns,path,name,label,description)
	Dim ftFSO
	Set ftFSO=CreateObject("Scripting.FileSystemObject")
	SessionOutput(" path: " & path )
	if not ftFSO.FolderExists(path) then
		ftFSO.CreateFolder path
	end if
	Dim ftfFSO
	Dim ftfFile
	Dim ftfFileName
	ftfFileName = path & "/" & name & ".rdf"
	SessionOutput(" SKOS ftfFileName: " & ftfFileName )
	Set ftfFSO = CreateObject("Scripting.FileSystemObject")
	Set ftfFile = ftfFSO.CreateTextFile(ftfFileName,True,False)			
	ftfFile.Write"<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
	ftfFile.Write"<rdf:RDF" & vbCrLf
	ftfFile.Write"  xmlns:skos=""http://www.w3.org/2004/02/skos/core#""" & vbCrLf
	ftfFile.Write"  xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""" & vbCrLf
	ftfFile.Write"  xml:base=""" & ns & "/" & path & "/"">" & vbCrLf
	ftfFile.Write"  <skos:Concept rdf:about=""" & ns & "/" & path & "/"  & utf8(name) & """>" & vbCrLf
	ftfFile.Write"    <skos:inScheme rdf:resource=""" & ns & "/" & path & "/" & """/>" & vbCrLf
	if label <> "" then
		ftfFile.Write"    <skos:prefLabel xml:lang=""no"">" & utf8(label) & "</skos:prefLabel>" & vbCrLf
	else
		ftfFile.Write"    <skos:prefLabel xml:lang=""no"">" & utf8(toLabel(name)) & "</skos:prefLabel>" & vbCrLf
	end if
	ftfFile.Write"    <skos:definition xml:lang=""no"">" & description & "</skos:definition>" & vbCrLf
	ftfFile.Write"  </skos:Concept>" & vbCrLf
	ftfFile.Write"</rdf:RDF>"
	ftfFile.Close
	Set ftFSO= Nothing
End Sub

Sub writeHtmlElement(ns,path,name,label,description)
	Dim ftFSO
	Set ftFSO=CreateObject("Scripting.FileSystemObject")
	SessionOutput(" path: " & path )
	if not ftFSO.FolderExists(path) then
		ftFSO.CreateFolder path
	end if
	Dim ftfFSO
	Dim ftfFile
	Dim ftfFileName
	ftfFileName = path & "/index.html"
	SessionOutput(" ftfFileName: " & ftfFileName )
	Set ftfFSO = CreateObject("Scripting.FileSystemObject")
	Set ftfFile = ftfFSO.CreateTextFile(ftfFileName,True,False)			
	ftfFile.Write"<!DOCTYPE html>" & vbCrLf
	ftfFile.Write"<html lang=""no"">" & vbCrLf
	ftfFile.Write"	<head>" & vbCrLf
	ftfFile.Write"	  <meta charset=""utf-8""/>" & vbCrLf
	ftfFile.Write"	  <title>" & name & "</title>" & vbCrLf
	ftfFile.Write"	</head>" & vbCrLf
	ftfFile.Write"	<body>" & vbCrLf
	ftfFile.Write"	  <h1>" & name & "</h1>" & vbCrLf
	'ftfFile.Write"    <p>xml:base=" & ns & "</p>" & vbCrLf
	if label <> "" then
		ftfFile.Write"    <p>presentasjonsnavn = " & label & "</p>" & vbCrLf
	end if
	ftfFile.Write"    <p>definisjon = " & description & "</p>" & vbCrLf
	'ftfFile.Write"    <p>http-URI = " & ns & "/" & path & "/" & name & "</p>" & vbCrLf
	ftfFile.Write"    <p>http-URI = " & ns & "/" & utf8(path) & "</p>" & vbCrLf
	ftfFile.Write"	</body>" & vbCrLf
	ftfFile.Write"</html>"
	ftfFile.Close
	Set ftFSO= Nothing
End Sub

listRDFfraModell
