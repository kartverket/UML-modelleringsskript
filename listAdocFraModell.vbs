Option Explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: listAdocFraModell
' Author: Tore Johnsen, Åsmund Tjora
' Purpose: Generate documentation in AsciiDoc syntax
' Original Date: 08.04.2021
'
' Version: 0.31 Date: 2022-08-05 Jostein Amlien: Ingen endring av funksjonalitet, kun refaktorering. Isolert adoc-syntaks og utskrift/lagring fra modell-logikken
' Version: 0.30 Date: 2022-07-04 Jostein Amlien: Lempa på rekkefølgekrav i hovedrutina, lagt til flere rutiner for Asciidoc-syntaks, formattert tekst og output, refaktorert Sub Relasjoner
' Version: 0.29 Date: 2022-06-17 Jostein Amlien: Definert og tatt i bruk noen enkle funksjoner for Asciidoc-syntaks
' Version: 0.28 Date: 2022-06-10 Kent Jonsrud: dersom diagrammer har beskrivelse så legges denne inn i alt=
' Version: 0.27 Date: 2022-01-17 Kent Jonsrud: endra Alt= til alt= på alternative bildetekster
' Version: 0.26 Date: 2021-12-15 Kent Jonsrud: retting av småfeil etter forrige retting
' Version: 0.25 Date: 2021-12-14 Kent Jonsrud: skille ocl fra beskrivelse med linjeskift før --, og komma fjernes kun fra bildetekst
' Version: 0.24 Date: 2021-12-09 Kent Jonsrud: AS på 2. nivå (===), FT og UP på nivåer under ned til 5. nivå (=====), tilpasset :toclevel: 4 og [discrete]
' Version: 0.23 Date: 2021-12-08 Kent Jonsrud: AS på 3. nivå, FT og UP på samme nivå under (kan justeres på linje ca. 200)
' Version: 0.22 Date: 2021-12-07 Kent Jonsrud: linjeskift i noter endres ikke lenger til blanke
' Version: 0.21 Date: 2021-11-25 Kent Jonsrud: endra nøsting til å nøste kun fire nivå ned (AS(FT og UP(FT og UP og UP::FT og (UP/)UP2::FT etc.)))
' Version: 0.20 Date: 2021-11-15 Kent Jonsrud: endra på Alt= , dobbeltparenteser, kun en lenke til ekstern kodeliste, nøsting av underpakker
' Version: 0.19 Date: 2021-10-05 Kent Jonsrud: satt inn Alt= i alle image:
' Version: 0.18 Date: 2021-09-28 Kent Jonsrud: bytta til eksplisitt skillelinje ("'''") og rydda vekk død kode
' Version: 0.17 Date: 2021-09-28 Kent Jonsrud: tatt bort eksplisitt nummerering av figurer
' Version: 0.16 Date: 2021-09-21 Kent Jonsrud: flyttet supertypen til slutt og laget hyperlinker til subtypene
' Version: 0.15 Date: 2021-09-17 Kent Jonsrud: smårettinger
' Version: 0.14 Date: 2021-09-16 Tore Johnsen/Kent Jonsrud: hyperlenker til egenskapenes typer innenfor modellen og absolutte lenker til basistyper
' Version: 0.13 Date: 2021-09-10 Kent Jonsrud: smårettinger
' Version: 0.12 Date: 2021-09-09 Kent Jonsrud: smårettinger, bedre angivelse av skille mellom klassene
' Version: 0.11 Date: 2021-09-05 Kent Jonsrud: Assosiasjonsnavn, ikke hente eksterne koder, sideskift for pakker i pdf
' Version: 0.10 Date: 2021-08-10 Kent Jonsrud: forbedra ledetekster
' Version: 0.9 Date: 2021-08-08 Kent Jonsrud: skriver ut navn og beskrivelse på alle operasjoner på objekttyper og datatyper
' Version: 0.8 Date: 2021-08-06 Kent Jonsrud: skriver ut alle restriksjoner på objekttyper og datatyper
' Version: 0.7 Date: 2021-07-08 Kent Jonsrud: retta en feil ved utskrift av roller
' Version: 0.6 Date: 2021-06-30 Kent Jonsrud: leser kodelister fra levende register
' Version: 0.5 Date: 2021-06-29 Kent Jonsrud: error if role list is not shown
' Date: 2021-06-24 Kent Jonsrud: endra skriptnavn fra AdocTest til listAdocFraModell
' Version: 0.4 Date: 2021-06-14 Kent Jonsrud: case-insensitiv test på navnet på tagged value SOSI_bildeAvModellelement
' Version: 0.3 Date: 2021-06-01 Kent Jonsrud: retta bildesti til app_img
' Version: 0.2 Date: 2021-04-16 Kent Jonsrud: tagged value SOSI_bildeAvModellelement på pakker og klasser: verdien vises som ekstern sti til bilde
' Date: 2021-04-15 Kent Jonsrud: diagrammer legges i underkatalog med navn enten verdien i tagged value SOSI_kortnavn eller img.
' Date: 2021-04-09/14 Kent Jonsrud:  - tagged value lists come right after definition, on packages and classes - "Spesialisering av" changed to Supertype, no list of subtypes shown
' - removed formatting in notes, except CRLF - show stereotype on attribute "Type" if present - roles shall have same simple look and structure as attributes
' - Relasjoner changed to Roller, show only ends with role names (and navigable ?) - tagged values on CodeList classes, empty tags suppressed (suppress only those from the standard profile?), heading?
' - simpler list for codelists with more than 1 code, three-column list when Defaults are used (Utvekslingsalias)
'
' TBD: tagged values on roles
' TBD: fjerne komma og sette inn - for blanke i diagrammnavn for enklere_filnavn.png ?
' TBD: show stereotype on Type DONE
' TBD: show abstract on Type 
' TBD: show navigable 
' TBD: show association type 
' TBD: output operations and constraints right after attributes and roles DONE
' TBD: codes with tagged values
' TBD: output info on associations if present
' TBD: Hva med tagged values på koder?
' TBD: if tV SOSI_bildeAvModellelement (på koder og egenskaper) -> Session.Output("image::"& tV &".png["& tV &"]")
' TBD: special handling of classes that have tV with names like FKB-A etc. and are subtypes of feature types
' TBD: write adoc and diagram files to a subfolder, ensure utf-8 in adoc (no &#229)
'		==== «dataType» Matrikkelenhetreferanse
'		Definisjon: Mulighet for &#229; koble matrikkelenhet til objekt i SSR for &#229; oppdatere bruksnavn i matrikkelen.
' TBD: opprydding !!!
'
Dim imgfolder, imgparent
Dim imgFSO
'
' Project Browser Script main function
Sub OnProjectBrowserScript()

	Dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()

	Select Case treeSelectedType

		Case otPackage
			Repository.EnsureOutputVisible "Script"
			Repository.ClearOutput "Script"
			
			' Code for when a package is selected
			Dim thePackage As EA.Package
			set thePackage = Repository.GetTreeSelectedObject()

			imgfolder = "diagrammer"
			Set imgFSO=CreateObject("Scripting.FileSystemObject")
			imgparent = imgFSO.GetParentFolderName(Repository.ConnectionString())  & "\" & imgfolder
			if not imgFSO.FolderExists(imgparent) then
				imgFSO.CreateFolder imgparent
			end if
			
'''			imgparent = ""   ''' brukes ved ny kjøring der alle diagremmer allerede er produseret
			
			Session.Output("// Start of UML-model")
			Dim topplevel  			
			topplevel = 2
			Call ListPakke( topplevel, thePackage)
			Session.Output("// End of UML-model")
		Case Else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK

	End Select
	Set imgFSO = Nothing
End Sub


''  Sub ListAsciiDoc(pakkelevel, thePackage)
Sub ListPakke(pakkelevel, thePackage)

'----------------- Overskrift og beskrivelse -----------------

	dim stereotype, overskrift, prefiks
	stereotype = tekstformatStereotype(thePackage.Element.Stereotype)

	if stereotype <> "" then
		prefiks = "Pakke: " & stereotype 
	else
		Call settInnSideskift                 ''' Hvorfor det ?
		call settInnSkillelinje

		if pakkelevel >= 4 then
			prefiks = "Underpakke: " 
		else
			prefiks = "Pakke: " 
		end if

	end if
	overskrift = prefiks & thePackage.Element.Name
	overskrift = adocOverskrift( pakkelevel, overskrift ) 

''	Call skrivIngress( pakkelevel, thePackage.Element, overskrift)
	call skrivTekst( overskrift)
	call skrivTekst( adocDefinisjonsAvsnitt( thePackage.Element) )
	call skrivProfilParameterTabell( pakkelevel, thePackage.Element) 


'----------------- Bilder og diagram-----------------

	call bildeAvModellelement( thePackage.Element)	


	Dim projectclass As EA.Project
	set projectclass = Repository.GetProjectInterface()
	
	dim diagramfil
	Dim diag As EA.Diagram
	dim alternativbildetekst
	
	For Each diag In thePackage.Diagrams
		if imgparent = "" then 
			'' bruk eksisterende diagrammer, ikke lag nye
		else
			Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, imgparent & "\" & diag.Name & ".png", 1)
			Repository.CloseDiagram(diag.DiagramID)
		end if
		
		if diag.Notes <> "" then
			alternativbildetekst = diag.Notes
		else
			alternativbildetekst = "Diagram med navn " & diag.Name & " som viser UML-klasser beskrevet i teksten nedenfor."
		end if
		
		diagramFil = imgfolder & "\" & diag.Name & ".png"

		call settInnBilde(diag.Name, diagramFil, alternativbildetekst)
	Next

'-----------------Elementer----------------- 

	Dim element As EA.Element 
	For each element in thePackage.Elements
		If isFeatureOrDataType(element)  Then	
			Call ObjektOgDatatyper(pakkelevel+1, element, thePackage)
			
		Elseif isCodelist(element) Then
			Call Kodelister(pakkelevel+1, element, thePackage)
			
		End if
	Next

'----------------- Underpakker ----------------- 

	dim nesteLevel

'	ALT 1 Underpakker flatt på samme nivå som Application Schema
'	nesteLevel = pakkelevel

'	ALT 2 Nøsting av pakker ned til nivå 4 under Application Schema
	nestelevel = pakkelevel + 1
	if pakkelevel = 4 then nestelevel = 4

'	ALT 3 TBD Nøsting helt ned med utskrift av Pakke::Klasse (Pakke/Pakke2::Klasse TBD)
' 	nestelevel = pakkelevel + 1

	dim pack as EA.Package
	for each pack in thePackage.Packages
		Call ListPakke(nestelevel, pack)
	next

end sub

'-----------------ObjektOgDatatyper-----------------
Sub ObjektOgDatatyper(elementLevel, element, pakke)

	dim overskrift
	overskrift = elementOverskrift(elementLevel, element, pakke) 

''	call skrivIngress( elementLevel, element, overskrift)
	call skrivTekst( overskrift)
	call skrivTekst( adocDefinisjonsAvsnitt( element) )
	call skrivProfilParameterTabell( elementLevel, element) 
	
	call bildeAvModellelement( element)

	call skrivUnderOverskrift(elementLevel, "Egenskaper")
'	if element.Attributes.Count = 0 then
'		skrivTekst("Inneholder ingen egenskaper") 
'	end if
	Dim att As EA.Attribute
	for each att in element.Attributes
		call skrivTabell( attributtbeskrivelse(att) )
	next

	call Relasjoner( element, adocUnderOverskrift(elementLevel, "Roller") )
	call Operasjoner( element, adocUnderOverskrift(elementLevel, "Operasjoner") )
	call Restriksjoner( element, adocUnderOverskrift(elementLevel, "Restriksjoner") )
	call ArvOgRealiseringer( element, adocUnderOverskrift( elementLevel, "Arv og realiseringer") )

end sub



sub ArvOgRealiseringer( element, tabellOverskrift)
	 
	Dim con As EA.Connector
	Dim supplier As EA.Element
	Dim client As EA.Element
	
	dim externalPackage

	dim supertyper, subtyper, realiseringer
	dim tabell
	
	tabell = adocTabellstart("20,80", tabellOverskrift)

' Supertype
	supertyper = false
	For Each con In element.Connectors
		set supplier = Repository.GetElementByID(con.SupplierID)
		If con.Type = "Generalization" And supplier.ElementID <> element.ElementID Then
			supertyper = true
			call utvidTabell( tabell, adocTabellRad("Supertype: ", targetLink(supplier)) )
		End If
	Next

' Spesialiseringer (subtyper) av klassen
	subtyper = false
	For Each con In element.Connectors
		If con.Type = "Generalization" Then
			set supplier = Repository.GetElementByID(con.SupplierID)
			set client = Repository.GetElementByID(con.ClientID)
			If supplier.ElementID = element.ElementID then 'dette er en generalisering
				if subtyper = false then
					call utvidTabell( tabell, adocTabellRad("Subtyper:", "") )
					subtyper = true
				End If
				call utvidTabell( tabell, targetLink(client) & adocLinjeskift() )
			End If
		End If
	Next

	realiseringer = false
	For Each con In element.Connectors  
'Må forbedres i framtidige versjoner dersom denne skal med 
'- full sti (opp til applicationSchema eller øverste pakke under "Model") til pakke som inneholder klassen som realiseres
		set supplier = Repository.GetElementByID(con.SupplierID)
		If con.Type = "Realisation" And supplier.ElementID <> element.ElementID Then
			set externalPackage = Repository.GetPackageByID(supplier.PackageID)
			if realiseringer = false Then
				call utvidTabell( tabell, adocTabellRad("Realisering av: ", "") )
				realiseringer = true
			end if
			call utvidTabell( tabell, getPath(externalPackage) & "::" & stereotypeNavn(supplier) & adocLinjeskift())
		end if
	next

	if supertyper or subtyper or realiseringer then
		call skrivTabell( tabell)
	end if

End sub

'-----------------ObjektOgDatatyper End-----------------


'-----------------CodeList-----------------

Sub Kodelister(elementLevel, element, pakke)

	dim overskrift
	overskrift = elementOverskrift(elementLevel, element, pakke) 
	
''	call skrivIngress( elementLevel, element, overskrift)
	call skrivTekst( overskrift)
	call skrivTekst( adocDefinisjonsAvsnitt( element) )
	call skrivProfilParameterTabell( elementLevel, element) 

	call bildeAvModellelement( element )

	if element.Attributes.Count > 0 then  
		dim tabelloverskrift
		tabelloverskrift = adocUnderOverskrift(elementLevel, "Koder i modellen")
	
		CALL skrivTabell( modellkoder(element, tabelloverskrift ) )
	else
	''' Da må kodelista være ekstern ....
	end if
	
end sub


function modellkoder(element, tabellOverskrift)

	Dim att As EA.Attribute	
	dim utvekslingsalias 
	dim tabell, tabellRad
	
	utvekslingsalias = false
	for each att in element.Attributes
		if att.Default <> "" then
			utvekslingsalias = true
			exit for
		end if
	next

	if utvekslingsalias then
		tabell = adocTabellstart("25,60,15", tabellOverskrift)
		tabellRad = adocTabellHode3( "Kodenavn:", "Definisjon:", "Utvekslingsalias:" ) 
	else
		tabell = adocTabellstart("20,80", tabellOverskrift)
		tabellRad = adocTabellHode( "Navn:", "Definisjon:" ) 		
	end if
	call utvidTabell( tabell, tabellRad )		

	for each att in element.Attributes
		tabellRad = adocTabellRad( att.Name, getCleanDefinition(att.Notes) )   
		if utvekslingsalias then call utvidTabellRad( tabellRad, adocTabellCelle(att.Default) )
		call utvidTabellRad( tabellRad, bildeAvAttributt(att, "kodelistekode")  )
		
		call utvidTabell( tabell, tabellrad )
	next
	
	modellkoder = tabell
End function

'-----------------CodeList End-----------------


function elementOverskrift(elementLevel, element, pakke)
	
	dim elementnavn, overskrift
	elementnavn = stereotypeNavn(element) 
	if isAbstract(element) then
		elementnavn = adocKursiv( elementnavn & " (abstrakt)" )    '''' NYTT: gjort abstracte klasser kursiv
	end if

	if elementLevel > 4 then   
		overskrift = adocOverskrift( 4, pakke.Name & "::" & elementnavn)
	else
		overskrift = adocOverskrift( elementLevel, elementnavn) 
	end if
	
	dim bokmerke
	bokmerke = merge( adocSkillelinje, adocBokmerke(element) )
	
	elementOverskrift = merge( bokmerke, overskrift)
end function


sub skrivIngress( elementLevel, element, overskrift)

	call skrivTekst( overskrift)
	call skrivTekst( adocDefinisjonsAvsnitt( element) )
	call skrivProfilParameterTabell( elementLevel, element) 

end sub


sub skrivProfilParameterTabell( elementLevel, element) 
	dim overskrift, tabell, listTags, tag

	overskrift = adocUnderOverskrift(elementLevel, "Profilparametre i tagged values")
	tabell = adocTabellstart("20,80", overskrift)
	listTags = false
	for each tag in element.TaggedValues
		if tag.Value = "" then	
		elseif tag.Name = "persistence" or tag.Name = "SOSI_melding" then  	''  hopp over disse
		elseif LCase(tag.Name) = "sosi_bildeavmodellelement" then 			''  tas separat
		else
			call utvidTabell( tabell, adocTabellRad( tag.Name, tag.Value) )
			listTags = true
		end if
	next

	if listTags then 
		call skrivTabell( tabell )
	end if
	
end sub



function attributtbeskrivelse( att)

	dim tabell
	tabell = adocTabellstart("20,80", "")
	
	call utvidTabell( tabell, adocTabellHode( "Navn:", att.Name ) )
	
	call utvidTabell( tabell, adocTabellRad( "Definisjon:", getCleanDefinition(att.Notes) ) )
	
	call utvidTabell( tabell, adocTabellRad( "Multiplisitet:", bounds(att) ) )
	
	if att.Default <> "" then
		call utvidTabell( tabell, adocTabellRad( "Initialverdi:", att.Default ) )
	end if

	if not att.Visibility = "Public" then
		call utvidTabell( tabell, adocTabellRad( "Visibilitet:", att.Visibility ) )
	end if
	
	dim typ, stereo
	if att.ClassifierID <> 0 then
		if isElement(att.ClassifierID) then		
			stereo = Repository.GetElementByID(att.ClassifierID).Stereotype
			typ = adocLink( att.Type, tekstformatStereotype(stereo) & att.Type )
		else
			typ = att.Type
		end if
	else
		typ = adocEksternLink( sosiBasistype( att.Type), att.Type)
	end if
	call utvidTabell( tabell, adocTabellRad( "Type:", typ) )


	dim tabellrad, listTags
	listTags = false 
	tabellrad = adocTabellRad( "Profilparametre i tagged values: ", "")
	dim tag as EA.TaggedValue
	for each tag in att.TaggedValues
''		if tag.Value = "" then												'' 	hopp over tomme tagger
''		elseif tag.Name = "persistence" or tag.Name = "SOSI_melding" then  	''  hopp over disse også
''		elseif LCase(tag.Name) = "sosi_bildeavmodellelement" then 			''  tas separat, hopp over
		call utvidTabellRad( tabellRad, tag.Name & ": " & tag.Value & adocLinjeskift() )
		listTags = true
	next
	if listTags then call utvidTabell( tabell, tabellrad)
		
	attributtbeskrivelse = tabell
end function


'-----------------Relasjoner-----------------
sub Relasjoner( element, underOverskrift)
	Dim con
	Dim supplier
	Dim client
	
'assosiasjoner
' skriv ut roller - sortert etter tagged value sequenceNumber TBD

	For Each con In element.Connectors
		If con.Type = "Association" or con.Type = "Aggregation" Then

			'' TBD:	Egenassosiasjoner.
			'	Kan løses ved å erstatte elseif med if under
			'   Eller legge det inn som et innledende særtilfelle
			'
			If element.elementID = con.SupplierID  Then 
				'dette elementet er suppliersiden - implisitt at fraklasse er denne klassen
				Call skrivRelasjon( con.ClientEnd, con.clientID, con.SupplierEnd, con, underOverskrift) 

			elseIf element.elementID = con.ClientID  Then
				'dette elementet er clientsiden, (rollen er på target)
				Call skrivRelasjon( con.SupplierEnd, con.supplierID, con.ClientEnd, con, underOverskrift) 	
				
			End If	
		End If
	Next

end sub

sub skrivRelasjon( targetEnd, targetID, currentEnd, connector, underOverskrift)
''
''	targetEnd angir Rollen: navn, def, multiplisitet og navigerbarhet
''	targetID angir klassen som Rollen peker på
''	currentEnd angir evt. aggregering, potensielt også rolle og mult. for denne enden
''	connector inneholder type og navn på selve konnektoren
''	underOverskrift skives ut bare en gang, den tømmes etter å ha blitt brukt en gang
''
	if targetEnd.Role = "" then
		exit sub
	end if

	dim tabell 
	tabell = adocTabellstart("20,80", underOverskrift)	
	
	call utvidTabell( tabell, adocTabellHode( "Rollenavn:", targetEnd.Role ) )

	If targetEnd.RoleNote <> "" Then
		call utvidTabell( tabell, adocTabellRad( "Definisjon:", getCleanDefinition(targetEnd.RoleNote) )  )
	End If
	If targetEnd.Cardinality <> "" Then
		call utvidTabell( tabell, adocTabellRad( "Multiplisitet:", "[" & targetEnd.Cardinality & "]" )   )  '''' tekstformat
	End If 
	
	DIM conType
	If currentEnd.Aggregation <> 0 Then
		if currentEnd.Aggregation = 2 then
			conType = "Komposisjon " & connector.Type
		else
			conType = "Aggregering " & connector.Type
		end if
		call utvidTabell( tabell, adocTabellRad( "Assosiasjonstype:", conType) )
	End If
'	
'	If currentEnd.Aggregation = 0 Then
'		conType = ""
'''		conType = connector.Type     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	elseif currentEnd.Aggregation = 2 then
'		conType = "Komposisjon " & connector.Type
'	else
'		conType = "Aggregering " & connector.Type
'	end if
'	if conType <> "" then call utvidTabell( tabell, adocTabellRad( "Assosiasjonstype:", conType) )

	If connector.Name <> "" Then call utvidTabell( tabell, adocTabellRad( "Assosiasjonsnavn:", connector.Name ) )

	dim textVar, targetReferanse
	targetReferanse = targetLink( Repository.GetElementByID(targetID) )
	textVar = "Til klasse"
	If targetEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
	ElseIf targetEnd.Navigable = "Non-Navigable" Then 
		textVar = textVar + adocKursiv(" (ikke navigerbar):") 
	Else 
		textVar = textVar + ":" 
	End If
	call utvidTabell( tabell, adocTabellRad( textVar, targetReferanse ) )

	if false then
		If currentEnd.Role <> "" Then
			call utvidTabell( tabell, adocTabellRad( "Fra rolle:", currentEnd.Role ) )
		End If
		If currentEnd.RoleNote <> "" Then
			call utvidTabell( tabell, adocTabellRad( "Fra rolle definisjon:", getCleanDefinition(currentEnd.RoleNote) ) )
		End If
		If currentEnd.Cardinality <> "" Then
			call utvidTabell( tabell, adocTabellRad( "Fra multiplisitet:", currentEnd.Cardinality )  )
		End If
	End If
	
	call skrivTabell( tabell)
	underOverskrift = ""   '' overskrift bare foran den første tabellen
end sub

'-----------------Relasjoner End-----------------



'-----------------Operasjoner-----------------
sub Operasjoner(element, underOverskrift)

	Dim meth as EA.Method
	dim tabell
					
	For Each meth In element.Methods
		tabell = adocTabellstart("20,80", underOverskrift) 
		
		call utvidTabell( tabell, adocTabellHode( "Navn:", meth.Name ) )
		call utvidTabell( tabell, adocTabellRad( "Beskrivelse:", getCleanDefinition(meth.Notes) ) )

''''''''''''''''''''''''''''''''''''''''''''''
'''		call utvidTabell( tabell, adocTabellRad( "Stereotype:", meth.Stereotype) )
'''		call utvidTabell( tabell, adocTabellRad( "Retur type:", meth.ReturnType) )
'''		call utvidTabell( tabell, adocTabellRad( "Oppførsel:", meth.Behaviour) )

		call skrivTabell( tabell)
		underOverskrift = " "
	Next

end sub
'-----------------Operasjoner End-----------------


'-----------------Restriksjoner-----------------
sub Restriksjoner( element, underOverskrift)

	Dim constr as EA.Constraint
	dim tabell
	
	For Each constr In element.Constraints
		tabell = adocTabellstart("20,80", underOverskrift) 
		
		call utvidTabell( tabell, adocTabellHode( "Navn:", Trim(constr.Name) ) )
		call utvidTabell( tabell, adocTabellRad( "Beskrivelse:", getCleanRestriction(constr.Notes) ) )
		
'''''''''''''''''
'''		call utvidTabell( tabell, adocTabellRad( "Type:", constr.Type) )
'''		call utvidTabell( tabell, adocTabellRad( "Status:", constr.Status) )
'''		call utvidTabell( tabell, adocTabellRad( "Vekt:", constr.Weight) )

		call skrivTabell( tabell)
		underOverskrift = " "
	Next

end sub
'-----------------Restriksjoner End-----------------




' -------------------  Bilde av modellelement  ----------------

sub bildeAvModellelement( element)

	dim standardTekst
	dim standardAlternativ(1)
	if isFeatureOrDataType(element)	then
		standardTekst = "Illustrasjon av objekttype "
		standardAlternativ(0) = "Bilde av et eksempel på objekttypen "
'''' 		Tatt bort formuleringen om påtegning av geometri
''		standardAlternativ[1] = ", eventuelt med påtegning av streker som viser hvor geometrien til objektet skal måles fra."
	elseif isCodelist(element) then
		standardTekst = "Illustrasjon av kodeliste: "
		standardAlternativ(0) = "Illustrasjon av hva kodelisten "
		standardAlternativ(1) = " kan inneholde." 
	else
		standardTekst = "Illustrasjon av pakke "
		standardAlternativ(0) = "Bildet viser en illustrasjon av innholdet i UML-pakken "
		standardAlternativ(1) = ", der alle detaljene kommer i teksten nedenfor."
	end if


	dim tag as EA.TaggedValue
	dim bilde, bildetekst, alternativbildetekst

 if isFeatureOrDataType(element) then
''		Tilnærming fra objekttyper og datatyper:
''
	for each tag in element.TaggedValues								
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then

			bilde = getTaggedValue(element, "SOSI_bildeAvModellelement") 
			
			if getTaggedValue(element, "SOSI_bildetekst") <> "" then 
				bildetekst = getTaggedValue(element, "SOSI_bildetekst")
			else
				bildetekst = standardTekst & element.Name & ""
			end if
			
			if getTaggedValue(element,"SOSI_alternativbildetekst") <> "" then 
				alternativbildetekst = getTaggedValue(element, "SOSI_alternativbildetekst")
			else
				alternativbildetekst = standardAlternativ(0) & element.Name 
			end if
			
			call settInnBilde( bildetekst, bilde, alternativbildetekst) 
			
		end if
	next
	
 elseif isCodelist(element)	then
''		Tilnærming fra kodelister:
''
	for each tag in element.TaggedValues										
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then	
			bildetekst = standardTekst & element.Name
			bilde = tag.Value		
			alternativbildetekst = standardAlternativ(0) & element.Name & standardAlternativ(1)

			call settInnBilde( bildetekst, bilde, alternativbildetekst)
		end if
	next
	
 else   '''  Tilnærming fra Pakker

	for each tag in element.TaggedValues
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then

			bilde = getTaggedValue(element, "SOSI_bildeAvModellelement") 
			
			if getTaggedValue(element, "SOSI_bildetekst") <> "" then 
				bildetekst = getTaggedValue(element, "SOSI_bildetekst")
			else
				bildetekst = standardTekst & element.Name & ""
			end if
			
			if getTaggedValue(element,"SOSI_alternativbildetekst") <> "" then 
				alternativbildetekst = getTaggedValue(element, "SOSI_alternativbildetekst")
			else
				alternativbildetekst = standardAlternativ(0) & element.Name & standardAlternativ(1)
			end if
			
			call settInnBilde(bildetekst, bilde, alternativbildetekst) 

		end if
	next

 end if

exit sub
	'' Ny felles tilnærming kan komme her....
	
	bilde = taggedValueFraElement(element, "sosi_bildeavmodellelement")   ''    ????

end sub



'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Func Name: attrbilde(att)
' Author: Kent Jonsrud
' Date: 2021-09-16
' Purpose: skriver ut lenke til bilde av element ved siden av elementet

sub attrbilde(att,typ)
''	call skrivTekst( bildeAvAttributt(att, typ) )	
''exit sub 
	dim tag as EA.TaggedValue
	for each tag in att.TaggedValues								
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			Session.Output(" +")
			Session.Output("Illustrasjon av " & typ & " "&att.Name&"")
			Session.Output("image:"&tag.Value&"[link="&tag.Value&",width=100,height=100, alt=""Bilde av " & typ & " "&att.Name&" som er forklart i teksten.""]")
		end if
	next
end sub

'-------------------------------------------------------------END--------------------------------------------------------------------------------------------

function bildeAvAttributt( att, typ)
	dim tag as EA.TaggedValue
	dim bildetekst, alternativbildetekst, res
	
	for each tag in att.TaggedValues								
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			bildetekst = "Illustrasjon av " & typ & " " & att.Name &""		
			alternativbildetekst = "Bilde av " & typ & " " & att.Name & " som er forklart i teksten."
			res = adocInlineBilde( bildetekst, tag.Value, alternativbildetekst)	

			exit for  '' eller skal det være mulig å ha flere bilder av samme attributt ???
		end if
	next
	
	bildeAvAttributt = res
end function 



'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Func Name: isElement
' Author: Kent Jonsrud
' Date: 2021-07-13
' Purpose: tester om det finnes et element med denne ID-en.

function isElement(ID)
	isElement = false
	if Mid(Repository.SQLQuery("select count(*) from t_object where Object_ID = " & ID & ";"), 113, 1) <> 0 then
		isElement = true
	end if
end function
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'-----------------Funksjon for full path-----------------
function getPath(package)
	dim path
	dim parent
	if package.parentID <> 0 then
'		Session.Output(" -----DEBUG getPath=" & getPath & " package.Name = " & package.Name & " package.ParentID = " & package.ParentID & " package.Element.Stereotype = " & package.Element.Stereotype & " ----- ")
		if package.Element.Stereotype = "" then
			path = package.Name
		else
			path = tekstformatStereotype( package.Element.Stereotype) & package.Name
		end if

		if ucase(package.Element.Stereotype) <> "APPLICATIONSCHEMA" then
			set parent = Repository.GetPackageByID(package.ParentID)
			path = getPath(parent) + "/" + path
		end if
	end if
	getPath = path
end function
'-----------------Funksjon for full path End-----------------


'----------------  Funksjoner for å lese tagged values -------------------
'
function taggedValueFraElement(element, byVal tagName)
	tagName = LCase(tagName)
	for each tag in element.TaggedValues
		if LCase(tag.Name) = tagName and tag.Value <> "" then
			taggedValueFraElement = tag.Value
			exit for
			exit function
		end if
	next
	taggedValueFraElement = ""
end function


function getTaggedValue(element,taggedValueName)
	dim i, existingTaggedValue
	getTaggedValue = ""
	for i = 0 to element.TaggedValues.Count - 1
		set existingTaggedValue = element.TaggedValues.GetAt(i)
		if LCase(existingTaggedValue.Name) = LCase(taggedValueName) then
			getTaggedValue = existingTaggedValue.Value
		end if
	next
end function

function getPackageTaggedValue(package,taggedValueName) 	'''' kan fjernes
	dim i, existingTaggedValue
	getPackageTaggedValue = ""
	for i = 0 to package.element.TaggedValues.Count - 1
		set existingTaggedValue = package.element.TaggedValues.GetAt(i)
		if LCase(existingTaggedValue.Name) = LCase(taggedValueName) then
			getPackageTaggedValue = existingTaggedValue.Value
		end if
	next
end function
'
'----------------  Funksjoner for å lese tagged values End -------------------


'-----------------Function getCleanDefinition Start-----------------
function getCleanDefinition(byVal txt)
	'removes all formatting in notes fields, except crlf
	Dim res, tegn, i, u, forrige
	
	txt = Trimutf8(txt)

	call ErstattTegn( txt, "|", "\|")
	call ErstattTegn( txt, "((", "( (")
	call ErstattTegn( txt, "))", ") )")

	For i = 1 To Len(txt)
		tegn = Mid(txt,i,1)
'		'for adoc \|
'		if tegn = "|" then
'			res = res + "\"
'		end if
'		if tegn = "(" and forrige = "(" then
'			res = res + " "
'		end if
'		if tegn = ")" and forrige = ")" then
'			res = res + " "
'		end if
''			if tegn = "," then tegn = " " 

		'for xml
		If tegn = "<" Then
			u = 1
''			tegn = " "
		end if 
		If tegn = ">" Then
			u = 0
			tegn = " "
		end if
		
		if u = 0 then
			res = res + tegn
		end if

	Next

	getCleanDefinition = res
exit function

	'removes all formatting in notes fields, except crlf
''	Dim res, tegn, i, u, forrige
	u=0
	getCleanDefinition = ""
	forrige = " "
	res = ""
	txt = Trimutf8(txt)

	For i = 1 To Len(txt)
		tegn = Mid(txt,i,1)
		'for adoc \|
		if tegn = "|" then
			res = res + "\"
		end if
		if tegn = "(" and forrige = "(" then
			res = res + " "
		end if
		if tegn = ")" and forrige = ")" then
			res = res + " "
		end if

''			if tegn = "," then tegn = " " 
		'for xml
		If tegn = "<" Then
			u = 1
			tegn = " "
		end if 
		If tegn = ">" Then
			u = 0
			tegn = " "
		end if
		
		if u = 0 then
			res = res + tegn
		end if

		forrige = tegn
	'	Session.Output(" tegn" & tegn)
	Next

	getCleanDefinition = res
end function
'-----------------Function getCleanDefinition End-----------------

'-----------------Function getCleanRestriction Start-----------------
function getCleanRestriction(txt)
	'removes all formatting in notes fields, except crlf
	Dim res, tegn, i, u, forrige, v, kommentarlinje
	kommentarlinje = 0
	u=0
	v=0
	getCleanRestriction = ""
	forrige = " "
		res = ""
		txt = Trimutf8(txt)
		For i = 1 To Len(txt)
			tegn = Mid(txt,i,1)
			'for adoc \|
			if tegn = "|" then
				res = res + "\"
			end if
			if tegn = "(" and forrige = "(" then
				res = res + " "
			end if
			if tegn = ")" and forrige = ")" then
				res = res + " "
			end if
			
	'		if tegn = "-" and forrige <> "-" then
	'			u = 1
	'		end if
			if tegn = "-" then
				if forrige = "-" then
					u = 0
					if kommentarlinje > 0 then
						res = res + " + " + vbCrLf  + "-"
					else
						res = res + vbCrLf  + "-"
					end if
					kommentarlinje = kommentarlinje + 1
					forrige = " "
					v = 1
				else
					u = 1
				end if
			else
				if forrige = "-" and v = 0 then
					res = res + "-"
					u = 0
				end if
				v = 0
			end if

		'	if tegn = "," then tegn = " " 
			'for xml
			If tegn = "<" Then
				u = 1
				tegn = " "
			end if 
			If tegn = ">" Then
				u = 0
				tegn = " "
			end if
			if u = 0 then
				res = res + tegn
			end if

			forrige = tegn
		'	Session.Output(" tegn" & tegn)
		Next

	getCleanRestriction = res
end function
'-----------------Function getCleanRestriction End-----------------

'-----------------Function getCleanBildetekst Start-----------------
function getCleanBildetekst(byVal txt)                                 ''' Ikke i bruk

	txt = getCleanDefinition(txt)
	
	call ErstattTegn( txt, ",", " ")
	
	getCleanBildetekst = txt	
exit function

	'removes all formatting in notes fields, except crlf
	Dim res, tegn, i, u, forrige
	u=0
	getCleanBildetekst = ""
	forrige = " "
	res = ""
	txt = Trimutf8(txt)
	For i = 1 To Len(txt)
		tegn = Mid(txt,i,1)
		'for adoc \|
		if tegn = "|" then
			res = res + "\"
		end if
		if tegn = "(" and forrige = "(" then
			res = res + " "
		end if
		if tegn = ")" and forrige = ")" then
			res = res + " "
		end if
		if tegn = "," then tegn = " " 
		'for xml
		If tegn = "<" Then
			u = 1
			tegn = " "
		end if 
		If tegn = ">" Then
			u = 0
			tegn = " "
		end if
		if u = 0 then
			res = res + tegn
		end if

		forrige = tegn
	'	Session.Output(" tegn" & tegn)
	Next

	getCleanBildetekst = res
end function
'-----------------Function getCleanBildetekst End-----------------


'-----------------Function Trimutf8 Start-----------------
function Trimutf8(txt)
	'convert national characters back to utf8

	Dim inp
	inp = Trim(txt)

	call ErstattKodeMedTegn( inp, 230, "æ")
	call ErstattKodeMedTegn( inp, 248, "ø")
	call ErstattKodeMedTegn( inp, 229, "å")
	call ErstattKodeMedTegn( inp, 198, "Æ")
	call ErstattKodeMedTegn( inp, 216, "Ø")
	call ErstattKodeMedTegn( inp, 197, "Å")
	call ErstattKodeMedTegn( inp, 233, "é")
	
''	call ErstattKodeMedTegn( inp, 167, "§")
	
	Trimutf8 = inp
exit function	
'	dim res, tegn, i, u, ÉéÄäÖöÜü-Áá &#233; forrige &#229;r i samme retning skal den h&#248; prim&#230;rt prim&#230;rt
	
	if InStr(1,inp,"&#230;",0) <> 0 then
		inp = Replace(inp,"&#230;","æ",1,-1,0)
	end if
	if InStr(1,inp,"&#248;",0) <> 0 then
		inp = Replace(inp,"&#248;","ø",1,-1,0)
	end if
	if InStr(1,inp,"&#229;",0) <> 0 then
		inp = Replace(inp,"&#229;","å",1,-1,0)
	end if
	if InStr(1,inp,"&#198;",0) <> 0 then
		inp = Replace(inp,"&#198;","Æ",1,-1,0)
	end if
	if InStr(1,inp,"&#216;",0) <> 0 then
		inp = Replace(inp,"&#216;","Ø",1,-1,0)
	end if
	if InStr(1,inp,"&#197;",0) <> 0 then
		inp = Replace(inp,"&#197;","Å",1,-1,0)
	end if
	if InStr(1,inp,"&#233;",0) <> 0 then
		inp = Replace(inp,"&#233;","é",1,-1,0)
	end if
	Trimutf8 = inp
end function

SUB ErstattKodeMedTegn( txt, byVal tallkode, tegn)
	
	tallkode = "&#" & tallkode & ";"
	if InStr(1, txt, tallkode, 0) <> 0 then
		txt = Replace(txt, tallkode, tegn, 1, -1, 0)
	end if

end SUB

SUB ErstattTegn( txt, tegn, nytttegn)
	
	if InStr(1, txt, tegn, 0) <> 0 then
		txt = Replace(txt, tegn, nytttegn, 1, -1, 0)
	end if

end SUB

'-----------------Function Trimutf8 End-----------------


'-----------------Function nao Start-----------------
function nao()
	' I just want a correct xml timestamp to document when the script was run
	dim m,d,t,min,sek,tm,td,tt,tmin,tsek
	y =  Year(Date) & "-"
	tm = innledendeNull( Month(Date)) & "-"
	td = innledendeNull( Day(Date))   & "T"
	tt = innledendeNull( Hour(Date))   & ":"
	tmin = innledendeNull( Minute(Date)) & ":"
	tsek = innledendeNull( Second(Date)) & "Z"
	
	nao = y & tm & td & tt & tmin & tsek 
exit function 
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

function innledendeNull(n)
	dim res

	res = FormatNumber(n,0,0,0,0)
	if n < 10 then 	res = "0" & res
	if n = 0 then res = "00" 
	
	innledendeNull = res
end function 
'-----------------Function nao End-----------------


'====== Hjelpefunksjoner som er uavhengig av adoc  =================

'--------   Funksjoner som returnerer rein tekst slik den skal se ut i rapporten  ---------------

function bounds( att)
''  Returnerer en formattert tekst som angir nedre og øvre grense for et intervall
	bounds = att.LowerBound & ".." & att.UpperBound
	bounds = "[" & bounds & "]"
end function

function tekstformatStereotype( stereotype)
	if stereotype <> "" then 
		tekstformatStereotype = "«" & stereotype & "» "
	else
		tekstformatStereotype = ""
	end if
end function

function stereotypeNavn( element)
	stereotypeNavn = tekstformatStereotype( element.Stereotype) & element.Name
end function

'--------   Funksjoner som forholder seg til UML-modelllen i EA  ---------------

function isFeatureOrDataType(element)
	dim ster 
	ster = Ucase(element.Stereotype)
	isFeatureOrDataType = (ster = "FEATURETYPE" OR ster = "DATATYPE" OR ster = "UNION" )
end function

function isCodelist(element)
	dim ster 
	ster = Ucase(element.Stereotype)
	isCodelist = (ster = "CODELIST" OR ster = "ENUMERATION" OR element.Type = "Enumeration" )
end function

function isAbstract(element)
	isAbstract = (element.Abstract = 1 and Ucase(element.Stereotype) = "FEATURETYPE" )
end function

function sosiBasistype( attrType)
	'' trenger også å sjekke om typen faktisk er en SOSI basistype TBD
	sosiBasistype = "http://skjema.geonorge.no/SOSI/basistype/" & attrType
end function


' ===========  	Funksjoner og rutiner for adoc-kode  ==================


'' --- Asciidoc-rutiner for dokumentoppdeling
'''
sub settInnSideskift
	call skrivTekst( adocAvsnittSkille( ))
	call skrivTekst( adocKommentar(" *********** Sideskift *********** ") )
	call skrivTekst( adocPageBreak())
	
end sub

function adocPageBreak()
	adocPageBreak = "<<<"
end function


sub settInnSkillelinje
	call skrivTekst( adocSkillelinje)
end sub

function adocSkillelinje
	dim res(2)
	res(0) = adocAvsnittSkille( )
	res(1) = adocKommentar(" ----------- Skillelinje -----------") 
	res(2) = adocBreak()
	
	adocSkillelinje = res
end function

function adocBreak()
	adocBreak = "'''"
end function


function adocKommentar(kommentar)
	adocKommentar = "// " & kommentar
end function

function adocAvsnittSkille( )
	adocAvsnittSkille = " "
end function

function adocLinjeskift( )
	adocLinjeskift = " +"
end function
	

''' --- Ascidoc for overskrifter --------------------------------------
'''
sub skrivUnderOverskrift(elementLevel, overskrift)
	skrivTekst( adocUnderOverskrift(elementLevel, overskrift) )
end sub


function adocUnderOverskrift(elementLevel, overskrift)
''  Returnerer asciidoc-kode til en underoverskrift for en del av beskrivelsen av et element
''  Overskriften skal være ett nivå lavere enn elementet som den er en del av.
''  Den skal være diskret, dvs. ikke skal vises i innholdsfortegnelsen
''	
	dim res(2)
	res(0) = adocAvsnittSkille()
	res(1) = "[discrete]"
	res(2) = adocOverskrift(elementLevel+1, overskrift)   
	
	adocUnderOverskrift = res
end function

function adocVanligOverskrift(elementLevel, overskrift)
''  Returnerer asciidoc-kode til en overskrift for et modellelement
''  Overskriften skal være på samme nivå som elementet den innleder beskrivelsen til.
''  Den skal vises i innholdsfortegnelsen
''	
	dim res(1)
	res(0) = adocAvsnittSkille()
	res(1) = adocOverskrift(elementLevel, overskrift)   
	
	adocVanligOverskrift = res
end function

function adocOverskrift(byVal level, overskrift)
''  En overskrift kan være på nivå 0-5. Den angis med 1-6 "=" før overskriftsteksten 

	if level > 5 then  	level = 5

	if level >= 0 then
		adocOverskrift = string(level+1, "=") & " " & overskrift

	else       ''   level < 0 gir ingen mening...
		adocOverskrift = overskrift
	end if
	
end function



''' --- Ascidoc-funksjoner for linker og referanser  ------------------
'''
function adocBokmerke(element)
	adocBokmerke = "[[" & LCase(element.Name) & "]]"
end function

function targetLink( target)
	targetLink = adocLink( target.Name, stereotypeNavn(target) )
end function

function adocLink( link, tekst)
	adocLink = "<<" & LCase(link) & ", " & tekst & ">>"  
end function

function adocEksternLink( link, tekst)
	adocEksternLink = link & "[" & tekst & "]" 
end function 


''' --- Asciidoc for tabeller
'''

function adocTabellavslutning()
''  Returnrer asciidoc-kode for å avslutte en tabell
	dim res(1)
	res(0) = "|==="	
	res(1) = adocKommentar("Slutt på tabell __________________")
	
	adocTabellavslutning = res
end function


sub skrivTabell(tabell)
	call avsluttTabell( tabell)
	call skrivTekst( tabell)
end sub

sub avsluttTabell(tabell)
	tabell = merge(tabell, adocTabellavslutning )
end sub

sub utvidTabell( tabell, byVal tabellRad)
''  Tabell utvides med tabellrad og returneres
''	Det forutsettes at tabellraden inneholder riktig antall tabellceller
''
	tabell = merge(tabell, tabellRad)
end sub


sub utvidTabellRad( tabellRad, byVal tillegg)
''  Tabellrad utvides med et tillegg og returneres
''	Enten vil tillegg representerte ei ny tabellcelle som legges til raden
''	Eller tekst som legges til i siste celle av raden
''
	tabellRad = merge(tabellRad, tillegg)
end sub

function adocTabellstart( kolonneBredder, overskrift )
''  Returnerer asciidoc-kode for å opprette en tabell med overskrift og angitte kolonnebredder
	
	dim topp(2)
	topp(0) = adocKommentar("Topp av tabell __________________")
	topp(1) = "[cols=""" & kolonneBredder & """]"
	topp(2) = "|==="

	if  isArray(overskrift) then
		adocTabellstart = merge( overskrift, topp)
	elseif overskrift <> "" then 
		adocTabellstart = merge( overskrift, topp)
	else
		adocTabellstart = topp 
	end if

end function


function adocTabellHode( parameter, verdi)
    adocTabellHode = adocTabellRad( adocBold(parameter), adocBold(verdi) ) 
end function

function adocTabellHode3( parameter, verdi, ekstra)
    adocTabellHode3 = adocTabellRad3( adocBold(parameter), adocBold(verdi), adocBold(ekstra) ) 
end function

function adocTabellRad( parameter, verdi)
''  Returnerer asciidoc-kode for å skive ut en rad i en tabell med to kolonner

	dim res(1)
	res(0) = "|" & parameter & " "
	res(1) = "|" & verdi & " "

	adocTabellRad = res
end function

function adocTabellRad3( parameter, verdi, ekstra)
''  Returnerer asciidoc-kode for å skive ut en rad i en tabell med tre kolonner

	dim res(2)
	res(0) = "|" & parameter & " "
	res(1) = "|" & verdi & " "
	res(2) = "|" & ekstra & " "

	adocTabellRad3 = res
end function

function adocTabellCelle( innhold)
	adocTabellCelle = "|" & innhold & " "
end function


''' Asciidoc for bilder
'''
sub settInnBilde(bildetekst, bilde, alternativbildetekst)
''	Setter inn et bilde i full størrelse med innledende skillelinje

	call settInnSkillelinje

	call skrivTekst(adocBilde(bildetekst, bilde, alternativbildetekst) )	
end sub


function adocBilde(bildetekst, bilde, alternativbildetekst)
	dim res(), size
	size = ""
	redim res(2)
'''	res(0) = "." & bildetekst
	res(0) = adocBildeTekst(bildetekst)
	res(1) = adocBildeLink(bilde, alternativbildetekst, size)
	res(2) = adocAvsnittSkille( )
	
	adocBilde = res
end function

function adocInlineBilde( bildetekst, bilde, alternativbildetekst)
	dim res(2)
	res(0) = " +"
	res(1) = bildetekst
	res(2) = adocBildeLink( bilde, alternativbildetekst, "width=100")

	adocInlineBilde = res
end function 

function adocBildeTekst(tekst)
	adocBildeTekst = "." & tekst
end function

function adocBildeLink(bilde, alternativbildetekst, imagesize )
''		adocBildeLink = "image::" & bilde & "[link=" & bilde & ", alt=""" & alternativbildetekst & """]"

	dim res
	res =       "image::" & bilde 
	res = res & "[link=" & bilde 
	if imagesize <> "" then res = res & ", " & imagesize 
	res = res & ", alt=""" & alternativbildetekst 
	res = res & """]"
	
	adocBildeLink = res

end function


''' Ascidoc-funksjoner for tekstformatering
'''
function adocBold( tekst)
''	Returnerer asciidoc-kode for feit/bold tekst
	adocBold = "**" & tekst & "**"
end function 

function adocKursiv( tekst)
''	Returnerer asciidoc-kode for kursiv tekst
	adocKursiv = "__" & tekst & "__"
end function 


''' --- Ascidoc for definisjoner
'''
function adocDefinisjonsAvsnitt( element)
	adocDefinisjonsAvsnitt = adocBold("Definisjon:") & " " & getCleanDefinition(element.Notes)
end function 



' =============  	Rutiner for oppsamling av tekst og utskrift til output    =============

''' --- Rutiner og funksjoner for aggregering av tekst før utskrift
'''

function merge( ByVal list, byVal tillegg)
''  Slår sammen to variabler, gjerne arrayer, til en ny array
''
	if not isArray(list) Then list = array(list)
	if not isArray(tillegg) Then tillegg = array(tillegg)

	dim i, start
	dim res()
	redim res(UBound(list))
	for i = 0 to UBound(list) 
		res(i) = list(i)
	next
	start = UBound(res) + 1

	REDIM preserve res(start + UBound(tillegg) )
	for i = 0 to UBound(tillegg) 
		res(start + i) = tillegg(i)
	next
	
	merge = res
end function


''' --- Rutine for utskrift av tekstlinjer
'''
'''  Utskrift til fil kan legges inn i denne delen

sub skrivTekst(byVal tekst)
	if not isArray(tekst) then 
		call skrivTekstlinje(tekst)
	else
		dim t
		for each t in tekst
			skrivTekstlinje(t)
		next
	end if 
end sub

sub skrivTekstlinje(tekst)
	if tekst <> "" then Session.Output(tekst)
end sub

'====================================================

OnProjectBrowserScript

