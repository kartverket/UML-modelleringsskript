Option Explicit

!INC Local Scripts.EAConstants-VBScript
'' !INC adocHjelpefunksjoner
'' !INC dokumentHjelpefunksjoner


' Script Name: listAdocFraModell
' Author: Tore Johnsen, Åsmund Tjora
' Purpose: Generate documentation in AsciiDoc syntax
' Original Date: 08.04.2021
'
' 
' Versjon: 0.33-1 Dato: 2024-01-05 Jostein Amlien: Omgruppert fila slik at rutinene er gruppert i moduler, med overskrifter. Ingen endring av koden.
' Versjon: 0.33 Dato: 2023-03-01 Jostein Amlien: Ny funksjonalitet: pakkeavhengigheter, eksterne modellelementer, assosisasjoner og aggregeringer, basisTyper. Rydding i kode.
' Versjon: 0.32 Dato: 2023-01-31 Jostein Amlien: Refaktorering. Sjekk av sosiBasisTyper og definisjonstekster. Prefiks av bokmerke, pakkeoverskrifter, filtrere tagger. 
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
'		Definisjon: Mulighet for å koble matrikkelenhet til objekt i SSR for å oppdatere bruksnavn i matrikkelen.
' TBD: opprydding !!!
'
DIM rootId
DIM prefiksBokmerke
Dim imgfolder, imgparent
Dim imgFSO

''  ----------------------------------------------------------------------------

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
			rootId = thePackage.PackageId
			
			imgfolder = "diagrammer"
			Set imgFSO=CreateObject("Scripting.FileSystemObject")
			imgparent = imgFSO.GetParentFolderName(Repository.ConnectionString())  & "\" & imgfolder
			if not imgFSO.FolderExists(imgparent) then
				imgFSO.CreateFolder imgparent
			end if
			
'''			imgparent = ""   ''' brukes ved ny kjøring der alle diagrammer allerede er produseret
			
'''			prefiksBokmerke = "XY"   ''' kan brukes for å reduserte flertydige bokmerker der dokumentet skal inngå i et samledokuemet 
'''
			Session.Output("// Start of UML-model")
			Dim topplevel  			
'''			topplevel = 1    '''  foreslått endring 
			topplevel = 2
			
			Call ListPakke( topplevel, thePackage)
			Session.Output("// End of UML-model")
		Case Else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK

	End Select
	Set imgFSO = Nothing
End Sub

''  ----------------------------------------------------------------------------

Sub ListPakke(pakkelevel, thePackage)

	dim pakkeElement
	set pakkeElement = thePackage.Element
	
'----------------- Overskrift og beskrivelse -----------------

	dim overskrift, prefiks

If false then   ''''''''' Erstatta av linjene under
	if pakkeElement.Stereotype = "" then

		Call settInnSideskift                 ''' Hvorfor det ?
		call settInnSkillelinje

		if pakkelevel >= 4 then
			prefiks = "Underpakke: " 
		else
			prefiks = "Pakke: " 
		end if
	else
		prefiks = "Pakke: " 
	end if

	overskrift = prefiks + stereotypeNavn( pakkeElement)
else	 ''''''''''''  erstatter linjene over
	
	call settInnSkillelinje	

	if pakkeElement.Stereotype <> "" then
		overskrift = stereotypeNavn( pakkeElement)
	else
		overskrift = "Pakke: " + pathTilInternPakke(thePackage.PackageID)
''		overskrift = "Pakke: " + pathTilInternPakke(thePackage.ParentID) + "::" + thePackage.Name

	end if

end if

	call skrivTekst( adocOverskrift( pakkelevel, overskrift ) )
	
	call skrivTekst( adocDefinisjonsAvsnitt( pakkeElement) )
 
	call skrivProfilParameterTabell( pakkelevel+1, pakkeElement)   '' øker med ett nivå 

  
	call Pakkeavhengigheter( pakkeElement, adocUnderOverskrift( pakkelevel+1, "Avhengigheter") )   

'----------------- Bilder og diagram-----------------

	call bildeAvModellelement( pakkeElement)	


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
			alternativbildetekst = getCleanDefinition(diag.Notes)  '' Gjenstår: fjerne avsnitt i teksten
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

'	ALT 4  Alle undeliggende pakker skrives ut på nivå 3
'	nestelevel = 3

	dim pack as EA.Package
	for each pack in thePackage.Packages
		Call ListPakke(nestelevel, pack)
	next

end sub

''  ----------------------------------------------------------------------------

sub Pakkeavhengigheter( element, tabellOverskrift)

	Dim con As EA.Connector
	dim target As EA.Element	

	dim realCelle    : realCelle = adocTabellCelle("") 		'' Tabellcelle for hvor elementet er realisert fra
	dim depSuppCelle : depSuppCelle = adocTabellCelle("") 	'' Tabellcelle for avhengigheter  
	
	DIM pakkeReferanse, elementReferanse, targetReferanse
	For Each con In element.Connectors
		'' En connector peker fra client til supplier, fx:
			'' Client er realisert fra eller avhengig av Supplier
			
		'' La Target være i andre enden av connector sett fra element
		if element.ElementID = con.ClientID then
			set target = Repository.GetElementByID(con.SupplierID)
		elseif element.ElementID = con.SupplierID Then 
			set target = Repository.GetElementByID(con.ClientID)
		end if
		'''  ASSERT target.Type = "Package" AND element.Type = "Package"
		
		pakkeReferanse = pathTilEksternPakke(target.PackageID) 
		elementReferanse = adocUnderstrek(stereotypeNavn(target))
		targetReferanse = pakkeReferanse + "::" + elementReferanse	

		if element.ElementID = con.ClientID then
			'' Elementet peker på target
			if con.Type = "Dependency" then
				call utvidTabellCelle( depSuppCelle, targetReferanse )	
			elseif con.Type = "Realisation" then
				call utvidTabellCelle( realCelle, targetReferanse )	
			end if
		elseif element.ElementID = con.SupplierID then
			'' Elementet pekes på av target
			'' Det er ikke krav om å dokumentere dette 
			targetReferanse = ""
		end if		
	Next
	
	dim tabell 
	tabell = adocTabellstart("20,80", tabellOverskrift)

	if isArray( realCelle) then
		call utvidTabell( tabell, merge( adocTabellCelle("Realisert fra:"), realCelle) )
	end if
	if isArray( depSuppCelle) then   
		call utvidTabell( tabell, merge( adocTabellCelle("Avhengig av:"), depSuppCelle) )
	end if

	if isArray( realCelle) or isArray( depSuppCelle) then call skrivTabell( tabell)
	
End sub   


''  ----------------------------------------------------------------------------

'-----------------ObjektOgDatatyper med ArvOgRealisering -----------------

Sub ObjektOgDatatyper(elementLevel, element, pakke)

	call skrivTekst( elementOverskrift(elementLevel, element, pakke) )
	call skrivTekst( adocDefinisjonsAvsnitt( element) )
	call skrivProfilParameterTabell( elementLevel, element) 
	
	call bildeAvModellelement( element)

	call skrivUnderOverskrift(elementLevel, "Egenskaper")
'''	if element.Attributes.Count = 0 then
'''		skrivTekst("Inneholder ingen egenskaper")
'''	end if
	Dim att As EA.Attribute
	for each att in element.Attributes
		call skrivTabell( attributtbeskrivelse(att) )
	next

	call Relasjoner( element, adocUnderOverskrift(elementLevel, "Roller") )
	
	call Operasjoner( element, adocUnderOverskrift(elementLevel, "Operasjoner") )
	
	call Restriksjoner( element, adocUnderOverskrift(elementLevel, "Restriksjoner") )
	
	call ArvOgRealiseringer( element, adocUnderOverskrift( elementLevel, "Arv og realiseringer") )
'	call ArvOgRealiseringer( element, adocUnderOverskrift( elementLevel, "Avhengigheter") )

end sub

''  ----------------------------------------------------------------------------

sub ArvOgRealiseringer( element, tabellOverskrift)
	 
	Dim con As EA.Connector
	dim target As EA.Element
	
	dim superCelle : superCelle = adocTabellCelle("")		'' Tabellcelle for elementets supertype(r) 
 	dim subCelle   : subCelle = adocTabellCelle("") 			'' Tabellcelle for elementets subtyper
	dim realCelle  : realCelle = adocTabellCelle("") 		'' Tabellcelle for hvor elementet er realisert fra

	DIM pakkeReferanse, elementReferanse, targetReferanse

	For Each con In element.Connectors
		'' En connector peker fra client til supplier, fx:
			'' client er realisert fra supplier
			'' client er en subtype av supplier

		dim targetID		
		'' La target være i andre enden av connector sett fra element
		if element.ElementID = con.ClientID then
			targetID = con.SupplierID
		elseif element.ElementID = con.SupplierID Then 
			targetID = con.ClientID
		else
			EXIT SUB
		end if
		set target = Repository.GetElementByID(targetID)
		
		If con.Type = "Generalization" then 
			''	Hovedregelen for generalisering er at target er intern 
			''  TBD: Ta høyde for at supertypen er ekstern
			elementReferanse = targetLink(target)
			if targetID = con.SupplierID then
				'' vis supertyper med pakkereferanse			
'				pakkeReferanse = pathTilInternPakke(target.PackageID) 
'				targetReferanse = pakkeReferanse + "::" + elementReferanse
'				call utvidTabellCelle( superCelle, targetReferanse )	
				call utvidTabellCelle( superCelle, pathTilInterntElement(target) )	
			else
				'' vis subtyper uten pakkereferanse
				'' Forutsett at subtypen er i ei pakke under samme skjema
				call utvidTabellCelle( subCelle, elementReferanse)
			end if
		elseIf con.Type = "Realisation" then
			''	Hovedregelen for realisering er at target er ekstern 
			if targetID = con.SupplierID then
				'' Vis hvor elementet er realisert fra
				pakkeReferanse = pathTilEksternPakke(target.PackageID) 
				elementReferanse = adocUnderstrek(stereotypeNavn(target))
				targetReferanse = pakkeReferanse + "::" + elementReferanse
				call utvidTabellCelle( realCelle, targetReferanse)	
			else 
				'' Ikke påkrevd å vise hvor et elementet er blitt realisert
			end if
		elseif false then
			'' Andre avhengigheter 
			elementReferanse = target.Type + " " + pathTilEksternPakke(target.PackageID)  
			targetReferanse = elementReferanse + "::" + adocUnderstrek(stereotypeNavn(target))
			
			dim connCelle  : connCelle = adocTabellCelle("")	
			if targetID = con.SupplierID  Then	'' element er avhengig av supplier
				call utvidTabellCelle( connCelle, "Til " & targetReferanse)
			elseif targetID = con.SupplierID then '' element er avhengig av client 
				call utvidTabellCelle( connCelle, "Fra " & targetReferanse)
			end if		
		end if

	Next
	
	dim tabell : tabell = adocTabellstart("20,80", tabellOverskrift)

	if isArray( superCelle) then
		call utvidTabell( tabell, merge( adocTabellCelle("Supertype:"), superCelle))
	end if
	if isArray( subCelle) then
		call utvidTabell( tabell, merge( adocTabellCelle("Subtyper:"), subCelle))
	end if
	if isArray( realCelle) then
		call utvidTabell( tabell, merge( adocTabellCelle("Realisert fra:"), realCelle))
	end if

	if isArray( superCelle) or isArray( subCelle) or isArray( realCelle) then 
		call skrivTabell( tabell)
	end if
End sub   


'-----------------ObjektOgDatatyper / ArvOgRealiseringer   End-----------------


''  ----------------------------------------------------------------------------
' 					Kodelister  		
''  ----------------------------------------------------------------------------

Sub Kodelister(elementLevel, element, pakke)

	call skrivTekst( elementOverskrift(elementLevel, element, pakke)  )
	call skrivTekst( adocDefinisjonsAvsnitt( element) )
	call skrivProfilParameterTabell( elementLevel, element) 

	call bildeAvModellelement( element )

	if element.Attributes.Count > 0 then  
		dim tabelloverskrift
		tabelloverskrift = adocUnderOverskrift(elementLevel, "Koder i modellen") 			 ''*****************''
	
		CALL skrivTabell( modellkoder(element, tabelloverskrift ) )
	else
	''' Da må kodelista være ekstern ....
	end if
	
end sub

''  ----------------------------------------------------------------------------

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
'''		IF att.Notes = "" THEN tabellRad = adocTabellRad( att.Name, adocBold("ADVARSEL: KODEDEFINISJON MANGLER"))  '''
		if utvekslingsalias then call utvidTabellRad( tabellRad, adocTabellCelle(att.Default) )
		call utvidTabellRad( tabellRad, bildeAvAttributt(att, "kodelistekode")  )
		
		call utvidTabell( tabell, tabellrad )
	next
	
	modellkoder = tabell
End function

'-----------------CodeList End-----------------

''  ----------------------------------------------------------------------------

sub skrivProfilParameterTabell( elementLevel, element) 
	dim overskrift, tabell, listTags, tag

	overskrift = adocUnderOverskrift(elementLevel, "Profilparametre i tagged values")
	tabell = adocTabellstart("20,80", overskrift)
	listTags = false
	for each tag in element.TaggedValues
		if tag.Value = "" then	
		elseif tag.Name = "persistence" or tag.Name = "SOSI_melding" then  	''  hopp over disse
''		elseif tag.Name = "SOSI_navn" then  	''  hopp over denne også
		elseif tag.Name = "byValuePropertyType" OR  tag.Name = "isCollection" OR  tag.Name = "noPropertyType" then ''hopp over
		elseif tag.Name = "asDictionary" AND tag.Value = "false" then ''hopp over
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

''  ----------------------------------------------------------------------------

function attributtbeskrivelse( att)

	dim tabell, tabellRad, tabellCelle
	tabell = adocTabellstart("20,80", "")
	
	call utvidTabell( tabell, adocTabellHode( "Navn:", att.Name ) )

	dim definisjon
	definisjon = getCleanDefinition(att.Notes)
'''	If definisjon = "" Then definisjon = adocBold("ADVARSEL: EGENSKAPSDEFINISJON MANGLER")   '''
	tabellRad = adocTabellRad( "Definisjon:", definisjon) 
	call utvidTabell( tabell, tabellRad )

	call utvidTabell( tabell, adocTabellRad( "Multiplisitet:", bounds(att) ) )
	
	if att.Default <> "" then
		call utvidTabell( tabell, adocTabellRad( "Initialverdi:", att.Default ) )
	end if

	if not att.Visibility = "Public" then
		call utvidTabell( tabell, adocTabellRad( "Visibilitet:", att.Visibility ) )
	end if
	

	call utvidTabell( tabell, adocTabellRad( "Type:", attributtype(att)	) )


	tabellCelle = adocTabellCelle( "")
	dim tag as EA.TaggedValue
	for each tag in att.TaggedValues
		if tag.Value = "" then												'' 	hopp over tomme tagger
		elseif tag.Name = "persistence" or tag.Name = "SOSI_melding" then  	''  hopp over disse også
		elseif LCase(tag.Name) = "sosi_bildeavmodellelement" then 			''  tas separat, hopp over
''		elseif LCase(tag.Name) = "sosi_navn" then 							''  hopp over
		else
			call utvidTabellCelle( tabellCelle, tag.Name & ": " & tag.Value )
		end if
	next
	if isArray( tabellCelle) then 
		tabellRad = merge( adocTabellCelle("Profilparametre i tagged values: "), tabellCelle)
		call utvidTabell( tabell, tabellrad)
	end if
	attributtbeskrivelse = tabell
end function


''  ----------------------------------------------------------------------------
''				Operasjoner og restriksjoner
''  ----------------------------------------------------------------------------

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

''  ----------------------------------------------------------------------------

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

'	-----------------	Operasjoner og Restriksjoner 	End	--------------------




'===============================================================================
'
'		MODUL for relasjoner og Assosiasjonsroller   ################
'
'===============================================================================

''  ----------------------------------------------------------------------------
'					Relasjoner		
''  ----------------------------------------------------------------------------
'
sub Relasjoner( element, underOverskrift)
	
	dim tabell 
	tabell = adocTabellstart("20,80", underOverskrift)	
'assosiasjoner
' skriv ut roller - sortert etter tagged value sequenceNumber TBD

	Dim con
	For Each con In element.Connectors
		If con.Type = "Association" or con.Type = "Aggregation" Then

			'' TBD:	Egenassosiasjoner.
			'	Kan løses ved å erstatte elseif med if under
			'   Eller legge det inn som et innledende særtilfelle
			
			'' Det er hensiktsmesig å rapportere aggregreringene i modellen slik EA velger å framstille dem i diagrammene
			call fiksKonnektor( con)

			dim aggType
			dim realiserRolle : realiserRolle = false
			If element.elementID = con.SupplierID  and element.elementID = con.ClientID Then
				'' Egenassosiasjon
				realiserRolle = TRUE
				tabell = beskrivRolle( tabell, con.ClientEnd,   con.clientID,   aggregationType(con.SupplierEnd))
				tabell = beskrivRolle( tabell, con.SupplierEnd, con.supplierID, aggregationType(con.ClientEnd))
				tabell = beskrivKonnektor( tabell, con)

			elseIf element.elementID = con.SupplierID  Then 
				'dette elementet er suppliersiden - implisitt at fraklasse er denne klassen

				aggType = aggregationType(con.SupplierEnd)
				realiserRolle = realiserbarRolle(con.ClientEnd, aggType) 
				if realiserRolle then
				''	if aggType = "Assosiasjon" and con.Type = "Aggregation" then aggType = adocKursiv("Assosiasjon")  

					tabell = beskrivRolle( tabell, con.ClientEnd, con.clientID, aggType)
					tabell = beskrivKonnektor( tabell, con)
					if false then tabell = beskrivCurrentEnd( tabell, con.SupplierEnd )
				end if
			elseIf element.elementID = con.ClientID  Then
				'dette elementet er clientsiden, (rollen er på target)
				
				aggType = aggregationType(con.ClientEnd)
				realiserRolle = realiserbarRolle(con.SupplierEnd, aggType) 
				if realiserRolle then
				''	if aggType = "Assosiasjon" and con.Type = "Aggregation" then aggType = adocKursiv("Assosiasjon")  
				
					tabell = beskrivRolle( tabell, con.SupplierEnd, con.supplierID, aggType)
					tabell = beskrivKonnektor( tabell, con)
					if false then tabell = beskrivCurrentEnd( tabell, con.ClientEnd )
				end if
			End If	

			if realiserRolle then
				call skrivTabell( tabell)

				''  Initier neste tabell, uten overskrift
				tabell = adocTabellstart("20,80", "")	
			end if
		End If
	Next
end sub

''  ----------------------------------------------------------------------------

function realiserbarRolle( r, aggType)

	realiserbarRolle = ( r.Role <> "" or r.RoleNote <> "" or r.Cardinality <> "" or r.Navigable = "Navigable" or aggType <> "Assosiasjon" )

end function

''  ----------------------------------------------------------------------------

function beskrivRolle( byVal tabell, targetEnd, targetID, aggregeringsType)
''
''	targetEnd angir Rollen: navn, def, multiplisitet og navigerbarhet
''	targetID angir klassen som Rollen peker på
''	aggregeringsType 
''	
''	tabell inneholder en initiert tabell, med eller uten overskrift 
''

	dim target : set target = Repository.GetElementByID(targetID)

	call utvidTabell( tabell, adocTabellHode( "Rollenavn:", targetEnd.Role ) )

	dim definisjon
	definisjon = getCleanDefinition(targetEnd.RoleNote)
''	If definisjon = "" Then	definisjon = adocBold("ADVARSEL: ROLLEDEFINISJON MANGLER")   '''
	call utvidTabell( tabell, adocTabellRad( "Definisjon:", definisjon )  )


	dim multiplisitet 	: multiplisitet = targetEnd.Cardinality    '''' tekstformat
	if multiplisitet <> "" then 
		multiplisitet = "[" + targetEnd.Cardinality + "]"
		call utvidTabell( tabell, adocTabellRad( "Multiplisitet:", multiplisitet ) )
''	else
''		call utvidTabell( tabell, adocTabellRad( "Multiplisitet:", adocBold("1") ) )
	end if


	call utvidTabell( tabell, adocTabellRad( "Assosiasjonstype:", aggregeringsType) )     ''''  ****************
''	call utvidTabell( tabell, adocTabellRad( "Aggregeringstype:", aggregeringsType) )     ''''  ****************
	
	dim textVar 		: textVar = "Til klasse"
	dim navigerbarhet 	: navigerbarhet = targetEnd.Navigable
	'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
	If navigerbarhet = "Navigable" Then 
		textVar = textVar + adocKursiv(" (navigerbar):") 
''		textVar = "Navigerbar til:"
	ElseIf navigerbarhet = "Non-Navigable" Then 
		textVar = textVar + adocKursiv(" (ikke navigerbar):") 
''		textVar = "Ikke navigerbar til:"		
'	Elseif navigerbarhet = "Unspecified" Then 
'		textVar = textVar + + adocBold(" (Uspesifisert):") 
''		textVar = "Til klasse"
	Else 
		textVar = textVar + ":" 
''		textVar = "Til klasse:"
	End If
	
'''	dim pakkeReferanse : pakkeReferanse = pathTilEksternPakke( target.PackageID)
'	dim pakkeReferanse : pakkeReferanse = pathTilInternPakke( target.PackageID)
'	dim elementReferanse : elementReferanse = targetLink( target )
'	dim targetReferanse : targetReferanse = pakkeReferanse + "::" + elementReferanse
'
'	call utvidTabell( tabell, adocTabellRad( textVar, targetReferanse ) )

	call utvidTabell( tabell, adocTabellRad( textVar, pathTilInterntElement(target) ) )

	beskrivRolle = tabell
end function

''  ----------------------------------------------------------------------------

function beskrivKonnektor( byVal tabell, connector)
	dim konnNavn : 	konnNavn = stereotypeNavn(connector)
	
	if konnNavn <> "" then call utvidTabell( tabell, adocTabellRad( "Konnektor: ", konnNavn) )

	if connector.Type = "Aggregation" or konnNavn <> "" then
	'	call utvidTabell( tabell, adocTabellRad( "Assosiasjonstype:", connector.Type) )
		call utvidTabell( tabell, adocTabellRad( "Konnektortype:", connector.Type ) ) '' Trenger vi denne ?
	end if
	
	beskrivKonnektor = tabell
end function

''  ----------------------------------------------------------------------------

function beskrivCurrentEnd( byVal tabell, currentEnd )

	If currentEnd.Role <> "" Then
		call utvidTabell( tabell, adocTabellRad( "Fra rolle:", currentEnd.Role ) )
	End If
	If currentEnd.RoleNote <> "" Then
		call utvidTabell( tabell, adocTabellRad( "Fra rolle definisjon:", getCleanDefinition(currentEnd.RoleNote) ) )
	End If
	If currentEnd.Cardinality <> "" Then
		call utvidTabell( tabell, adocTabellRad( "Fra multiplisitet:", currentEnd.Cardinality )  )
	End If

	beskrivCurrentEnd = tabell
end function

''  ----------------------------------------------------------------------------

sub fiksKonnektor( connector)
	'' Denne funksjonen gjennomfører de samme tilpasnignene som EA gjør i diagrammene

	if connector.Type = "Aggregation" then
		'' En Aggregation kan ikke ha Assosiasjonsroller i begge ender
		if connector.SupplierEnd.Aggregation = 0 and connector.ClientEnd.Aggregation = 0 then
			connector.SupplierEnd.Aggregation = 1
		end if
		
		'' En Aggregation kan ikke ha ensidig retning mot destinasjon (target, supplier)	
		if connector.Direction = "Source -> Destination" then 
			connector.Direction = "Unspecified"
			connector.SupplierEnd.Navigable = "Unspecified" 
		end if
		
	end if
end sub

''  ----------------------------------------------------------------------------

function aggregationType(rolleEnde)

	dim aggType, res
	aggType = rolleEnde.Aggregation
	
	if aggType = 0 then 
		res = "Assosiasjon"
	elseif aggType = 1 then 
		res = "Aggregering"
	elseif aggType = 2 then 
		res = "Komposisjon"
	else
		call skrivTekst("SYSTEMFEIL:  Assosiasjon har en uventa aggregeringstype:" & aggtype)
		exit function
	end if

	aggregationType = res	
end function



'	============================================================================
'
'					MODUL   navigeringEA    
'
'	============================================================================

'--------   Funksjoner som forholder seg til UML-modelllen i EA  ---------------

''  ----------------------------------------------------------------------------
'				Elementtyper
''  ----------------------------------------------------------------------------

function isFeatureOrDataType(element)
	dim ster 
	ster = Ucase(element.Stereotype)
	isFeatureOrDataType = (ster = "FEATURETYPE" OR ster = "DATATYPE" OR ster = "UNION" )
end function

''  ----------------------------------------------------------------------------

function isCodelist(element)
	dim ster 
	ster = Ucase(element.Stereotype)
	isCodelist = (ster = "CODELIST" OR ster = "ENUMERATION" OR element.Type = "Enumeration" )
end function

''  ----------------------------------------------------------------------------

function isAbstract(element)
	isAbstract = (element.Abstract = 1 and Ucase(element.Stereotype) = "FEATURETYPE" )
end function

''  ----------------------------------------------------------------------------

function sosiBasistype( attrType)
	'' trenger å sjekke om typen faktisk er en SOSI basistype
	
	dim basisTyper
	basisTyper = "Date, Time, DateTime, Number, Decimal, Integer, Real, Vector" 
	basisTyper = basisTyper + ", CharacterString, Boolean, URI, Any, Record, LanguageString"
	basisTyper = basisTyper + ", GM_Point, GM_Curve, GM_Surface, GM_Solid"
	basisTyper = basisTyper + ", GM_Primitive, GM_MultiSurface" 
	basisTyper = basisTyper + ", Punkt, Sverm, Kurve, Flate" 

	basisTyper = Split(basisTyper, ", ")

	dim typ, res
	for each typ in basisTyper
		if typ = attrType then 
			res = "http://skjema.geonorge.no/SOSI/basistype/"  & attrType
			exit for
		end if
	next

	sosiBasistype = res
end function

''  ----------------------------------------------------------------------------

function erLovligBasisType( attType) 
	dim lovligetyper
''	lovligetyper = Split("string, integer, date, boolean", ", ") 
	lovligetyper = "string, integer, date, boolean"  '' typer for xml
	lovligetyper = Split(lovligetyper, ", ") 
	
	erLovligBasisType = false
	dim typ
	for each typ in lovligetyper 
		if typ = attType then erLovligBasisType = true
	next
	
end function

''  ----------------------------------------------------------------------------

function attributtype(att)	
	dim typ

	if att.ClassifierID = 0 then  ''	attributtet har ingen referanse til en classifier
		dim uri 
		uri = sosiBasistype( att.Type)
		
		if uri <> "" then
			typ = adocEksternLink( uri, att.Type)
		elseif erLovligBasisType( att.Type) then   '' lovlig ihht f.eks. xml
			typ = att.Type
		else 
			typ = adocBold("Ukjent type: ") + att.Type
		end if

	elseif att.ClassifierID > 0 then  '' referanse til en klasse i modellen
	
		dim classifier as EA.Element   '' for å angi attributtets datatype
		set classifier = Repository.GetElementByID(att.ClassifierID) '' denne forutsetter at ID peker på noe...

		'' Sjekk om classifier skulle være en datatype definert utafor scope
		if erEksternPakke(classifier.PackageID) then
			'' classifier er ekstern og derfor ikke beskrevet i dette dokumentet
			typ = pathTilEksternPakke(classifier.PackageID) + "::" + adocUnderstrek(stereotypeNavn( classifier))
		else
			typ = targetLink( classifier)
		end if
		
	end if
	
	attributtype = typ
end function 


''  ----------------------------------------------------------------------------
'				Referanser til pakker i modellregisteret
''  ----------------------------------------------------------------------------

function erEksternPakke( pakkeID)  '' pakke i et annet skjama
''	dim rootId  '' global

	dim pakke
	set pakke = Repository.GetPackageByID(pakkeID)

	dim  res
	if pakkeID = rootId then  '' Vi har nådd toppen av denne modellen: pakka er lokal
		res = false
	elseif pakke.parentID = rootId then  '' Vi har nådd toppen av denne modellen: pakka er lokal
		res = false
	elseif pakke.name = "SOSI Model" then  '' Vi har nådd toppen av modellregisteret
		res = true
	elseif pakke.parentID = 0 then	'' Vi har nådd toppen av modellregisteret: pakka er ekstern
		res = true
	elseif pakke.Element.Stereotype <> "" then '' Vi har nådd et annet applikasjonskjema: pakka er ekstern
		res = true
	else
		res = erEksternPakke( pakke.ParentID) 
	end if
	
	erEksternPakke = res
end function

''  ----------------------------------------------------------------------------

function pathTilEksternPakke( pakkeID)  '' pakke i et annet skjama

	dim pakke
	set pakke = Repository.GetPackageByID(pakkeID)

	dim pakkenavn, res
	pakkenavn = pakke.name 

	if pakke.parentID = 0 then	'' Vi har nådd toppen av modellregisteret: pakka er ekstern
		res = ""
	elseif pakkenavn = "SOSI Model" then  '' Vi har nådd toppen av modellregisteret
		res = ""
	elseif pakke.Element.Stereotype <> "" then '' Vi har nådd et applikasjonskjema: pakka er ekstern
		res = pakkenavn
	else
		dim path
		path = pathTilEksternPakke( pakke.ParentID) 

		if path <> "" then 
			res = path + "::" + pakkenavn
		else
			res = pakkenavn	
		end if

	end if
	
	pathTilEksternPakke = res
end function

''  ----------------------------------------------------------------------------

function pathTilInterntElement( target)

	pathTilInterntElement = pathTilInternPakke(target.PackageID) + "::" + targetLink(target)

end function

''  ----------------------------------------------------------------------------

function pathTilInternPakke( pakkeID)
''  Denne brukes for å sette overskrift på pakkene
''	dim rootId  '' global

	dim pakke
	set pakke = Repository.GetPackageByID(pakkeID)

	dim pakkenavn, res
	pakkenavn = pakke.name

	if pakke.parentID = rootId then  '' Vi har nådd toppen av denne modellen: pakka er lokal
		res = pakkenavn
	elseif pakke.parentID = 0 then	'' Vi har nådd toppen av modellregisteret: pakka er ekstern
		res = ""
	elseif pakke.Element.Stereotype <> "" then '' Vi har nådd et annet applikasjonskjema: pakka er ekstern
		res = ""
	else
		dim parentPath
		parentPath = pathTilInternPakke( pakke.ParentID) 

		if parentPath = "" then
			res = ""
		else
			res = parentPath + "::" + pakke.element.name
		end if
	end if
	
	pathTilInternPakke = res
end function


''	============================================================================
'					MODUL: taggedValues
'
'		Høsting av tagged values fra modellen  
'
''	============================================================================


'----------------  Funksjoner for å lese tagged values -------------------
'
function taggedValueFraElement(element, byVal tagName)
	tagName = LCase(tagName)
	dim tag
	for each tag in element.TaggedValues
		if LCase(tag.Name) = tagName and tag.Value <> "" then
			taggedValueFraElement = tag.Value
			exit for
		end if
	next

end function

''  ----------------------------------------------------------------------------

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

''  ----------------------------------------------------------------------------

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



''	============================================================================
'
'				MODUL: dokumentHjelpefunksjoner
'		Hjelpefunksjoner for tekst som er uavhengig av adoc  
'
''	============================================================================

''  ----------------------------------------------------------------------------
''  			Utseende av tekst i dokumentet
''  ----------------------------------------------------------------------------


function bounds( att)
''  Returnerer en formattert tekst som angir nedre og øvre grense for et intervall
	bounds = att.LowerBound & ".." & att.UpperBound
	bounds = "[" & bounds & "]"
end function

''  ----------------------------------------------------------------------------

function tekstformatStereotype( stereotype)
	if stereotype <> "" then 
		tekstformatStereotype = "«" & stereotype & "» "
	else
		tekstformatStereotype = ""
	end if
end function

''  ----------------------------------------------------------------------------

function stereotypeNavn( element)
	dim stereo
	stereo = tekstformatStereotype(element.Stereotype) 
	if stereo = "" and element.Type = "Enumeration" then  
		stereo = tekstformatEnumeration( element)  
	end if

	stereotypeNavn = stereo + element.Name
end function

''  ----------------------------------------------------------------------------

function tekstformatEnumeration( element)
	if element.Type = "Enumeration" then  
		tekstformatEnumeration = """Enumeration"" "
	end if
end function

''  ----------------------------------------------------------------------------

function bokmerke(element)
	'' prefiksBokmerke er en global variabel som settes innledningsvis
	bokmerke = prefiksBokmerke + LCase(element.Name)
end function

''  ----------------------------------------------------------------------------

function targetLink( element)
	'' intern hyperlenke i dokumentet som peker til beskrivelsen av et element
	targetLink = adocLink( bokmerke(element), stereotypeNavn(element) )
end function


''  ----------------------------------------------------------------------------

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


''  ----------------------------------------------------------------------------
''				Tabeller
''  ----------------------------------------------------------------------------

sub skrivTabell(tabell)
	call avsluttTabell( tabell)
	call skrivTekst( tabell)
end sub

sub avsluttTabell(tabell)
	tabell = merge(tabell, adocTabellavslutning )
end sub

''  ----------------------------------------------------------------------------

sub utvidTabell( tabell, tabellRad)
''  Tabell utvides med tabellrad 
''	Det forutsettes at tabellraden inneholder riktig antall tabellceller
''
	tabell = merge(tabell, tabellRad)
end sub

sub utvidTabellRad( tabellRad, byVal tabellcelle)
''  Tabellrad utvides med ei ny tabellcelle som legges til raden
''
	tabellRad = merge(tabellRad, tabellcelle)
end sub

sub utvidTabellCelle( celle, ekstraLinje )
''  Tabellcelle utvides med ei ekstraLinje og returneres
''
	if ekstraLinje <> "" then celle = merge( celle, ekstraLinje & adocLinjeskift )
end sub

''  ----------------------------------------------------------------------------


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



''  ----------------------------------------------------------------------------
''  			Aggregering av generert tekst før utskrift
''  ----------------------------------------------------------------------------

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


'  ----------------------------------------------------------------------------
'  		Utskrift av generert tekst
'  			TBD:  Direkte utskrift til fil
'  ----------------------------------------------------------------------------
'
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

''  ----------------------------------------------------------------------------

sub skrivTekstlinje(tekst)
	if tekst <> "" then Session.Output(tekst)
end sub

''  ----------------------------------------------------------------------------



'===============================================================================
'
'		MODUL for Bilder   ################
'
'===============================================================================


''  ----------------------------------------------------------------------------

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


''  ----------------------------------------------------------------------------
' Func Name: attrbilde(att)
' Author: Kent Jonsrud
' Date: 2021-09-16
' Purpose: skriver ut lenke til bilde av element ved siden av elementet
'
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
'
'-------------------------------------------------------------END---------------

''  ----------------------------------------------------------------------------

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

''  ----------------------------------------------------------------------------

sub settInnBilde(bildetekst, bilde, alternativbildetekst)
''	Setter inn et bilde i full størrelse med innledende skillelinje

	call settInnSkillelinje

	call skrivTekst(adocBilde(bildetekst, bilde, alternativbildetekst) )	
end sub



''	============================================================================
'					MODUL: tekstformatering
'		Hjelpefunksjoner for tekst som er uavhengig av adoc  
'
''	============================================================================


''  ----------------------------------------------------------------------------
'
'     brukes kun til debugging av definisjoner
'
''  ----------------------------------------------------------------------------
''' --- Ascidoc for definisjoner
''
function adocDefinisjonsAvsnitt( element)
	dim definisjon, advarsel
	definisjon = getCleanDefinition(element.Notes)
	advarsel = "ADVARSEL: DEFINISJON MANGLER"
''	if definisjon = "" then definisjon = adocBold(advarsel)
	adocDefinisjonsAvsnitt = adocBold("Definisjon:") & " " & definisjon
	
end function 

''  ----------------------------------------------------------------------------
'
'				Rensing av modelltekst som skal brukes i dokumentet
'
''  ----------------------------------------------------------------------------

function getCleanDefinition(byVal txt)
	'removes all formatting in notes fields, except crlf
	Dim res, tegn, i, u
	
	txt = Trimutf8(txt)

	call ErstattTegn( txt, "|", "\|")
	call ErstattTegn( txt, "((", "( (")
	call ErstattTegn( txt, "))", ") )")

	For i = 1 To Len(txt)
		tegn = Mid(txt,i,1)

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

end function

''  ----------------------------------------------------------------------------

function getCleanRestriction( byval txt)
	'removes all formatting in notes fields, except crlf
	Dim res, tegn, i, u, forrige, v, kommentarlinje
	kommentarlinje = 0
	u=0
	v=0
	getCleanRestriction = ""
	forrige = " "
	res = ""
	txt = Trimutf8(txt)
	
	call ErstattTegn( txt, "|", "\|")
	call ErstattTegn( txt, "((", "( (")
	call ErstattTegn( txt, "))", ") )")

	For i = 1 To Len(txt)
		tegn = Mid(txt,i,1)
		
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

	Next

	getCleanRestriction = res
end function

''  ----------------------------------------------------------------------------

function getCleanBildetekst(byVal txt)                                 ''' Ikke i bruk

	dim res
	res = getCleanDefinition(txt)
	
	call ErstattTegn( res, ",", " ")
	
	getCleanBildetekst = res	

end function


''  ----------------------------------------------------------------------------
'							 Trimutf8 	
''  ----------------------------------------------------------------------------

function Trimutf8(txt)
	'convert national characters back to utf8
'	dim res, tegn, i, u, ÉéÄäÖöÜü-Áá &#233; forrige &#229;r i samme retning skal den h&#248; prim&#230;rt prim&#230;rt

	Dim inp
	inp = Trim(txt)

	call ErstattKodeMedTegn( inp, 230, "æ")
	call ErstattKodeMedTegn( inp, 248, "ø")
	call ErstattKodeMedTegn( inp, 229, "å")
	call ErstattKodeMedTegn( inp, 198, "Æ")
	call ErstattKodeMedTegn( inp, 216, "Ø")
	call ErstattKodeMedTegn( inp, 197, "Å")
	call ErstattKodeMedTegn( inp, 233, "é")
	
	call ErstattKodeMedTegn( inp, 167, "§")
	
	Trimutf8 = inp
end function

''  ----------------------------------------------------------------------------

SUB ErstattKodeMedTegn( txt, byVal tallkode, tegn)
	
	tallkode = "&#" & tallkode & ";"
	if InStr(1, txt, tallkode, 0) <> 0 then
		txt = Replace(txt, tallkode, tegn, 1, -1, 0)
	end if

end SUB

''  ----------------------------------------------------------------------------

SUB ErstattTegn( txt, tegn, nytttegn)
	
	if InStr(1, txt, tegn, 0) <> 0 then
		txt = Replace(txt, tegn, nytttegn, 1, -1, 0)
	end if

end SUB


''  ----------------------------------------------------------------------------
''				Timestamp
''  ----------------------------------------------------------------------------

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
end function

''  ----------------------------------------------------------------------------

function innledendeNull(n)
	dim res

	res = FormatNumber(n,0,0,0,0)
	if n < 10 then 	res = "0" & res
	if n = 0 then res = "00" 
	
	innledendeNull = res
end function 

''  ----------------------------------------------------------------------------


''	============================================================================
'							MODUL adocSyntaks
'
' 	Tilgjengelige tekstutformingsfunksjoner ihht. adoc-syntaksen   #############
'
''	============================================================================

''  ----------------------------------------------------------------------------
'
'  	Formatering av tekst (ord og fraser):
'		bold, kursiv, understrek, bokstavlig, bokstavligCelle
'
''  ----------------------------------------------------------------------------

function adocBold( tekst)
''	Returnerer asciidoc-kode for feit/bold tekst
''
	adocBold = adocFormat( tekst, "**", "")
end function 

function adocKursiv( tekst)
''	Returnerer asciidoc-kode for kursiv tekst
''
	adocKursiv = adocFormat( tekst, "__", "")
end function 

function adocUnderstrek( tekst)
''	Returnerer asciidoc-kode for understreka tekst
''
	adocUnderstrek = adocFormat( tekst, "##", "underline")
end function 

function adocFormat( tekst, format, rolle)
''	Returnerer asciidoc-kode for formattert tekst
	if tekst = "" then
		adocFormat = tekst
	elseif rolle = "" then
		adocFormat = format + tekst + format
	elseif format = "#" or format = "##" then
		adocFormat = "[." + rolle + "]" + format + tekst + format
	else
		adocFormat = format + tekst + format
	end if
end function 


''  ----------------------------------------------------------------------------
'
'	Innsetting av bilder 
'
''  ----------------------------------------------------------------------------

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

function adocBildeTekst(tekst)
	adocBildeTekst = "." + tekst
end function



' 	----------------------------------------------------------------------------
'
'	Ombrekking
'
'	----------------------------------------------------------------------------

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


function adocAvsnittSkille( )
	adocAvsnittSkille = " "
end function

function adocLinjeskift( )
	adocLinjeskift = " +"
end function
	
	
'	----------------------------------------------------------------------------
'
'	Overskrifter
'
'	----------------------------------------------------------------------------

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
''  En overskrift kan være på nivå 0-5. Den angis med 1-6 stk "=" før overskriftsteksten 

	if level > 5 then  	level = 5

	if level >= 0 then
		adocOverskrift = string(level+1, "=") & " " & overskrift

	else       ''   level < 0 gir ingen mening...
		adocOverskrift = overskrift
	end if
	
end function


'	----------------------------------------------------------------------------
'	
'	Lenker og referanser  
'
'	----------------------------------------------------------------------------

''' --- Ascidoc-funksjoner for linker og referanser  ------------------
'''

function adocLink( link, tekst)
	adocLink = "<<" + link + ", " + tekst + ">>"  
end function

''  ----------------------------------------------------------------------------

function adocEksternLink( link, tekst)
	adocEksternLink = link + "[" + tekst + "]" 
end function 

''  ----------------------------------------------------------------------------

function adocBokmerke(element)
	adocBokmerke = "[[" + bokmerke(element) + "]]"
end function

'	----------------------------------------------------------------------------
'
'	Tabeller  
'
'	----------------------------------------------------------------------------

''' --- Asciidoc for tabeller
'''
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

''  ----------------------------------------------------------------------------

function adocTabellavslutning()
''  Returnerer asciidoc-kode for å avslutte en tabell
	dim res(1)
	res(0) = "|==="	
	res(1) = adocKommentar("Slutt på tabell __________________")
	
	adocTabellavslutning = res
end function

''  ----------------------------------------------------------------------------

function adocTabellCelle( innhold)
	adocTabellCelle = "|" & innhold & " "
end function


''  ----------------------------------------------------------------------------

''  ----------------------------------------------------------------------------

function adocKommentar(kommentar)
	adocKommentar = "// " & kommentar
end function


'====================================================

OnProjectBrowserScript

'====================================================


