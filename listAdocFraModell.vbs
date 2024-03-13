Option Explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: listAdocFraModell
' Author: Tore Johnsen, Åsmund Tjora
' Purpose: Generate documentation in AsciiDoc syntax
' Original Date: 08.04.2021
'
' Versjon: 0.37 Dato: 2024-14-03 Jostein Amlien: Gjennomgang av globale parametre. Modul for forflata beskrivelse av realiserte objekttyper til egen fil.
' Versjon: 0.36 Dato: 2024-02-02 Jostein Amlien: Rapport skrives til fil. Feilretting i utskift av rolletagger. Modul for hjelpefunksjoner
' Versjon: 0.35 Dato: 2024-01-19 Jostein Amlien: Feilretting i utskift av tagger på egenskaper. Definert standard overskrift på kodelister.
' Versjon: 0.34 Dato: 2024-01-18 Jostein Amlien: Globale styreparamtre, åpn diagrammer, realiserigner til kodeliste, roller sortert på sequenceNo, refaktorering med vekt på tabeller og tagger
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
''  ----------------------------------------------------------------------------


''  ----------------------------------------------------------------------------
'
' Project Browser Script main function
Sub OnProjectBrowserScript

	Dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()

	Select Case treeSelectedType

		Case otPackage
			Repository.EnsureOutputVisible "Script"
			Repository.ClearOutput "Script"
			
			' Code for when a package is selected

			Dim valgtPakke As EA.Package
			set valgtPakke = Repository.GetTreeSelectedObject()

			InitierGlobaleParametre
			
			GlobaleParametreForValgtPakke( valgtPakke)
			
			Session.Output "// Modellrapport "   + valgtPakke.element.name
			Session.Output "// Start of UML-model"		

			set utfil = tomTekstfil( filnavn + ".adoc")
			Call SkrivModellrapport( toppnivaa, valgtPakke)  '' full modellrapport
			utfil.close

			Session.Output "// End of UML-model"

			visSosiFormatRealisering = true	
			Call listRealiserteObjekttyper( valgtPakke)  '' flat beskrivelse av alle realiserte objekttyper

			Session.Output "// Dokumentasjon ferdig "
		Case Else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK

	End Select

End Sub


''  --------	Globale styringsparametre	------------------------------------

dim filnavn  '' navn på fil med modellrapporten

'' 	Styring av innhold
dim debugModell 
dim ignorerSosiformatTagger

dim genererDiagrammer : genererDiagrammer = true

dim standardTabellFormat
dim alleTaggerISammeTabellrad : alleTaggerISammeTabellrad = false
dim alternativBetegnelseForInitialverdi

dim detaljnivaa, nedersteOverskiftsnivaa
dim toppnivaa, oversteOverskiftsnivaa

dim visSosiFormatRealisering '' viser sosi-format i vedlegget 


''  ----------------------------------------------------------------------------

sub InitierGlobaleParametre
	''	parametre bør helst hentes fra en konfigurasjonsfil eller fra modellen
		
	filnavn = "modellRapport"

	''	--	Definer hvor mye info som skal skrives ut i rapporten
''	debugModell = true 	'' skriver ut noe innhold som kan avsløre modellfeil 
''	debugModell = false  '' filterer bort noe innhold som ikke skal rapporteres

'	ignorerSosiformatTagger = false '' Ta med tagger for SOSI-format i rapporten
	ignorerSosiformatTagger = true 	'' Utelat tagger for SOSI-format i rapporten
	
	genererDiagrammer = true  	'' regenererer alle diagrammer
''	genererDiagrammer = false  	'' anta at alle diagrammer er på plass
	
	''  ---------  	Utseende av tabeller	------------------------------------
	
	standardTabellFormat = "20,80"
	
	''	tabellhode for koder
''	alternativBetegnelseForInitialverdi = "Kodeverdi:"   

'	alleTaggerISammeTabellrad = true   '' alle taggene samla i en tabellrad
	alleTaggerISammeTabellrad = false  '' en tabellrad for hver tag
	
	
	''	-----	Styring av overskriftsnivåer   ---------------------------------
	toppnivaa = 2   
	detaljnivaa = 5	'' mest detaljerte overskriftsnivvå
	nedersteOverskiftsnivaa = 5
	oversteOverskiftsnivaa = 1
	
'		dim topplevel
'''		topplevel = 1    '''  foreslått endring
'		topplevel = 2
end sub


''  --------------   Globale variable avleda fra valgt pakke   -----------------

DIM rootId
DIM prefiksBokmerke

dim utkatalog	'' full path til hovedkatalogen for det genererte dokumentet 
dim imgfolder 	'' underkatalog med bildefiler for diagrammer

dim utfil

Dim projectclass As EA.Project 


''  ----------------------------------------------------------------------------

sub GlobaleParametreForValgtPakke( rotPakke) 
	''	--	Ta vare på hvilken pakke som blir dokumentert
	rootId = rotPakke.PackageId
	
	'' La navneromsforkortelsen til skjemaet styre prefiks
	dim xmlns
	xmlns = taggedValueFraElement(rotPakke.Element,"xmlns")

	prefiksBokmerke = xmlns
	
	dim sosiKortnavn
	sosiKortnavn = taggedValueFraElement(rotPakke.Element,"SOSI_kortnavn")

	dim pakkenavn
	pakkenavn = rotPakke.Element.Name
	
''	genererDiagrammer = false

	Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
	utkatalog = FSO.GetParentFolderName(Repository.ConnectionString()) & "\"
	if sosiKortnavn <> "" then 
		utkatalog = utkatalog & sosiKortnavn & "\"
		imgfolder = "Diagrammer\"
	elseif pakkenavn <> "" then
		utkatalog = utkatalog & pakkenavn & "\"
		imgfolder = "Diagrammer\"	
	else
		imgfolder = "Diagrammer" & xmlns & "\"	
	end if
	
	if not FSO.FolderExists(utkatalog) then FSO.CreateFolder utkatalog
	
	''	opprett katalog for bilder og diagrammer
	Dim imgparent : imgparent = utkatalog  & imgfolder
	if not FSO.FolderExists(imgparent) then FSO.CreateFolder imgparent
	
	set projectclass = Repository.GetProjectInterface()
end sub


''	============================================================================
'
'							MODUL Modellrapport
'
''	============================================================================

Sub SkrivModellrapport(pakkelevel, thePackage)


	dim pakkeElement
	set pakkeElement = thePackage.Element
	
'	dim xmlns : xmlns = taggedValueFraElement(pakkeElement,"xmlns")
'	if prefiksBokmerke = "" and xmlns <> "" then prefiksBokmerke = xmlns

'----------------- Overskrift og beskrivelse -----------------

	dim overskrift, pakketype

If false then   ''''''''' Erstatta av linjene under
	if pakkeElement.Stereotype = "" then

		SettInnTekst sideskift()              ''' Hvorfor det ?
		SettInnTekst skillelinje()

		if pakkelevel >= 4 then
			pakketype = "Underpakke: " 
		else
			pakketype = "Pakke: " 
		end if
	else
		pakketype = "Pakke: " 
	end if

	overskrift = pakketype + stereotypeNavn( pakkeElement)
else	 ''''''''''''  erstatter linjene over
	
''	SettInnTekst skillelinje()
	if  thePackage.packageId <> rootId then SettInnTekst sideskift()

	if pakkeElement.Stereotype <> "" then
		overskrift = stereotypeNavn( pakkeElement)   
	else
		overskrift = "Pakke: " + pathTilInternPakke(thePackage.PackageID)
	end if

end if

	SettInnTekst nummerertOverskrift( pakkelevel, overskrift ) 
	
	SettInnTekst bold("Definisjon:") & " " & definisjon(pakkeElement)

''  ----------------------------------------------------------------------------

	dim tittel 
	tittel = "Profilparametre i tagged values"
	tittel = unummerertOverskrift(pakkelevel +1, tittel)
	call SettInnSomTabell( pakkeTagger( pakkeElement), "", tittel)

	tittel = unummerertOverskrift(pakkelevel +1, "Avhengigheter")
	call SettInnSomTabell( pakkeavhengigheter( pakkeElement), "", tittel)
	
'----------------- Bilder og diagram-----------------

	SettInnBilde bildeAvPakke(pakkeElement)

	
	Dim diag As EA.Diagram
	For Each diag In thePackage.Diagrams
			
		SettInnDiagram(diag)
	
	next

'-----------------Elementer----------------- 

	Dim element As EA.Element 
	For each element in thePackage.Elements
	
		call beskrivElement(pakkelevel+1, element, thePackage)		

	Next

'----------------- Underpakker ----------------- 

	dim nesteLevel

'	ALT 1 Underpakker flatt på samme nivå som Application Schema
'	nesteLevel = pakkelevel

'	ALT 2 Nøsting av pakker ned til nivå 4 under Application Schema
	nestelevel = pakkelevel + 1
	if pakkelevel = 4 then nestelevel = 4

'	ALT 3 TBD Nøsting helt ned 
'	med utskrift av Pakke::Klasse (Pakke/Pakke2::Klasse TBD)
' 	nestelevel = pakkelevel + 1

'	ALT 4  Alle undeliggende pakker skrives ut på nivå 3
'	nestelevel = 3


	dim pack as EA.Package
	for each pack in thePackage.Packages
	
		Call SkrivModellrapport(nestelevel, pack)
		
	next

end sub

''  ----------------------------------------------------------------------------

function pakkeavhengigheter( pakkeelement )

	dim pakker, realisertFra, avhengigAv
	pakker = elementAvhengigAv( pakkeelement, "Dependency")
	if isArray(pakker) then 
		pakker = genererPakkeNavnListe(pakker)
		avhengigAv = array("Avhengig av", pakker)
	end if

	pakker = elementAvhengigAv( pakkeelement, "Realisation")
	if isArray(pakker) then 
		pakker = genererPakkeNavnListe(pakker)
		realisertFra = array("Realisert fra", pakker)
	end if
	
	pakkeavhengigheter = array( realisertFra, avhengigAv)
	
End function   

''  ----------------------------------------------------------------------------

sub BeskrivElement(elementLevel, element, pakke)

	if not (isFeatureOrDataType(element) or isCodelist(element) ) then EXIT sub
	
	SettInnTekst elementOverskrift(elementLevel, element, pakke) 

	SettInnTekst bold("Definisjon:") & " " & definisjon(element)

	dim undernivaa : undernivaa = elementLevel +1

	dim underTittel
	underTittel = "Profilparametre i tagged values"
	underTittel = unummerertOverskrift( undernivaa, underTittel)
	call SettInnSomTabell( elementTagger( element), "", underTittel)

	If isFeatureOrDataType(element)  Then	
		SettInnBilde bildeAvObjekttype(element)

		underTittel = unummerertOverskrift( underNivaa, "Egenskaper")
		call Egenskaper(element, underTittel)

		underTittel = unummerertOverskrift( underNivaa, "Roller") 
		call Relasjoner( element, underTittel)
		
		underTittel = unummerertOverskrift( underNivaa, "Operasjoner")
		call Operasjoner( element, underTittel)  ''
		
		underTittel = unummerertOverskrift( underNivaa, "Restriksjoner")
		call Restriksjoner( element, underTittel)  ''

	Elseif isCodelist(element) Then
		SettInnBilde bildeAvKodeliste(element)
		
		underTittel = unummerertOverskrift( undernivaa, "Koder")
		Call Kodeliste( element, underTittel)
 
	End if

	underTittel = unummerertOverskrift( undernivaa, "Arv og realiseringer")
	call SettInnSomTabell( arvOgRealiseringer( element), "", underTittel)

end sub

''  ----------------------------------------------------------------------------

'-----------------ObjektOgDatatyper med ArvOgRealisering -----------------

Sub Egenskaper(element, tittel)

	if element.Attributes.Count = 0 then EXIT SUB

	SettInnTekst tittel
	dim tabellFormat : tabellFormat = "20,80"	
	Dim att As EA.Attribute
	for each att in element.Attributes
	
		call SettInnSomTabell(  attributtbeskrivelse(att), tabellFormat, "")	

	next
	
end sub 

''  ----------------------------------------------------------------------------

function arvOgRealiseringer( element)
	dim elementListe 
	dim supertype, subtyper, realisertFra
	

	elementListe = elementAvhengigAv( element, "Generalization")
	if isArray(elementListe) then 
		elementListe = genererInternPathListe(elementListe)
		supertype = array("Supertype:", elementListe)
	end if

	elementListe = elementGirForingerFor( element, "Generalization")
	if isArray(elementListe) then  
		elementListe = genererElementNavnListe(elementListe)  '
		subtyper = array("Subtyper:", elementListe)
	end if

	elementListe = elementAvhengigAv( element, "Realisation")
	if isArray(elementListe) then   
		elementListe = genererPakkeNavnListe(elementListe)
		realisertFra = array("Realisert fra:", elementListe)
	end if

	arvOgRealiseringer = array( supertype, subtyper, realisertFra)
	
end function

'-----------------ObjektOgDatatyper / ArvOgRealiseringer   End-----------------


''  ----------------------------------------------------------------------------
' 					Kodelister  		
''  ----------------------------------------------------------------------------

Sub Kodeliste(element, tittel)

	if element.Attributes.Count > 0 then  

		dim koder : koder = modellkoder(element)

		dim hode : hode = koder(0)
		dim tabellFormat : tabellFormat = "20,80"	
		if UBound(hode) = 2 then tabellFormat = "30,60,10"
			
		call SettInnSomTabell( koder, tabellFormat, tittel) 
	end if

end sub

''  ----------------------------------------------------------------------------

function modellkoder(element)

	dim hode : hode = kodeTabellHode(element)
		
	dim treKolonner : treKolonner = ( Ubound(hode) = 2)
	
	dim liste() : redim liste(element.Attributes.Count)
	liste(0) = hode

	dim i :	i = 1
	
	Dim att As EA.Attribute	
	dim def, kode
	for each att in element.Attributes
		
		def = definisjon(att)
		def = def + bildeAvAttributt(att, "kodelistekode")

		if treKolonner then 
			kode = array( att.Name, def, att.Default) 
		else
			kode = array( att.Name, def )   
		end if
		
		liste(i) = kode
		i = i +1
	next
	
	modellkoder = liste
	
End function

''  ----------------------------------------------------------------------------

function kodeTabellHode(element)

	dim hode : hode = array( "Kodenavn:", "Definisjon:")

	dim att
	for each att in element.Attributes
		if att.Default <> "" then 
			hode = array( "Kodenavn:", "Definisjon:", "Utvekslingsalias:")
			if alternativBetegnelseForInitialverdi <> "" then 
				hode(2) = alternativBetegnelseForInitialverdi
			end if
			
			exit for
		end if
	next
	
	dim i
	for i = 0 to UBound(hode)
		hode(i) = bold(hode(i))
	next

	kodeTabellHode = hode

End function

'-----------------CodeList End-----------------

''  ----------------------------------------------------------------------------

function attributtbeskrivelse( att)

	dim navn, def, mult, init, visib, typ
	
	navn = array(bold("Navn:"), bold(att.Name) ) 

	def = array("Definisjon:", definisjon(att) )

	mult = array("Multiplisitet:", bounds(att))

	if att.Default <> "" then
		init = array("Initialverdi:", att.Default )
	end if

	if not att.Visibility = "Public" then
		visib = array( "Visibilitet:", att.Visibility )
	end if
		
	typ = array( "Type:", attributtype(att)	) 
	
	dim attributt
	attributt = array ( navn, def, mult, init, visib, typ)
	
	attributtbeskrivelse = merge(attributt, egenskapsTagger( att))

end function


''  ----------------------------------------------------------------------------
''				Operasjoner og restriksjoner
''  ----------------------------------------------------------------------------

sub Operasjoner(element, tittel)

	Dim meth as EA.Method
	dim tabell

	For Each meth In element.Methods
		call SettInnSomTabell( beskrivOperasjon( meth), "", tittel)
		tittel = ""
	next

end sub

''  ----------------------------------------------------------------------------

function beskrivOperasjon( meth)

	dim navn : navn = array( bold("Navn:"), bold( meth.Name) ) 
	dim beskrivelse : beskrivelse = array("Beskrivelse:", definisjon( meth) )
	
	beskrivOperasjon = array( navn, beskrivelse)

exit function

	dim ster, ret, behav
	
	ster = array( "Stereotype:", meth.Stereotype) 
	ret =  array( "Retur type:", meth.ReturnType) 
	behav = array( "Oppførsel:", meth.Behaviour) 
	
	beskrivOperasjon = array( navn, beskrivelse, ster, ret, behav)

end function

''  ----------------------------------------------------------------------------

sub Restriksjoner( element, underOverskrift)

	Dim constr as EA.Constraint
	For Each constr In element.Constraints

		call SettInnSomTabell( restrik( constr), "20,80", underOverskrift)
		underOverskrift = ""
	next
	
end sub

''  ----------------------------------------------------------------------------

function restrik( constr)

	dim navn, beskrivelse, typ, oclKode
	
	
	dim ocl : ocl = (constr.Type = "OCL")
	
	dim oclConstraint  '' , posInv
	if ocl then 
		dim noter : noter = split(constr.Notes, "inv:")
		beskrivelse = noter(0)
		if UBound(noter) > 0 then 
			oclConstraint = Trimutf8("inv:" + noter(1))
			oclKode = bokstavlig(oclConstraint)
		end if
	else
		beskrivelse = constr.Notes
	end if
	
	navn = array( bold("Navn:"), bold( trim( constr.Name)) ) 
	typ = array( "Type:", constr.Type)
	beskrivelse = array("Beskrivelse:", getCleanRestriction(beskrivelse) ) 
	if not isEmpty(oclKode) then oclKode = array("OCL kode:", oclKode)

	restrik = array( navn, beskrivelse, typ, oclKode)

exit function
	dim status : status = array( "Status:", constr.Status)
	dim vekt : vekt= array( "Vekt:", constr.Weight)
	restrik = array( navn, beskrivelse, typ, oclKode, status, vekt)

end function


'	-----------------	Operasjoner og Restriksjoner 	End	--------------------


'===============================================================================
'
'		MODUL for relasjoner og Assosiasjonsroller   
'
'===============================================================================
'

''  ----------------------------------------------------------------------------
'					Relasjoner		
''  ----------------------------------------------------------------------------
'
sub Relasjoner( element, byval underOverskrift)
''	finn gjerne et nut navn til denne rutina
''
'assosiasjoner
' skriv ut roller - sortert etter tagged value sequenceNumber

	dim rollesamling  	'' array av rolle-arryer
	rollesamling = sorterteRoller( element)
	if isEmpty(rollesamling) then 	EXIT sub

	SettInnTekst underOverskrift
	dim rolle 
	for each rolle in rollesamling
''		call SettInnSomTabell(  beskrivRolle( rolle), "25,75", "")
		call SettInnSomTabell(  beskrivRolle( rolle), "", "")
	next
	

end sub

''  ----------------------------------------------------------------------------

function beskrivRolle(rolle)

	dim seqNo : seqNo = rolle(0)

	dim con : set con = rolle(1)
	dim current : set current = rolle(2)
	dim target : set target = rolle(3)
	dim targetID : targetID = rolle(4)
	
	dim aggType, res
	aggType = aggregationType(current)
	
	dim rol, roltag, konn, konntag
	
	rol = rolleBeskrivelse( target, targetID, aggtype)
	roltag = rolleTagger(target) 	
	rol = merge( rol, roltag)
	
	konn = merge (konnektor(con), konnektorTagger(con) )


''	beskrivRolle = array( rol, roltag, konn, konntag)

	beskrivRolle = merge( rol, konn)
end function

''  ----------------------------------------------------------------------------

function alleRoller( element)
	dim rolle   		'' array
	dim rollesamling  	'' array av rolle-arryer
	
	dim beggeEnder : beggeEnder = array( "Supplier", "Client")
	Dim con
	For Each con In element.Connectors
		dim rolleEnde   
		for each rolleEnde in beggeEnder  
			rolle = identifiserRolle( element.elementID, con, rolleEnde)
			if isArray(rolle) then 
				rollesamling = merge( rollesamling, array(rolle))
''				rollesamling = append( rollesamling, rolle)
			end if
		next   
	next

	alleRoller = rollesamling
end function

''  ----------------------------------------------------------------------------

function sorterteRoller( element)
	dim rolle   		'' array
	dim rollesamling  	'' array av rolle-arryer
	
	rollesamling = alleRoller( element)

	if isEmpty(rollesamling) then 	EXIT function

	dim sekvens()			'' array av sekvensnummre
	redim sekvens(UBound(rollesamling))
	for i = 0 to UBound(rollesamling)
		sekvens(i) = rollesamling(i)(0)
	next

	dim indeks
	indeks = sortertIndeks(sekvens)

	dim res
	res = rolleSamling
	dim i, j
''	for each i in indeks
	for j = 0 to UBound(indeks) 
		res(j) = rollesamling(indeks(j))
	next

	sorterteRoller = res
end function

''  ----------------------------------------------------------------------------

function identifiserRolle( elemID, con, clientEllerSupplier)
	'' Returnerer en referanse til en rolle som skal beskrives, 
	'' dvs. en connector av typen "Association" eller "Aggregation"
	
	if con.Type <> "Association" and con.Type <> "Aggregation" Then
		EXIT function		
	end if

''  Det er hensiktsmesig å rapportere aggregreringene i modellen 
	''  på samme måte som EA velger å framstille dem i diagrammene
	call fiksKonnektor( con)
	
	dim targetID, target, current
	if elemID = con.SupplierID and clientEllerSupplier="Client" then
		targetID = con.ClientID
		set target = con.ClientEnd	
		set current = con.SupplierEnd	
		
	elseif elemID = con.ClientID and clientEllerSupplier="Supplier" then
		targetID = con.SupplierID
		set target = con.SupplierEnd
		set current = con.ClientEnd
	
	else 
		exit function
	end if
	
	if target.Role <> "" then   ''erstatter realiserbarRolle( target, aggType)
		dim seqNo
		seqNo = taggedValueFraRolle(target, "sequenceNumber")

		if seqNo <> "" then seqNo = Cint(seqNo)		

		identifiserRolle = array( seqNo, con, current, target, targetID)
	end if

end function

''  ----------------------------------------------------------------------------


''  UTGÅR  ########
function realiserbarRolle( r, aggType)

	dim res
	res =  r.Role <> "" or r.RoleNote <> "" or r.Cardinality <> ""
	res =  res or r.Navigable = "Navigable" or aggType <> "Assosiasjon" 

	realiserbarRolle = res
end function


''  ----------------------------------------------------------------------------

function rolleBeskrivelse( targetEnd, targetID, aggregeringsType)
''
''	targetEnd angir Rollen: navn, def, multiplisitet og navigerbarhet
''	targetID angir klassen som Rollen peker på
''	aggregeringsType 

	dim res 
	dim rolle
''	redim rolle(4)
	dim navn, definisjon, multiplisitet, assType, targetRef
	
	navn = targetEnd.Role
	if navn = "" and debugModell then navn = "ADVARSEL: ROLLENAVN MANGLER"
	navn = array( bold("Rollenavn:"), bold(navn) ) 
	res = navn
	
	definisjon = getCleanDefinition(targetEnd.RoleNote)
	If definisjon = "" and debugModell Then	
		definisjon = bold("ADVARSEL: DEFINISJON MANGLER") 	
	end if
	definisjon = array( "Definisjon:", definisjon )
	res = merge( res,  definisjon  )


	multiplisitet = targetEnd.Cardinality  '''' tekstformat
	if multiplisitet <> "" then 
		multiplisitet = "[" + targetEnd.Cardinality + "]"
		multiplisitet = array( "Multiplisitet:", multiplisitet )
	else
		multiplisitet = array(  "Multiplisitet:", " " ) 
	end if
	res = merge( res,  multiplisitet )

''	res = merge( res, array(  "Aggregeringstype:", aggregeringsType) )     
	assType = array(  "Assosiasjonstype:", aggregeringsType)
	res = merge( res, assType )    
	
	dim textVar 		: textVar = "Til klasse"
	dim navigerbarhet 	: navigerbarhet = targetEnd.Navigable
	'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
	If navigerbarhet = "Navigable" Then 
		textVar = textVar + kursiv(":") 
''		textVar = textVar + kursiv(" (navigerbar):") 
''		textVar = "Navigerbar til:"
	ElseIf navigerbarhet = "Non-Navigable" Then 
		textVar = textVar + kursiv(" (ikke navigerbar):") 
''		textVar = "Ikke navigerbar til:"		
'	Elseif navigerbarhet = "Unspecified" Then 
'		textVar = textVar + bold(" (Uspesifisert):") 
''		textVar = "Til klasse"
	Else 
		textVar = textVar + ":" 
''		textVar = "Til klasse:"
	End If
	
	dim targetElement : set targetElement = Repository.GetElementByID(targetID)
	
'''	dim pakkeReferanse 
'''	pakkeReferanse = pathTilEksternPakke( targetElement.PackageID)
'	dim pakkeReferanse 
'	pakkeReferanse = pathTilInternPakke( targetElement.PackageID)
'	dim elementReferanse 
'	elementReferanse = targetLink( targetElement )
'	targetRef = pakkeReferanse + "::" + elementReferanse

	targetRef = array(  textVar, pathTilInterntElement(targetElement) )
	res = merge( res, targetRef )

	rolle = array( navn, definisjon, multiplisitet, assType, targetRef)

	rolleBeskrivelse = rolle
end function

''  ----------------------------------------------------------------------------

function konnektor( connector)
	dim res 

	dim konnNavn : konnNavn = stereotypeNavn(connector)
''	
	if konnNavn <> "" then 
		res = array( "Konnektor: ", konnNavn)
		res = merge( res, array( "Konnektortype:", connector.Type )  )
	elseif connector.Type = "Aggregation" then
		'' Trenger vi denne ?
		res = array( "Konnektortype:", connector.Type ) 
	end if
		
	if not isEmpty(res) then konnektor = res

end function

''  ----------------------------------------------------------------------------

function currentEnd( ende )
	dim res 

	If ende.Role <> "" Then
		res = merge( res, array( "Fra rolle:", ende.Role ) )
	End If
	If ende.RoleNote <> "" Then
		dim def : def = getCleanDefinition(ende.RoleNote)
		res = merge( res, array( "Fra rolle definisjon:", def ) )
	End If
	If ende.Cardinality <> "" Then
		res = merge( res, array( "Fra multiplisitet:", ende.Cardinality ) )
	End If

	CurrentEnd = res
end function

''  ----------------------------------------------------------------------------

sub FiksKonnektor( connector)
	'' Denne funksjonen gjennomfører de tilpasningene som EA viser i diagrammene

	if connector.Type = "Aggregation" then
		'' En Aggregation kan ikke ha Assosiasjonsroller i begge ender
		if connector.ClientEnd.Aggregation > 0 then
		elseif connector.SupplierEnd.Aggregation = 0 then
			connector.SupplierEnd.Aggregation = 1
		end if
				
		'' En Aggregation kan ikke ha ensidig retning mot destinasjon 	
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
		SettInnTekst "SYSTEMFEIL:  Assosiasjon har aggregeringstype:" & aggtype
		exit function
	end if

	aggregationType = res	
end function


''  ----------------------------------------------------------------------------
''
''	Rutiner som brukes til å sortere på sekvensnummer
''
''  ----------------------------------------------------------------------------

function sortertIndeks(byval sekvens)
	dim res()
	redim res(UBound(sekvens)) ''  : res = order

	dim i, pos
	for i = 0 to UBound( sekvens)
		pos = posMinimum(sekvens)
		
		if isEmpty(pos) then pos = firstEmpty(sekvens)
		
		res(i) = pos
		sekvens(pos) = -1   '' marker elementet som identifisert
	next
	
	sortertIndeks = res
	
end function

''  ----------------------------------------------------------------------------

function firstEmpty(sekvens)
''  identifiserer posisjoen til det første tomme elementet i en sekvens
	dim j
	for j = 0 to UBound(sekvens) 
		if isEmpty(sekvens(j)) then 
			firstEmpty = j
			exit function
		end if
	next
end function

''  ----------------------------------------------------------------------------

function posMinimum(sekvens)
''  identifiserer posisjoen til det minste positive tallet i en sekvens
	dim min, pos, j, seqNo

	for j = 0 to UBound(sekvens)
		seqNo = sekvens(j)
		if not isNull(seqNo) and seqNo > 0 then    '' sjekk denne kandidaten
			if isEmpty(min) or seqNo < min  then  '' beste kandidat så langt
				min = seqNo
				pos = j
			end if
		end if
	next

	posMinimum = pos
	
end function

 
 

'	============================================================================
'
'					MODUL   navigeringEA    
'
'	============================================================================
'
'--------   Funksjoner som forholder seg til UML-modelllen i EA  ---------------

''  ----------------------------------------------------------------------------
'				Elementtyper
''  ----------------------------------------------------------------------------

function isFeatureOrDataType(element)

	isFeatureOrDataType = isFeatureType(element) or isDataType(element)

end function

''  ----------------------------------------------------------------------------

function isFeatureType(element)
	dim ster : ster = Ucase(element.Stereotype)
	
	isFeatureType = (ster = "FEATURETYPE")
end function

''  ----------------------------------------------------------------------------

function isDataType(element)
	dim ster : ster = Ucase(element.Stereotype)

	isDataType = (ster = "DATATYPE" OR ster = "UNION" )
end function

''  ----------------------------------------------------------------------------

function isCodelist(element)
	dim ster : ster = Ucase(element.Stereotype)
	dim sterOK
	sterOK = (ster = "CODELIST" OR ster = "ENUMERATION" )
	
	isCodelist = sterOK OR (element.Type = "Enumeration")
end function

''  ----------------------------------------------------------------------------

function isAbstract(element)
	dim ster : ster = Ucase(element.Stereotype)

	isAbstract = (element.Abstract = 1 and ster = "FEATURETYPE" )
end function

''  ----------------------------------------------------------------------------

function erSosiBasistype( attrType)
	'' trenger å sjekke om typen faktisk er en SOSI basistype
	
	''**  En mer generesk tilnærming ville være å traversere nettressusen under
	''	  og sett opp lista i bTyp ihht. typene som er definert der.
	dim path : path = "http://skjema.geonorge.no/SOSI/basistype/" 
	
	dim bTyp
	bTyp = "Date, Time, DateTime, Number, Decimal, Integer, Real, Vector" 
	bTyp = bTyp + ", CharacterString, Boolean, URI, Any, Record, LanguageString"
	bTyp = bTyp + ", GM_Point, GM_Curve, GM_Surface, GM_Solid"
	bTyp = bTyp + ", GM_Primitive, GM_MultiSurface" 
	bTyp = bTyp + ", Punkt, Sverm, Kurve, Flate" 

	dim basisTyper : basisTyper = Split(bTyp, ", ")

	erSosiBasistype = listeInneholderElement( basisTyper, attrType)  

end function

''  ----------------------------------------------------------------------------

function erLovligBasisType( attrType) 
	dim lovligetyper
''	lovligetyper = Split("string, integer, date, boolean", ", ") 
	lovligetyper = "string, integer, date, boolean"  '' typer for xml
	lovligetyper = Split(lovligetyper, ", ") 
	
	erLovligBasisType = listeInneholderElement( lovligetyper, attrType)

end function

''  ----------------------------------------------------------------------------

function listeInneholderElement( liste, element)
	listeInneholderElement = false
	
	dim el
	for each el in liste
		if el = element then listeInneholderElement = true
	next
end function

''  ----------------------------------------------------------------------------

function attributtype(att)	

	dim sosiPath : sosiPath = "http://skjema.geonorge.no/SOSI/basistype/" 
	if erSosiBasistype(att.Type) then
		attributtype = eksternLenke( sosiPath & att.Type, att.Type)

	elseif erLovligBasisType( att.Type) then   '' lovlig ihht f.eks. xml
		attributtype = att.Type

	elseif att.ClassifierID > 0 then  '' referanse til en klasse i modellen

		''  Attributtets datatype er referert fra att.ClassifierID 
		''  Må forutsette at ClassifierID faktisk peker på et element
		dim classifier as EA.Element   '' for å angi attributtets datatype
		set classifier = Repository.GetElementByID(att.ClassifierID) 

		'' Sjekk om classifier skulle være en datatype definert utafor scope
		dim typ
		if erEksternPakke(classifier.PackageID) then
			'' classifier er ekstern og derfor ikke beskrevet i dette dokumentet
			typ = pathTilEksternPakke(classifier.PackageID) 
			typ = typ + "::" + understrek(stereotypeNavn( classifier))
		else
			typ = targetLink( classifier)
		end if
		
		attributtype = typ		
	else
		attributtype = bold("Ukjent type: ") + att.Type	
	end if

end function 

''  ----------------------------------------------------------------------------

function elementGirForingerFor( element, conType)
	''  returnerer de elementene er avhengig av dette elementet
	''	i form av en array med IDer

	dim liste 

	DIM pakkeReferanse, elementReferanse  '' , targetReferanse
	dim con 
	for each con in element.Connectors 
		if element.ElementID = con.SupplierID and con.Type = contype then
			liste = merge(liste, con.ClientID)
		end if
	next
	elementGirForingerFor = liste
end function

''  ----------------------------------------------------------------------------

function elementAvhengigAv( element, conType)
	''  returnerer de elementene som dette elementet er avhengig av
	''	i form av en array med IDer
	dim liste 

	dim con  
	for each con in element.Connectors 
		if element.ElementID = con.ClientID and con.Type = contype then
			liste = merge(liste, con.SupplierID)
		end if
	next
	elementAvhengigAv = liste
end function


''  ----------------------------------------------------------------------------
'				Referanser til pakker i modellregisteret
''  ----------------------------------------------------------------------------

function genererPakkeNavnListe(IDliste)
	dim liste
	
	DIM pakkeReferanse, elementReferanse, targetReferanse
	dim id, target
	for each id in IDliste
		set target = Repository.GetElementByID(id)

		if target.Type <> "Boundary" then 					''''''''''''''''''
		pakkeReferanse = pathTilEksternPakke(target.PackageID) 
		elementReferanse = understrek(stereotypeNavn(target))
		targetReferanse = pakkeReferanse + "::" + elementReferanse	

		liste = merge(liste, targetReferanse)
		end if											'''''''''''''''''''''''
	next
	genererPakkeNavnListe = liste
end function

''  ----------------------------------------------------------------------------

function genererElementNavnListe(IDliste)
	dim liste
	
	dim id, target
	for each id in IDliste
		set target = Repository.GetElementByID(id)
		liste = merge(liste, targetlink(target))
	next
	genererElementNavnListe = liste
end function

''  ----------------------------------------------------------------------------

function genererInternPathListe(IDliste)
	dim liste
	
	dim id, target
	for each id in IDliste
		set target = Repository.GetElementByID(id)
		liste = merge(liste, pathTilInterntElement(target))
	next
	genererInternPathListe = liste
end function

''  ----------------------------------------------------------------------------

function erEksternPakke( pakkeID)  '' pakke i et annet skjama
''	dim rootId  '' global

	dim pakke
	set pakke = Repository.GetPackageByID(pakkeID)

	dim  res
	if pakkeID = rootId then  
		'' Vi har nådd toppen av denne modellen: pakka er lokal
		res = false
	elseif pakke.parentID = rootId then  
		'' Vi har nådd toppen av denne modellen: pakka er lokal
		res = false
	elseif pakke.name = "SOSI Model" then  
		'' Vi har nådd toppen av modellregisteret
		res = true
	elseif pakke.parentID = 0 then	
		'' Vi har nådd toppen av modellregisteret: pakka er ekstern
		res = true
	elseif pakke.Element.Stereotype <> "" then 
		'' Vi har nådd et annet applikasjonskjema: pakka er ekstern
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

	if pakke.parentID = 0 then	
		'' Vi har nådd toppen av modellregisteret: pakka er ekstern
		res = ""
	elseif pakkenavn = "SOSI Model" then  
		'' Vi har nådd toppen av modellregisteret
		res = ""
	elseif pakke.Element.Stereotype <> "" then 
		'' Vi har nådd et applikasjonskjema: pakka er ekstern
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

function pathTilInterntElement( element)

	dim tlink : tlink = targetLink(element)
	pathTilInterntElement = pathTilInternPakke(element.PackageID) + "::" + tLink

end function

''  ----------------------------------------------------------------------------

function pathTilInternPakke( pakkeID)
''  Denne brukes for å sette overskrift på pakkene
''	dim rootId  '' global

	dim pakke
	set pakke = Repository.GetPackageByID(pakkeID)

	dim pakkenavn, res
	pakkenavn = pakke.name

	if pakke.parentID = rootId then  
		'' Vi har nådd toppen av denne modellen: pakka er lokal
		res = pakkenavn
	elseif pakke.parentID = 0 then	
		'' Vi har nådd toppen av modellregisteret: pakka er ekstern
		res = ""
	elseif pakke.Element.Stereotype <> "" then 
		'' Vi har nådd et annet applikasjonskjema: pakka er ekstern
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
'
'					MODUL: taggedValues
'
'		Høsting av tagged values fra modellen  
'
''	============================================================================

'----------------  Funksjoner for å lese tagged values -------------------

function taggedValueFraElement(element, byVal tagName)
	tagName = LCase(tagName)
	taggedValueFraElement = ""

	dim tag
	for each tag in element.TaggedValues
		if LCase(tag.Name) = tagName and tag.Value <> "" then
			taggedValueFraElement = tag.Value
			exit for
		end if
	next

end function

''  ----------------------------------------------------------------------------

function taggedValueFraRolle(rolle, byVal tagName)
	tagName = LCase(tagName)

	dim tag
	for each tag in rolle.TaggedValues
		if LCase(tag.Tag) = tagName and tag.Value <> "" then
			taggedValueFraRolle = tag.Value
			exit for
		end if
	next

end function

''  ----------------------------------------------------------------------------

function egenskapsTagger( element) '' element er et atributt

	dim res()
	redim res(element.TaggedValues.Count )
	dim tagNr : tagNr = 0
	
	dim tag
	for each tag in element.TaggedValues
		if not ignorerTag(tag.Name) and tag.Value <> "" then 	
			res(tagNr) = array(tag.Name, tag.Value) 
			tagNr = tagNr + 1
		end if
	next

	if tagNr = 0 then exit function

	redim preserve res(tagNr-1)

	egenskapsTagger = res

	if alleTaggerISammeTabellrad then 
		'' Legg alle tagger i en og samme tabell-celle
		dim taggListe : taggListe = listeFraTabell(res, ": ")

		if  isEmpty( taggListe) then EXIT function 
		
	''	taggListe = array( "Profilparametre i tagged values:", taggListe )
		taggListe = array( "Tagged values:", taggListe )
		egenskapsTagger = array(taggListe)

	end if

end function

''  ----------------------------------------------------------------------------

function elementTagger( element) 

	dim antallTagger 
	antallTagger = element.TaggedValues.Count
	if antallTagger = 0 then 	EXIT function
	
	dim tagger()
	redim tagger(antallTagger )
	dim tagNr : tagNr = 0
	
	dim tag
	for each tag in element.TaggedValues
		if not ignorerTag(tag.Name) and tag.Value <> "" then 	
			tagger(tagNr) = array(tag.Name, tag.Value) 
			tagNr = tagNr + 1
		end if
	next
	
	if tagNr = 0 then exit function

	redim preserve tagger(tagNr-1)
	elementTagger = tagger

end function

''  ----------------------------------------------------------------------------

function pakkeTagger( element) '' element er ei pakke

	dim antallTagger 
	antallTagger = element.TaggedValues.Count
	if antallTagger = 0 then 	EXIT function
	
	dim tagger()
	redim tagger(antallTagger )
	dim tagNr : tagNr = 0
	
	dim tag
	for each tag in element.TaggedValues
		if not ignorerTag(tag.Name) and tag.Value <> "" then 	
			tagger(tagNr) = array(tag.Name, tag.Value) 
			tagNr = tagNr + 1
		end if
	next

	if tagNr = 0 then exit function
	
	redim preserve tagger(tagNr-1)
	pakkeTagger = tagger

end function

''  ----------------------------------------------------------------------------

function konnektorTagger( con) 

	dim antallTagger 
	antallTagger = con.TaggedValues.Count
	if antallTagger = 0 then 	EXIT function

	dim tagger()
	redim tagger(antallTagger )
	dim tagNr : tagNr = 0
	
	dim res, tag
	for each tag in con.TaggedValues
		if not ignorerTag(tag.Name) and tag.Value <> "" then 	
			tagger(tagNr) = array(tag.Name, tag.Value) 
			tagNr = tagNr + 1
		end if
	next
	
	redim preserve tagger(tagNr-1)

	konnektorTagger = tagger

end function

''  ----------------------------------------------------------------------------

function rolleTagger( rol) 

	dim antallTagger 
	antallTagger = rol.TaggedValues.Count
	if antallTagger = 0 then 	EXIT function
	
	dim tagger()
	redim tagger(antallTagger )
	dim tagNr : tagNr = 0

	dim res, tag
	for each tag in rol.TaggedValues
		if not ignorerTag(tag.Tag) and tag.Value <> "" then 	
			tagger(tagNr) = array(tag.Tag, tag.Value) 
			tagNr = tagNr + 1
		end if
	next

	redim preserve tagger(tagNr-1)

	rolleTagger = tagger

end function

'
'----------------  Funksjoner for å lese tagged values End -------------------

''  ----------------------------------------------------------------------------

function profilParametre( element) 
''	kalles ikke men testen kan brukes i rutina for pakketagger

	dim res
	dim tag
	for each tag in element.TaggedValues
		if ignorerTag(tag.Name) then 						''	pass
		elseif tag.Value <> "" then 						''	pass
		elseif tag.Name = "byValuePropertyType" then					''	pass
		elseif tag.Name = "isCollection" then						''	pass
		elseif tag.Name = "noPropertyType" then 						''	pass
		elseif tag.Name = "asDictionary" AND tag.Value = "false" then 	''	pass
		else
			res = merge ( res, array(tag.Name, tag.Value) )
		end if
	next

	profilParametre = res

end function

''  ----------------------------------------------------------------------------

function ignorerTag( byval navn)
	navn = LCase(navn) 
	dim ignorer

	ignorer = navn = "persistence" or navn = "sosi_melding" '' skriv ut BARE i debug-modus
	ignorer = ignorer AND not debugModell
	ignorer = ignorer or navn = "sosi_bildeavmodellelement"  '' 	tas separat, hopper over
		
	ignorerTag = ignorer or ignorerSosiFormatTag(navn)

end function
''  ----------------------------------------------------------------------------

function ignorerSosiFormatTag(tagnavn)

	if ignorerSosiformatTagger then
		dim sosiTagger 
		sosiTagger = array( "sosi_navn","sosi_lengde", "sosi_datatype") 
		ignorerSosiFormatTag = listeInneholder( sosiTagger, tagnavn) 
		
	else
		ignorerSosiFormatTag = false
	end if 
	
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
''  Returnerer en formattert tekst med nedre og øvre grense for et intervall
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
''		tekstformatEnumeration = """Enumeration"" "
		tekstformatEnumeration = "«Enumeration»"
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
	targetLink = internLenke( bokmerke(element), stereotypeNavn(element) )

end function

''  ----------------------------------------------------------------------------

function elementOverskrift(elementLevel, element, pakke)
	
	dim elementnavn, tittel
	elementnavn = stereotypeNavn(element) 
	if isAbstract(element) then '''' NYTT: gjort abstracte klasser kursiv
		elementnavn = kursiv( elementnavn & " (abstrakt)" )   
	end if

	if elementLevel > 4 then   
		tittel = nummerertOverskrift( 4, pakke.Name & "::" & elementnavn)
	else
		tittel = nummerertOverskrift( elementLevel, elementnavn) 
	end if
	
	dim res : res = overkriftMedBokmerke( tittel, bokmerke(element))
	elementOverskrift = merge( skillelinje(), res)
	
end function


''  ----------------------------------------------------------------------------
''				Tabeller
''  ----------------------------------------------------------------------------

sub SettInnSomTabell( byVal data, tabellFormat, overskrift)

''		if isEmpty(data) then EXIT sub

	if not isArray(data) then EXIT sub

	if UBound(data) < 0 then EXIT sub 

	if tabellFormat = "" then tabellFormat = standardTabellFormat

	SettInnTekst formatertTabell( data, tabellFormat, overskrift)

end sub

''  ----------------------------------------------------------------------------

function formatertTabell( byVal tabell, tabellFormat, overskrift)

	if not isArray(tabell) then EXIT function
'
	dim antallRader : antallRader = UBound(tabell) +1
	if antallRader = 0 then EXIT function 

	if tabellFormat = "" then tabellFormat = standardTabellFormat

	dim res() : redim res(UBound(tabell)+2)
	
	dim i, rad
	res(0) = array( overskrift, tabellstart( tabellFormat) )
	i = 1
	for each rad in tabell 
		rad = tabellRad( rad)
		if not isEmpty(rad) then
			res(i) = rad
			i = i +1
		end if
	next
	res(i) = tabellavslutning()
	redim preserve res(i)
	
	if i > 1 then 
		formatertTabell = res
	end if

end function

''  ----------------------------------------------------------------------------

function tabellRad( byval rad)
	
	if isEmpty( rad) then 	EXIT function

	if not isArray(rad) then
		tabellRad = tabellCelle( rad) '' " - skulle vært array... ")
		EXIT function
	end if

	dim res()
	redim res(UBound(rad) +1)
	dim i
	for i = 0 to UBound(rad) 
		res(i) = tabellCelle( rad(i) )
	next
	res(UBound(rad) +1) = " "  '' ihht adoc-konvesjon: blank linje
	
	tabellRad = res
	
end function


'  ----------------------------------------------------------------------------
'  		Utskrift av generert tekst
'  			TBD:  Direkte utskrift til fil
'  ----------------------------------------------------------------------------
'
sub SettInnTekst( tekst)
	if not isArray(tekst) then 
		if tekst <> "" then SettInnTekstLinje tekst
	else
		dim t
		for each t in tekst
			SettInnTekst t   ''rekursivt
		next
	end if 
end sub

''  ----------------------------------------------------------------------------

sub SettInnTekstLinje( tekstlinje)

	if isNull(utfil) then 
		Session.Output tekstlinje
	else
		utfil.write tekstlinje & vbCrLf
	end if
	
end sub

''  ----------------------------------------------------------------------------

function tomTekstfil( byval filnavn)
	filnavn = utkatalog + filnavn
	
	Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")

	dim overWrite : overWrite = true
	dim unicode : unicode = true

	set tomTekstfil = FSO.CreateTextFile( filnavn, overWrite, unicode)
	
end function


'===============================================================================
'
'		MODUL for hjelpefunksjoner  
'
'===============================================================================


function merge( ByVal list, byVal tillegg)
''  Forlenger ei liste (array) med et tillegg, og returnere ei ny liste
''
	if isEmpty(tillegg) Then 
		merge = list
		exit function
	elseif not isArray(tillegg) Then 
		tillegg = array(tillegg)
	end if
	
	if isEmpty(list) Then 
		merge = tillegg
		exit function
	elseif not isArray(list) then
		list = array(list)
	end if


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


''  ----------------------------------------------------------------------------

function listeFraTabell( byval tabell, skilletegn)
	''  Input er en tabell som består av en array av arrayer, kalt rader
	''	Hver rad gjøres om fra en array til en sammensatt streng med skilletegn
	''	Resultatet er ei liste av sammensatte strenger
	
	if not isArray(tabell) then EXIT function

	if skilletegn = "" then skilletegn = ": "

	dim res()
	redim res( UBound(tabell) )

	dim i, rad, nyrad
	for each rad in tabell
		if isArray(rad) then 
			res(i) = join( rad , skilletegn )
			i = i+1 
		end if
	next
	
	redim preserve res(i-1) 
	listeFraTabell = res
		
end function

''  ----------------------------------------------------------------------------

function listeInneholder(liste, verdi)

	listeInneholder = false
	
	dim ledd
	if isArray(liste) then
		for each ledd in liste
			if ledd = verdi then listeInneholder = true
		next
	end if
	
end function

'===============================================================================
'
'		MODUL for Bilder   
'
'===============================================================================

function bildeAvObjekttype(element)
	dim standardTekst, alt
	
	standardTekst = "Illustrasjon av objekttype "  + element.Name
	if isDataType(element) then
		standardTekst = "Illustrasjon av datatype "  + element.Name
	end if
	
	alt = "Bilde av et eksempel på objekttypen " + element.Name

	bildeAvObjekttype = bildeAvModellElement( element, standardTekst, alt)
	
end function 

''  ----------------------------------------------------------------------------

function bildeAvKodeliste(element)
	dim standardTekst, alt

	standardTekst = "Illustrasjon av kodeliste " + element.Name

	alt = "Illustrasjon av hva kodelisten "  + element.Name + " kan inneholde" 

	bildeAvKodeliste = bildeAvModellElement( element, standardTekst, alt)
	
end function 

''  ----------------------------------------------------------------------------

function bildeAvPakke(element)
	dim standardTekst, alt

	standardTekst = "Illustrasjon av pakke "

	alt = "Illustrasjon av innholdet i UML-pakken " + element.Name
	
	bildeAvPakke = bildeAvModellElement( element, standardTekst, alt)
	
end function 

''  ----------------------------------------------------------------------------

function bildeAvModellElement( element, standardBildeTekst, standardAltTekst)
''	**  Bør få et nytt navn ======


	dim bilde
	bilde = taggedValueFraElement(element, "SOSI_bildeavmodellelement")  

	if bilde = "" then EXIT FUNCTION
	
	dim bildeTekst
	bildeTekst = taggedValueFraElement(element, "SOSI_bildetekst")
	if bildeTekst = "" then bildeTekst = standardBildeTekst
	
	dim altBildeTekst
	altBildeTekst = taggedValueFraElement(element, "SOSI_alternativbildetekst")
	if altBildeTekst = "" then altBildeTekst = standardAltTekst
	
	bildeAvModellElement = bildeFrittstaaende(bildeTekst, bilde, altBildeTekst) 

end function

''  ----------------------------------------------------------------------------

sub SettInnDiagram(diag)

	dim altBildeTekst
	if diag.Notes <> "" then
		altBildeTekst = getCleanDefinition(diag.Notes)  
	else
		dim altTekst(2)
		altTekst(0) = "Diagram med navn " 
		altTekst(1) = diag.Name 
		altTekst(2) = " som viser UML-klasser beskrevet i teksten nedenfor."
		altBildeTekst = join( altTekst)
	end if
	

	dim diagramfil : diagramfil = imgfolder + diag.Name + ".png"

	if genererDiagrammer then 
		dim openDia : openDia = ( Repository.IsTabOpen(diag.Name) > 0 )

		dim pathToDia : pathToDia = utkatalog & diagramfil
		Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, pathToDia, 1)

		if not openDia then call Repository.CloseDiagram(diag.DiagramID)
	end if

	SettInnBilde bildeFrittstaaende(diag.Name, diagramfil, altBildeTekst) 

end sub

''  ----------------------------------------------------------------------------

function bildeAvAttributt( att, typ)

	dim bilde
	bilde = taggedValueFraElement(att, "sosi_bildeavmodellelement")
	if bilde <> "" then
		dim bildetekst, alternativTekst
		bildetekst = "Illustrasjon av " & typ & " " & att.Name &""		
		alternativTekst = "Bilde av " & typ & " " & att.Name 
		alternativTekst = alternativTekst & " som er forklart i teksten."
		
		bildeAvAttributt = bildeITekst( bildetekst, bilde, alternativTekst)	
	end if

end function 

''  ----------------------------------------------------------------------------

sub SettInnBilde( bilde)

	if isArray(bilde)  then
		SettInnTekst Skillelinje()
	elseif bilde <> "" then
		SettInnTekst Skillelinje()
	end if

	SettInnTekst bilde

end sub



''	============================================================================
'					MODUL: tekstformatering
'		Hjelpefunksjoner for tekst som er uavhengig av adoc  
'
''	============================================================================

function definisjon( element)

	dim advarsel : advarsel = bold("ADVARSEL: DEFINISJON MANGLER")	
	dim def : def = getCleanDefinition( element.Notes)

	if def = "" and debugModell then 
		definisjon = advarsel		
	else
		definisjon = def
	end if
	
end function 

''  ----------------------------------------------------------------------------
'
'				Rensing av modelltekst som skal brukes i dokumentet
'
''  ----------------------------------------------------------------------------

function getCleanDefinition(byVal txt)     '' Gjenstår: fjerne avsnitt i teksten
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

function getCleanBildetekst(byVal txt)                         ''' Ikke i bruk

	dim res
	res = getCleanDefinition(txt)
	
	call ErstattTegn( res, ",", " ")
	
	getCleanBildetekst = res	

end function


''  ----------------------------------------------------------------------------
'							 Trimutf8 	
''  ----------------------------------------------------------------------------

function trimUTF8(byval txt)
	'convert national characters back to utf8

	Dim inp
	txt = Trim(txt)

	call ErstattKodeMedTegn( txt, 230, "æ")
	call ErstattKodeMedTegn( txt, 248, "ø")
	call ErstattKodeMedTegn( txt, 229, "å")
	call ErstattKodeMedTegn( txt, 198, "Æ")
	call ErstattKodeMedTegn( txt, 216, "Ø")
	call ErstattKodeMedTegn( txt, 197, "Å")
	call ErstattKodeMedTegn( txt, 233, "é")
	
	call ErstattKodeMedTegn( txt, 167, "§")
	
	call ErstattBokstavkodeMedTegn( txt, "lt", "<")
	call ErstattBokstavkodeMedTegn( txt, "gt", ">")
	
	trimUTF8 = txt
end function

''  ----------------------------------------------------------------------------

SUB ErstattKodeMedTegn( txt, byVal tallkode, tegn)
	
	tallkode = "&#" & tallkode & ";"
	
	call ErstattTegn( txt, tallkode, tegn)

end SUB

''  ----------------------------------------------------------------------------

SUB ErstattBokstavkodeMedTegn( txt, byVal bokstavKode, tegn)
	
	bokstavKode = "&" & bokstavKode & ";"
	
	call ErstattTegn( txt, bokstavKode, tegn)

end SUB

''  ----------------------------------------------------------------------------

SUB ErstattTegn( txt, tegn, nytttegn)
	
	if InStr(1, txt, tegn, 0) <> 0 then
		txt = Replace(txt, tegn, nytttegn, 1, -1, 0)
	end if

end SUB


''  ----------------------------------------------------------------------------
''				Timestamp   IKKE I BRUK
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


''	============================================================================
'							MODUL adocSyntaks
'
' 	Tilgjengelige tekstutformingsfunksjoner ihht. adoc-syntaksen  ##############
'	Funksjoner som starter med adoc er ikke ment å brukes fra andre moduler
'
''	============================================================================

''  ----------------------------------------------------------------------------
'
'  	Formatering av tekst (ord og fraser):
'		bold, kursiv, understrek, bokstavlig, bokstavligCelle
'
''  ----------------------------------------------------------------------------

function bold( tekst)
''	Returnerer asciidoc-kode for feit/bold tekst
''
	bold = adocFormat( tekst, "**", "")
end function 

''  ----------------------------------------------------------------------------

function kursiv( tekst)
''	Returnerer asciidoc-kode for kursiv tekst
''
	kursiv = adocFormat( tekst, "__", "")
end function 

''  ----------------------------------------------------------------------------

function understrek( tekst)
''	Returnerer asciidoc-kode for understreka tekst
''
	understrek = adocFormat( tekst, "##", "underline")
end function 

''  ----------------------------------------------------------------------------

function bokstavlig( tekst)
''	Returnerer asciidoc-kode for tekst som skal gjensgis bokstavlig
''
	bokstavlig = array( "[literal]", tekst, avsnittSkille() )

end function

''  ----------------------------------------------------------------------------

function erBokstavlig( tekst)
	if isArray(tekst) then
		erBokstavlig =  (tekst(0) = "[literal]" and UBound(tekst) = 2)
	end if
end function

''  ----------------------------------------------------------------------------

function bokstavligCelle( tekst)
''	asciidoc-kode for tekst som skal gjensgis bokstavlig i en tabell	
''
	bokstavligCelle = array ( "l|", tekst(1))

end function

''  ----------------------------------------------------------------------------

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

function bildeFrittstaaende( byVal bildetekst, bilde, alternativtekst)

	dim bildelink 
	bildelink = adocBildelink(bilde, alternativtekst, "" )

	bildetekst = adocBildeTekst(bildetekst)

	bildeFrittstaaende = array( bildetekst, bildelink, avsnittSkille( ) )
end function

''  ----------------------------------------------------------------------------

function bildeITekst( byVal bildetekst, byval bilde, byval alternativtekst)

	dim bildelink 
	bildelink = adocBildelink(bilde, alternativtekst, "width=100")

	bildeITekst = array( linjeskift(), bildetekst, bildelink)
end function 

''  ----------------------------------------------------------------------------

function adocBildelink( byVal bilde, byVal alternativtekst, byVal imagesize)
	dim fnutt : fnutt = """"

	dim bildelink : bildelink = "link=" + bilde
	alternativtekst = "alt=" + fnutt + alternativtekst + fnutt 

	dim link
	if imagesize <> "" then 
		link = array( bildelink, alternativtekst)
	else
		link = array( bildelink, imagesize, alternativtekst)
	end if
	
	link = "[" + join( link, ", ") + "]"
	
	adocBildelink = "image::" + bilde + link	
end function

''  ----------------------------------------------------------------------------

function adocBildeTekst(tekst)
	adocBildeTekst = "." & tekst
end function


' 	----------------------------------------------------------------------------
'
'	Ombrekking
'
'	----------------------------------------------------------------------------


function sideskift()
	dim kommentar
	kommentar = adocKommentar(" *********** Sideskift *********** ")
	
	sideskift = array( avsnittSkille(), kommentar, adocPageBreak() )
	
end function

''  --------

function adocPageBreak()
	adocPageBreak = "<<<"
end function

''  ----------------------------------------------------------------------------

function skillelinje( )

	dim kommentar
	kommentar = adocKommentar(" ----------- Skillelinje -----------")

	skillelinje = array( avsnittSkille(), kommentar, adocBreak() )
	
end function

''  --------

function adocBreak()
	adocBreak = "'''"
end function

''  ----------------------------------------------------------------------------

function avsnittSkille( )

	avsnittSkille = " "
	
end function

''  ----------------------------------------------------------------------------

function linjeskift( )

	linjeskift = " +"
	
end function


'	----------------------------------------------------------------------------
'
'	Overskrifter
'
'	----------------------------------------------------------------------------

function unummerertOverskrift(byVal level, byval tittel)

	tittel = adocOverskrift(level, tittel)
	
	unummerertOverskrift = array( avsnittSkille(), "[discrete]", tittel )	
end function

''  ----------------------------------------------------------------------------

function nummerertOverskrift(byVal level, byval tittel)

	tittel = adocOverskrift(level, tittel)
	
	nummerertOverskrift = array( avsnittSkille(), tittel )
end function

''  ----------------------------------------------------------------------------

function adocOverskrift(byVal level, tittel)
''  En overskrift kan være på nivå 0-5. 
''	Den angis med 1-6 stk "=" før overskriftsteksten 

	if level > nedersteOverskiftsnivaa then level = nedersteOverskiftsnivaa

	if level >=  oversteOverskiftsnivaa then
		adocOverskrift = string(level, "=") & "= " & tittel

	else       ''   kan ikke lage overskrift
		adocOverskrift = tittel
	end if
	
end function

''  ----------------------------------------------------------------------------

function overkriftMedBokmerke( tittel, byval bokmerke)
	
	overkriftMedBokmerke = array( adocBokmerke(bokmerke), tittel)

end function


'	----------------------------------------------------------------------------
'	
'	Lenker og referanser  
'
'	----------------------------------------------------------------------------


function internLenke( bokmerke, tekst)

	internLenke = "<<" + bokmerke + ", " + tekst + ">>"  	
	
end function

''  ----------------------------------------------------------------------------

function eksternLenke( uri, tekst)

	eksternLenke = uri + "[" + tekst + "]" 	
	
end function 

''  ----------------------------------------------------------------------------

function adocBokmerke(bokmerke)

	adocBokmerke = "[[" + bokmerke + "]]"
	
end function


'	----------------------------------------------------------------------------
'
'	Tabeller  
'
'	----------------------------------------------------------------------------

function tabellStart( byval kolonneBredder )
	
	dim kommentar : kommentar = adocKommentar("Topp av tabell _______________")
	
	dim fnutt : fnutt = """"
	kolonnebredder = "[cols=" + fnutt + kolonneBredder + fnutt + "]"
	
	tabellStart = array(kommentar, kolonnebredder, "|===")
end function

''  ----------------------------------------------------------------------------

function tabellAvslutning()
''  Returnerer asciidoc-kode for å avslutte en tabell
	dim kommentar : kommentar = adocKommentar("Slutt på tabell _______________")
	
	tabellavslutning = array( "|===", kommentar)
end function

''  ----------------------------------------------------------------------------

function tabellCelle( byval innhold)

	if erBokstavlig( innhold) then

		tabellCelle = bokstavligCelle( innhold)
		
	elseif not isArray( innhold) then
	
		tabellCelle = "|" & innhold & " "
		
	else
		dim res()
		redim res(Ubound(innhold)+1)
		res(0) =  "|"

		dim i
		for i = 0 to Ubound(innhold)
			if not isEmpty(innhold(i)) then 
				res(i+1) = innhold(i) + linjeskift()
			end if
		next
		redim preserve res(i)
		tabellCelle = res
	end if
	
end function

''  ----------------------------------------------------------------------------

function adocKommentar(kommentar)
	adocKommentar = "// " & kommentar
end function


''	'====================================================
''
''	OnProjectBrowserScript
''
''	'====================================================


''	============================================================================
'							MODUL Realiserte Objekttyper
'
' 	Skriver ut oversikt over objekttypenes egenskper i forflata form
'	med eller uten kolomnner for SOSI-format
'	
'	Den globale parameteren visSosiFormatRealisering bestemmer om 
'	kolonnene for SOSI-format skal være med eller ikke. 
'
''	============================================================================


Sub listRealiserteObjekttyper( pakke)

''	visSosiFormatRealisering = true	
''	visSosiFormatRealisering = false	

	dim filnavn

	dim hode, i
	if visSosiFormatRealisering then
		hode = array( "Navn", "Type", "Mult.", "SOSI-navn", "SOSI-type")
		filnavn = "SOSIformatRealisering"
	else
		hode = array( "Navn", "Type", "Mult.")	
		filnavn = "RealiserteObjekttyper"
	end if
''	hode(0) = "GML-navn på egenskap/rolle"
''	hode(0) = "Navn på egenskap/rolle"

	for i = 0 to UBound(hode)
		hode(i) = bold(hode(i))
	next

	Session.Output "// Realiserte objekttyper i " + pakke.element.name
	Session.Output "// Start of UML-model"		

	set utfil = tomTekstfil( filnavn + ".adoc")
	
	call visPakkasRealiserteObjekttyper( pakke, hode)
	
	utfil.close
	
	Session.Output "// End of UML-model"


end sub

''  ----------------------------------------------------------------------------

Sub visPakkasRealiserteObjekttyper( pakke, tabellhode)

	dim pakkenavn : pakkenavn = pakke.Name

	dim tabellformat
	if UBound(tabellhode) = 2 then tabellformat = "20,20,10"
	if UBound(tabellhode) = 4 then tabellformat = "20,20,10,20,10"  

	dim overskriftsnivaa 
	overskriftsnivaa = 3
	
	dim overskrift
	overskrift = unummerertOverskrift( overskriftsnivaa, "Pakke : " + pakkenavn)
''	settInnTekst overskrift   ''  Pakka kan være tom for realiserbare klasser...
	overskriftsnivaa = 4
	
	Dim element As EA.Element
	For each element in pakke.Elements
		dim ster : ster = Ucase(element.Stereotype)
		If element.Type <> "Class" then  	'' hopp over
		elseif element.Abstract = 1 then	'' hopp over abstrakte klasser
		elseif (ster = "FEATURETYPE" or ster = "") Then
		
			dim elementNavn  '' navnet på featuretype
			elementNavn = pakkenavn + "::" + stereotypenavn( element)
''			elementNavn = "Objekttype: " + element.Name)
			settInnTekst unummerertOverskrift( overskriftsnivaa, elementNavn)
						
			if visSosiFormatRealisering then
				dim avgrensingslinjer, avgrensingsflater
				avgrensingslinjer = avgrensingsrelasjoner(element,"avgrensesAv")
				avgrensingsflater = avgrensingsrelasjoner(element,"avgrenser")
			
				dim navnestreng
				if UBound(avgrensingslinjer) >= 0 then 
					navnestreng = join( navneliste(avgrensingslinjer), ", ")
					settInnTekst bold("Avgrenses av: ") + navnestreng
				end if		
				if UBound(avgrensingsflater) >= 0 then 
					navnestreng = join( navneliste(avgrensingsflater), ", ")
					settInnTekst bold("Avgrenser: ") + navnestreng
				end if
			end if
								
			dim attributtListe
			attributtListe = egenskaperOgRoller( element, avgrensingslinjer)
			
			if UBound(attributtListe) > 0 then		
				attributtListe(0) = tabellhode
				call SettInnSomTabell( attributtListe, tabellFormat, "")
			end if
			
		End if		
	Next

	''	Gå rekursivt gjennom alle underpakker
	dim delpakke as EA.Package
	for each delpakke in pakke.Packages
	
		Call visPakkasRealiserteObjekttyper( delpakke, tabellhode)
		
	next
	
end sub

''  ----------------------------------------------------------------------------

function avgrensingsrelasjoner( featureType, relasjonsType)

	''	finn først featuretypens avgrensingsrelasjoner
	''	deretter fra supertypen, som legges først i lista

	dim liste()
	redim liste(featuretype.Connectors.Count)
	dim relNr : relNr = 0
	
	Dim con
	For Each con In featureType.Connectors
	
		dim avgrensing
		dim elementID
	
		if relasjonsType = "avgrensesAv" then
			elementID = avgrensesAv( featureType, con)
		elseif relasjonsType = "avgrenser" then
			elementID = avgrenser( featureType, con) 
		end if
		
		if elementID <> 0 then
			liste(relNr) = elementID
			relNr = relNr +1
		end if
		
	next
	redim preserve liste(relNr-1)

	if antallSupertyper( featureType) = 1 then
		'' egenskaper og roller fra supertypen legges først i lista
		dim super
		super = avgrensingsrelasjoner( superType(featureType), relasjonsType)
		avgrensingsrelasjoner = merge( super, liste)
	else
		avgrensingsrelasjoner = liste
	end if

end function

''  ----------------------------------------------------------------------------

function avgrensesAv( element, con)

'' Finn ut om elementet geometrisk kan avgrenses av et annet element (target)
'' I så fall returneres target-elementets ID

	dim elementId : elementId = element.elementID
	dim targetID, target

	if elementId = con.SupplierID then
		targetID = con.ClientID
		set target = con.ClientEnd	
		
	elseif elementId = con.ClientID  then
		targetID = con.SupplierID
		set target = con.SupplierEnd
	
	else 
		exit function
	end if

	'	Det er tre måter å avgøre om target er et avgrensingsobjekt:
	' SOSI 4.5:	conn.Stereotype = "Topo". Kan navigere fra current til target
	' SOSI 5.0:	element.Constraints inneholder 'KanAvgrensesAv Target.Name'
	' FKB-praksis:	target.Role = "avgrensesAv"
	'
	'	Det er bare FKB-varianten som er lagt inn
	'
	if inStr(LCase(target.Role),"avgrensesav") <> 0 then
		avgrensesAv = targetID
	end if
	
end function

''  ----------------------------------------------------------------------------

function avgrenser( element, con)

'' Finn ut om elementet geometrisk kan avgrense et annet element (target)
'' I så fall returneres target-elementets ID

	dim elementId : elementId = element.elementID
	dim targetID, current

	if elementId = con.SupplierID then
		targetID = con.ClientID
		set current = con.SupplierEnd	
		
	elseif elementId = con.ClientID  then
		targetID = con.SupplierID
		set current = con.ClientEnd
	
	else 
		exit function
	end if
	
	'	Det er tre måter å avgøre om target er ei flate som avgrenses:
	' SOSI 4.5:	conn.Stereotype = "Topo". Kan navigere fra target til current
	' SOSI 5.0:	element.Constraints inneholder 'KanAvgrensesAv Current.Name'
	' FKB-praksis:	current.Role = "avgrensesAv"
	'
	'	Det er bare FKB-varianten som er lagt inn
	'
	if inStr(LCase(current.Role),"avgrensesav") <> 0 then
		avgrenser = targetID
	end if
	
end function

''  ----------------------------------------------------------------------------

function egenskaperOgRoller( featureType, avgrensingslinjer)

	''	finn først featuretypens egen egenskaper og roller
	''	deretter fra supertypen, som legges først i lista
	
	dim egensk, roller
	egensk = egneEgenskaper( featureType, "", ".")
	roller = egneRoller( featureType, avgrensingslinjer)
	egensk = merge( egensk,	roller)
	
	
	if antallSupertyper( featureType) = 0 then
		''	Denne featuretypen har ingen supertyper
		dim plassTilTabellhode()
		redim plassTilTabellhode(0)   '' første linje i lista holdes tom
		egenskaperOgRoller = merge( plassTilTabellhode, egensk)  
		
''	elseif antallSupertyper(featureType) > 1 then
''		'' FEILSITUASJON

	else  '' harSupertype(featureType) = true
		'' egenskaper og roller fra supertypen legges først i lista
		dim super
		super = egenskaperOgRoller( superType(featureType), avgrensingslinjer)
		egenskaperOgRoller = merge( super, egensk)
	end if

end function

''  ----------------------------------------------------------------------------

function egneEgenskaper( element, byVal egenskapsgruppe, byval sosiPrikker)

	if egenskapsgruppe <> "" then egenskapsgruppe = egenskapsgruppe + "."
	sosiPrikker = sosiPrikker + "."
	
	dim liste()
	redim liste(element.Attributes.Count-1)

	dim radNr : radNr = 0

	Dim att As EA.Attribute
	for each att in element.Attributes
		Dim datatype As EA.Element

		dim egNavn, mult, dtyp
		if att.ClassifierID <> 0 then
			set datatype = Repository.GetElementByID(att.ClassifierID)
			dtyp = stereotypeNavn( datatype)
		else
			datatype = null
			dtyp = att.Type
		end if
		egNavn = egenskapsgruppe & att.Name 
		mult = "[" & att.LowerBound & ".." & att.UpperBound & "]"
		
		dim attrib
		attrib = array(egNavn, dtyp, mult)
		
		if visSosiFormatRealisering then
			dim sNavn, sType
			sNavn = sosiNavn(att, sosiPrikker)

			sType = sosiDatatype(att)
			
			if sType = "" and not isnull(datatype) then 
				'' hent sosi-typen fra datatypen
				if iscodelist( datatype) then 
					sType = sosiDatatype(dataType)
				elseif isDataType( dataType) then
					sType = "*"
				end if
			end if
			
			attrib = merge( attrib, array(sNavn, sType) )
		end if
		
		liste(radNr) = attrib
		radNr = radNr +1
		
		dim underEgenskaper : underEgenskaper = null
				
		if not isNull(datatype) then
			if LCase(datatype.Stereotype) = "datatype" then
				underEgenskaper = egneEgenskaper( datatype, egNavn, sosiPrikker)
			end if
			if not isNull(underEgenskaper) then
				redim preserve liste(UBound(liste) + UBound(underEgenskaper)+1 )
				dim eg
				for each eg in underEgenskaper
					liste(radNr) = eg				
					radNr = radNr +1
				next
			end if
		end if
		
	next

	redim preserve liste(radNr-1)
	
	egneEgenskaper = liste

end function

''  ----------------------------------------------------------------------------

function egneRoller( featureType, avgrensingslinjer)

	dim sosiPrikker : sosiPrikker = ".."
	
	dim liste()
	redim liste(featureType.Connectors.Count)

	dim radNr : radNr = 0
	
	dim rollesamling  	'' array av rolle-arryer
''	rollesamling = alleRoller( featureType)
	rollesamling = sorterteRoller( featureType)
	if isEmpty(rollesamling) then 	EXIT function
	
	dim target, targetID
	dim rollenavn, dtyp, mult
	dim datatype

	dim rolle   		
	for each rolle in rollesamling 
		targetID = rolle(4)

		if not listeInneholder(avgrensingslinjer, targetID) then
			set target = rolle(3)
			rollenavn = bold("Rolle: ") + target.Role
			mult = "[" & target.Cardinality & "]"
			if targetID <> 0 then
				set datatype = Repository.GetElementByID(targetID)
				dtyp = stereotypenavn( datatype)
			else
				set datatype = nothing
				dtyp = ""
			end if
			dim rolleEgenskaper 
			rolleEgenskaper = array(rollenavn, dtyp, mult)
			
			if visSosiFormatRealisering then
				dim sNavn, sType
				sNavn = taggedValueFraRolle(target, "SOSI_navn") 
				if sNavn <> "" then sNavn = sosiPrikker & sNavn
				sType = "REF"

				rolleEgenskaper = merge( rolleEgenskaper, array(sNavn, sType))
	''			rolleEgenskaper = array(rollenavn, dtyp, mult, sNavn, sType)
			end if

			liste(radNr) = rolleEgenskaper
			radNr = radNr +1
		end if
	next
	
	redim preserve liste(radNr-1)
	egneRoller = liste

end function

''  ----------------------------------------------------------------------------

function antallSuperTyper( elem)

	antallSuperTyper = 0
	
	dim conn as EA.Collection
	for each conn in elem.Connectors
		if conn.Type = "Generalization" and elem.ElementID = conn.ClientID then
			antallSuperTyper = antallSuperTyper +1
		end if
	next
		
end function

''  ----------------------------------------------------------------------------

function harSuperType( elem)

	harSuperType = false
	
	dim conn as EA.Collection
	for each conn in elem.Connectors
		if conn.Type = "Generalization" and elem.ElementID = conn.ClientID then
			harSuperType = true
			exit for
		end if
	next
		
end function

''  ----------------------------------------------------------------------------

function superType( elem)

	dim conn as EA.Collection
	for each conn in elem.Connectors
		if conn.Type = "Generalization" and elem.ElementID = conn.ClientID then
			set superType = Repository.GetElementByID(conn.SupplierID)				
			exit for
		end if
	next
		
end function

''  ----------------------------------------------------------------------------

function sosiNavn( element, sosiPrikker)

	dim sNavn : sNavn = taggedValueFraElement(element, "SOSI_navn")
	dim sosiGeometri : sosiGeometri = sosiGeometritype( element)

	if sosiGeometri <> "" then 
		sosiNavn = "." + sosiGeometri 
	elseif sNavn <> "" then 
		sosiNavn = sosiPrikker & sNavn
	end if

end function

''  ----------------------------------------------------------------------------

function sosiDatatype( element)

	dim sosiType, sosiLengde
	sosiType   = taggedValueFraElement(element, "SOSI_datatype")
	sosiLengde = taggedValueFraElement(element, "SOSI_lengde")
	
	sosiDatatype = sosiType & sosiLengde
	
end function

''  ----------------------------------------------------------------------------

function navneliste( idListe)

	if isNull(idListe) then exit function
	
	dim liste : liste = idListe
	dim id, element
	dim i : i = 0
	for each id in idliste
		set element = Repository.GetElementByID( id)
		liste(i) = element.Name
		i = i+1
	next

	navneliste = liste

end function

''  ----------------------------------------------------------------------------

function sosiGeometritype(element)

	dim gtype : gtype = element.Type

	if gtype = "Punkt" or gtype = "Kurve" or gtype = "Flate" then
	'' if listeInneholder(array("Punkt", "Kurve", "Flate"), gtype) then
		sosiGeometritype = UCase( gtype)

	'fra Ralisering i SOSI-format versjon 5.0 tabell 8.2:
	elseif gtype = "GM_Point" then
		sosiGeometritype = "PUNKT"
		
	elseif gtype = "GM_MultiPoint" then
		sosiGeometritype = "SVERM"
		
	elseif gtype = "GM_Curve" or gtype = "GM_CompositeCurve" then
		sosiGeometritype = "KURVE"
		
	elseif gtype = "GM_Surface" or gtype = "GM_CompositeSurface" then
		sosiGeometritype = "FLATE"
		
	'fra "etablert praksis"
	elseif gtype = "GM_Object" or gtype = "GM_Primitive" then
		sosiGeometritype = "OBJEKT"
		
	else
		sosiGeometritype = ""
	end if
	
end function


'====================================================

OnProjectBrowserScript

'====================================================
