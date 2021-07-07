option explicit 
 
 !INC Local Scripts.EAConstants-VBScript 
 
'  
' Script Name: realiserbarGMLformat50 
' Author: Kent Jonsrud - Section for standardization and technology development - Norwegian Mapping Authority
'
' Version: alfa0.3
' Date: 2021-07-07	renamed from realiserbarGML50.vb to realiserbarGMLformat50.vb
'
' Version: alfa0.2
' Date: 2018-03-27

' Purpose: Validerer om modellen er realiserbar etter standarden SOSI Realisering i GML v.5.0 
' Purpose: Validate model elements according to rules defined in the standard SOSI Regler for UML-modellering 5.0 
' Implemented rules: 
' /krav/tegnsett	Tegnsett for alle tegn i alle GML datasett skal være UTF-8, og dette skal dokumenteres i XML-attributtet encoding i XML-angivelsen i starten på fila.	15
' /krav/filhode	Datafiler som skal brukes til å utveksle geografisk informasjon på GML-format skal inneholde et standard filhode og et standard filsluttmerke. Det skal ikke være datainnhold etter filsluttmerket. Filhodet skal inneholde angivelse av tegnsett, formatversjon og angivelse av datasettets spesifikasjon.	15
' /krav/WFS-konteiner	 Konteineren skal være wfs:FeatureCollection i versjon 2.0. Konteineren skal ha navnerommet http://www.opengis.net/wfs/2.0/wfs og det skal pekes til XML-skjemabeskrivelse i http://schemas.opengis.net/wfs/2.0/wfs.xsd	15
' /krav/rekkefølge	Det er ved bruk av XML påkrevet å følge en fast rekkefølge på elementer innen samme grupperingsnivå, rekkefølgen i GML skal være den samme som rekkefølgen som er i modellen.	15
' /krav/GML-formatversjon	 Formatversjonen skal ha navnerommet http://www.opengis.net/gml/3.2 og peke til sin XML-skjemabeskrivelse i http://schemas.opengis.net/gml/3.2.1/gml.xsd Alternativt kan formatversjon være en eller flere fra GML 3.3 da denne direkte bygger på versjon 3.2.	16
' /krav/koordinatreferansesystem	Koordinatreferansesystemkoden skal være lik for alle geometrier i et datasett. Alle geometrier i datasettet skal ha oppgitt koordinatreferansesystemkode dersom dette ikke er oppgitt i filhodet.	17
' --/krav/produktnavnerom	Tagged value targetNamespace i UML-modellen skal realiseres som http-URI og URL til navnerommet. Benevnelsen https: skal ikke være med i en http-URI men den skal alternativt også kunne benyttes til å referere til skjemafila.	19
' --/krav/produktforstavelse	Tagged value xmlns i UML-modellen angir et XML kortnavn for navnerommet (vanligvis app) og skal være med som navneromsprefiks i datasett hvis datasettet ikke angir navnerommet som default namespace. (xmlns="http://skjema.geonorge.no/SOSI/produktspesifikasjon/Stedsnavn/5.0")	19
' --/krav/produktbeskrivelse	Tagged value xsdDocument i UML-modellen angir filnavn på ei GML-applikasjonsskjemafil som skal være tilgjengelig i navnerommet.	19
' --/krav/produktversjon	Tagged value version i UML-modellen beskriver versjonen i full detalj. Første del av versjon skal realiseres likt med siste del av navnerommet.	19
' /krav/objektstereotype	Klasser med stereotype «FeatureType» skal alltid realiseres direkte som XML-elementer på toppnivå under konteinerens XML-element <wfs:member>.	20
' /krav/objekttype	Modellelementnavnet på klasser med stereotype «FeatureType» skal realiseres i xsd-fila som <xsd:complexType> med endelsen "Type" etter klassenavnet. Klassenavnet benyttes direkte som navn på XML-elementet.	20
' /krav/objektidentifikator	Elementet skal ha XML-egenskapen gml:id med verdi som skal være unik innenfor datasettet.	21
' /krav/objektegenskap	Navn på egenskaper i klasser med stereotype «FeatureType» er modellelementnavn som skal realiseres ordrett som XML-elementer under objekttypens XML-element.		21
' /krav/objektegenskapstype	 Egenskapstyper som er en brukerdefinert klasse skal realiseres som en <xsd:complexType> med alt innhold fra denne klassen. Se 7.7. Egenskapstyper som er basistyper skal realiseres direkte som angitt xsd-basistype. Se Tabell 7.1	22
' /krav/tekst	 Egenskaper av type CharacterString der teksten inneholder "&", "<" eller ">" skal disse tegnene endres til henholdsvis &amp; &lt; &gt; fordi ellers vil disse tegnene kunne oppfattes som escapetegn eller start og slutt på XML-elementnavn. Alle andre tegn skal være korrekte og utranslitererte UTF-8-tegn.	23
' /krav/geometriegenskap3D	 For egenskaper med geometrityper i 3D (GM_Solid) skal koordinatreferansesystemet være et system med 3D koordinater.	23
' i--/krav/heleid2Dgeometri	Datasett og tjenester som erklærer at de er konforme med konformitetsklasse SOSI-GML-heleid2Dgeometri skal kun benytte geometritype fra Tabell 7.2.		23
' --/krav/heleid3Dgeometri	Datasett og tjenester som erklærer at de er konforme med konformitetsklasse SOSI-GML-heleid3Dgeometri skal kun benytte geometritype fra Tabell 7.2 og Tabell 7.3.		23
' --/krav/delt2Dgeometri	Datasett og tjenester som erklærer at de er konforme med konformitetsklasse SOSI-GML-delt2Dgeometri skal kun benytte geometritype fra tabell 7.4			24
' /krav/akserekkefølge	Rekkefølgen på aksene skal alltid være den akserekkefølgen som koordinatreferansesystembeskrivelsen angir. Se Tabell 6.1 Standardiserte koordinatsystemkoder.	24
' /krav/akseantall	 Antall akser skal alltid være lik det antallet som angitt i koordinatreferansesystembeskrivelsen. Se Tabell 6.1 Standardiserte koordinatsystemkoder.	24
' /krav/akseenhet	 Koordinaters enhet for hver akse skal være den samme enheten som er beskrevet i koordinatreferansesystembeskrivelsen. Se Tabell 6.1 Standardiserte koordinatsystemkoder.	25
' /krav/rollerekkefølge	 Alle assosiasjonsroller skal ha en tagged value sequenceNumber med verdi som angir rekkefølgen elementene skal komme i. Alle egenskaper uten tagged value sequenceNumber skal komme i den rekkefølge de er vist i modellen, og de skal komme før alle assosiasjonrollene.	25
' /krav/objekttyperolle	Navn på assosiasjonsroller fra aggregeringer eller vanlige assosiasjoner til klasser med stereotype «FeatureType» skal realiseres direkte som XML-elementnavn som inneholder xlink til det refererte objektet, som enten er internt i datasettet (local) eller i et eksternt datasett (remote).	25
' /krav/datatyperolle	Navn på assosiasjonsroller fra komposisjoner til klasser med stereotype «Union» eller «dataType» er modellelementnavn som skal realiseres direkte som XML-elementer inline i eierobjektet.	25
' /krav/datatype	Modellelementet skal realiseres direkte via navnet på egenskapen eller assosiasjonsrollen som peker til datatypeklassen.  Egenskaper i datatypen skal realiseres på samme måte som objektegenskaper. Assosiasjonsroller i datatypen skal realiseres på samme måte som vanlige roller.	28
' /krav/union	En klasse med stereotype «Union» beskriver et sett med mulige egenskaper. Kun en av egenskapene kan forekomme i hver instans. Modellelementet skal først realiseres direkte fra navnet på den egenskapen som bruker unionen og deretter klassenavnet til unionen og til slutt det valgte UML-modellelementnavnet i unionen.	29
' /krav/enumerering	En klasse med stereotype «enumeration» beskriver et lukket sett med lovlige koder. Kun en av disse kodene kan forekomme i en instans. Modellelementet skal realiseres direkte som enumererte verdier i GML-applikasjonsskjema.	30
' /krav/skjemakodeliste	Dersom modellelementet har en tagged value asDictionary med verdi false, eller mangler denne tagged value, skal kodene realiseres direkte som enumererte verdier i GML-applikasjonsskjema.	30
' --/krav/koderegister	 Dersom kodelista er implementert i et register, angitt med tagged value asDictionary = true, skal koden valideres mot verdier i det levende registeret.	30
' --/krav/koderegistersti 	Sti til registeret skal finnes i verdien til en tagged value codeList i kodelisteklassen. På egenskaper som bruker kodelista skal tilsvarende sti stå i en tagged value defaultCodeSpace.	30
' /krav/kode	 Elementer i en klasse med stereotype «CodeList» eller «enumeration» beskriver lovlige koder. Modellelementet skal realiseres slik at kodens navn benyttes direkte i datasettet. (Ref. krav om NCName på koder). Dersom koden har en initialverdi skal denne initialverdien benyttes i datasettet istedenfor kodens navn.	31
' /krav/nøsteretning	Det kreves at geometrien til ytre flateavgrensninger nøstes i retning mot klokka, og indre avgrensinger i retning med klokka.	32
' i--/krav/eldreGeometritype 	Modellering av geometri i eldre fagområdestandarder skal kunne realiseres regelstyrt til GML i en produktspesifikasjon ved bruk av Tabell 8.1.	35
' /krav/segmenttype	Geometrisegmenttyper skal være en av de som er beskrevet i Tabell 8.2			36
' --/krav/GMLtopologi 	Realisering av topologi i GML-applikasjonsskjema skal benytte topologityper som angitt i tabellen i ISO 19136 vedlegg D, se Tabell D.2 .	40
' /req/temporal/	Egenskaper med temporale datatyper skal realiseres	44
' /krav/abstraktGeometri 	Abstrakte geometrityper i modellen skal realiseres som enhver av de realiserbare subtypene. Se Tabell 8.9.	45
'
'
'
' /anbefaling/tekstformat	GML er et standardisert vokabular i tekstformatet XML og bør normalt ikke ha noe binærinnhold. GML-formatet bør ikke åpne for inkonsistens om datasettets tegnsett og bør ikke inneholde alternative binære tegnsettidentifikatorer som BOM (Byte Order Mark – ofte brukt til å angi tegnsett UTF-8 der formatet ikke har mulighet for å kode tegnsettinformasjon)	15
' /anbefaling/formKoordinatreferansesystem	Det anbefales å bruke http-URI-form eller URN på angivelsen av koordinatreferansesystemkodene for nyere datasett da disse kan brukes direkte som URL til beskrivelsen.		18
' /anbefaling/kodingsregel	Det anbefales å benytte en tagged value xsdEncodingRule i UML-modellen med verdien sosi50 som angir i detalj hvordan skjemagenereringen skal utføres i henhold til denne standarden. Alternativt kan verdien settes til sosi som angir beste praksis etter reglene i standarden SOSI 4.x.	19
' /anbefaling/enkeltGeometrisegment	Det anbefales å ikke legge inn flere geometrisegmenter i hver geometriprimitiv i datasett.	36
' /anbefaling/alternativeGeometrityper	For nyere datasett anbefales det å utveksle GML-data på den fulle modellbaserte måten som er beskrevet i vedlegg B i dette dokumentet. Programvare som robust skal kunne lese fra alle mulige kilder bør likevel kunne gjenkjenne alle de alternative forenklede geometritypene og mappe dem til de modellbaserte.		40
' /anbefaling/topologibruk 	Topologi bør ikke brukes i et applikasjonsskjema med mindre det er absolutt påkrevet.		40
' /anbefaling/CoveregeType 	Coverage-typer brukt i et applikasjonsskjema bør være enten RectifiedGridCoverage, ReferencableGridCoverage eller TimeSeriesTVP fra WaterML 2.0, eller en subtype.		41
' /anbefaling/tid 	 Det anbefales å benytte UTC som referansesystem for tid i alle temporale egenskaper da disse kan sammenlignes presist med andre temporale data uten å få inkonsistens på grunn av ulike tidssone og sommertidsoverganger.	45



'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Project Browser Script main function
'
sub OnProjectBrowserScript()
	
	Repository.EnsureOutputVisible("Script")
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	
	' Handling Code: Uncomment any types you wish this script to support
	' NOTE: You can toggle comments on multiple lines that are currently
	' selected with [CTRL]+[SHIFT]+[C].
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
 			
			if not thePackage.IsModel then
				'check if the selected package has stereotype applicationSchema 
 			
				if UCase(thePackage.element.stereotype) = UCase("ApplicationSchema") then 
				
					dim box, mess
					mess = "Indikerer om modellen kan realiseres etter standarden 'SOSI Realisering i GML versjon 5.0'"&Chr(13)&Chr(10)
					mess = mess + ""&Chr(13)&Chr(10)
					mess = mess + "En liste med kravene det testes for ligger i kildekoden (linje 15++)."&Chr(13)&Chr(10)
					mess = mess + ""&Chr(13)&Chr(10)
					mess = mess + "Starter validering av pakke [" & thePackage.Name &"]."&Chr(13)&Chr(10)

					box = Msgbox (mess, vbOKCancel, "realiserbarGML50 alfa0.2-2018-03-27")
					select case box
						case vbOK
							dim logLevelFromInputBox, logLevelInputBoxText, correctInput, abort
							logLevelInputBoxText = "Velg loggnivå."&Chr(13)&Chr(10)
							logLevelInputBoxText = logLevelInputBoxText+ ""&Chr(13)&Chr(10)
							logLevelInputBoxText = logLevelInputBoxText+ ""&Chr(13)&Chr(10)
							logLevelInputBoxText = logLevelInputBoxText+ "E - Feil (Error): kun meldinger om direkte feil."&Chr(13)&Chr(10)
							logLevelInputBoxText = logLevelInputBoxText+ ""&Chr(13)&Chr(10)
							logLevelInputBoxText = logLevelInputBoxText+ "W - Advarsel (Warning): melder både feil og advarsler."&Chr(13)&Chr(10)
							logLevelInputBoxText = logLevelInputBoxText+ ""&Chr(13)&Chr(10)
							logLevelInputBoxText = logLevelInputBoxText+ "Angi E eller W:"&Chr(13)&Chr(10)
							correctInput = false
							abort = false
							do while not correctInput
						
								logLevelFromInputBox = InputBox(logLevelInputBoxText, "Velg loggnivå", "W")
								select case true 
									case UCase(logLevelFromInputBox) = "E"	
										globalLogLevelIsWarning = false
										correctInput = true
									case UCase(logLevelFromInputBox) = "W"	
										globalLogLevelIsWarning = true
										correctInput = true
									case UCase(logLevelFromInputBox) = "D"	
										globalLogLevelIsWarning = true
										debug = true
										correctInput = true
									case IsEmpty(logLevelFromInputBox)
										MsgBox "Abort",64
										abort = true
										exit do
									case else
										MsgBox "Du valgte et ukjent loggnivå, velg 'E' eller 'W'.",48
								end select
							
							loop
							

							if not abort then
								'give an initial feedback in system output 
								Session.Output("realiserbarGML50 alfa0.2 startet. "&Now())
								'Check model for script breaking structures
								if scriptBreakingStructuresInModel(thePackage) then
									Session.Output("Kritisk feil: Kan ikke validere struktur og innhold før denne feilen er rettet.")
									Session.Output("Aborterer skript.")
									exit sub
								end if
							
							'	call populatePackageIDList(thePackage)
							'	call populateClassifierIDList(thePackage)
							'	call findPackageDependencies(thePackage.Element)
							'	call getElementIDsOfExternalReferencedElements(thePackage)
							'	call findPackagesToBeReferenced()
							'	call checkPackageDependency(thePackage)
							'	call dependencyLoop(thePackage.Element)
							  
                'For /req/Uml/Profile:
							  Set ProfileTypes = CreateObject("System.Collections.ArrayList")
							  Set ExtensionTypes = CreateObject("System.Collections.ArrayList")
							  Set CoreTypes = CreateObject("System.Collections.ArrayList")
							  reqUmlProfileLoad()
								'For /krav/18:
							'	set startPackage = thePackage
							'	Set diaoList = CreateObject( "System.Collections.Sortedlist" )
							'	Set diagList = CreateObject( "System.Collections.Sortedlist" )
							'	recListDiagramObjects(thePackage)

								Dim StartTime, EndTime, Elapsed
								StartTime = timer 
								startPackageName = thePackage.Name
								FindInvalidElementsInASPackage(thePackage) 
								Elapsed = formatnumber((Timer - StartTime),2)

								'final report
								Session.Output("-----Rapport for pakke ["&startPackageName&"]-----") 		
								Session.Output("   Antall feil funnet: " & globalErrorCounter) 
								if globalLogLevelIsWarning then
									Session.Output("   Antall advarsler funnet: " & globalWarningCounter)
								end if	
								Session.Output("   Kjøretid: " &Elapsed& " sekunder" )
							end if	
						case VBcancel
							'nothing to do						
					end select 
				else 
 				Msgbox "Pakken [" & thePackage.Name &"] har ikke stereotype «ApplicationSchema». Velg en pakke med stereotype «ApplicationSchema»." 
				end if
			else
			Msgbox "Pakken [" & thePackage.Name &"] er en rotpakke og har ikke stereotype «ApplicationSchema». Velg en vanlig pakke med stereotype «ApplicationSchema»."
 			end if
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
			Session.Prompt "[Warning] You must select a package with stereotype ApplicationSchema in the Project Browser to start the validation.", promptOK 
			
	end select
	
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: FindInvalidElementsInPackage
' Author: Kent Jonsrud
' Date: 2018-02-09
' Purpose: Test the content of the top level package

sub FindInvalidElementsInASPackage(package) 
			
	call kravProduktnavnerom(package)
	call kravProduktforstavelse(package)
	call kravProduktbeskrivelse(package)
	call kravProduktversjon(package)







	call FindInvalidElementsInPackage(package) 



end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------

'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: FindInvalidElementsInPackage
' Author: Kent Jonsrud
' Date: 2018-02-09
' Purpose: Main loop iterating all elements in selected package and all subpackages, conducting tests on their elements

sub FindInvalidElementsInPackage(package) 
			
 	dim elements as EA.Collection 
 	set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages) 
 	Dim myDictionary 
 	dim errorsInFunctionTests 

	if debug then Session.Output("Debug: package to be tested: [«" &package.element.Stereotype& "» " &package.Name& "].")

	dim i 
	for i = 0 to elements.Count - 1 
		dim currentElement as EA.Element 
		set currentElement = elements.GetAt( i ) 
		if debug then Session.Output("Debug: class to be tested: [«" &currentElement.Stereotype& "» " &currentElement.Name& "].")

		if currentElement.Type = "Class" Or currentElement.Type = "Enumeration" Or currentElement.Type = "DataType" then 

			call kravObjekttype(currentElement)

			if UCase(currentElement.Stereotype) = "FEATURETYPE"  Or UCase(currentElement.Stereotype) = "DATATYPE" Or UCase(currentElement.Stereotype) = "UNION" or currentElement.Type = "DataType" then
				dim attributesCollection as EA.Collection 
				set attributesCollection = currentElement.Attributes 
				 
				if attributesCollection.Count > 0 then 
					dim n 
					for n = 0 to attributesCollection.Count - 1 					 
						dim currentAttribute as EA.Attribute		 
						set currentAttribute = attributesCollection.GetAt(n) 

						if debug then Session.Output("Debug: attribute to be tested: [«" &currentAttribute.Stereotype& "» " &currentAttribute.Name& "].")

						call kravObjektegenskap(currentAttribute)
						call kravObjektegenskapstype(currentAttribute)

					next
				end if
						'if debug then Session.Output("Debug: role to be tested: [«" &role.Stereotype& "» " &role.Name& "].")
						'if debug then Session.Output("Debug: operation to be tested: [«" &operation.Stereotype& "» " &operation.Name& "].")
						'if debug then Session.Output("Debug: constraint to be tested: [«" &constraint.Stereotype& "» " &constraint.Name& "].")
			end if
			if UCase(currentElement.Stereotype) = "CODELIST"  Or UCase(currentElement.Stereotype) = "ENUMERATION" or currentElement.Type = "Enumeration" then
			end if
		end if
	next


	dim subP as EA.Package
	for each subP in package.packages
		call FindInvalidElementsInPackage(subP) 
	next




end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------

'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: kravProduktnavnerom
' Author: Kent Jonsrud
' Date: 2018-02-09
' Purpose: Test om tagged value targetNamespace finnes i pakka og om den har en gyldig uri som verdi.

sub kravProduktnavnerom(package)


end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: kravProduktforstavelse
' Author: Kent Jonsrud
' Date: 2018-02-09
' Purpose: Test om tagged value xmlns finnes i pakka og om den har en verdi. Info/advarsel dersom verdien ikke er app.

sub kravProduktforstavelse(package)
	dim forstavelse
	forstavelse = getPackageTaggedValue(package,"xmlns")
	if debug then Session.Output("Debug: package tagged value xmlns: " & forstavelse & " [«" &package.element.Stereotype& "» " &package.Name& "]. [/krav/produktforstavelse]")
	if len(forstavelse) = 0 then
			Session.Output("Error: missing package tagged value xmlns: [«" &package.element.Stereotype& "» " &package.Name& "]. [/krav/produktforstavelse]")
			globalErrorCounter = globalErrorCounter + 1
	end if
	if len(forstavelse) > 0 and forstavelse <> "app" then
			Session.Output("Warning: package tagged value xmlns is not app but: " & forstavelse & " [«" &package.element.Stereotype& "» " &package.Name& "]. [/krav/produktforstavelse]")
			globalWarningCounter = globalWarningCounter + 1
	end if
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: kravProduktbeskrivelse
' Author: Kent Jonsrud
' Date: 2018-02-09
' Purpose: Test om tagged value targetNamespace finnes i pakka og om den har en gyldig uri som verdi.

sub kravProduktbeskrivelse(package)


end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: kravProduktversjon
' Author: Kent Jonsrud
' Date: 2018-02-09
' Purpose: Test om tagged value targetNamespace finnes i pakka og om den har en gyldig uri som verdi.

sub kravProduktversjon(package)


end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: kravObjekttype
' Author: Kent Jonsrud
' Date: 2018-02-09
' Purpose: Test om tagged value targetNamespace finnes i pakka og om den har en gyldig uri som verdi.

sub kravObjekttype(element)


end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: kravObjektegenskap
' Author: Kent Jonsrud
' Date: 2018-02-09
' Purpose: Test om tagged value targetNamespace finnes i pakka og om den har en gyldig uri som verdi.

sub kravObjektegenskap(attr)


end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------



'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: kravObjektegenskapstype
' Author: Kent Jonsrud
' Date: 2018-03-27
' Purpose: Egenskapstyper skal være brukerdefinerte klasser eller kjente geometri- eller basistyper.

sub kravObjektegenskapstype(attr)
	'Iso 19109 Requirement /req/uml/profile - well known types. Including Iso 19103 Requirements 22 and 25
	if debug then Session.Output("Debug: datatype to be tested: [" &attr.Type& "].")
	if attr.ClassifierID <> 0 then
		dim datatype as EA.Element
		set datatype = Repository.GetElementByID(attr.ClassifierID)
		if datatype.Name <> attr.Type then
			Session.Output("Error: attribute [" &attr.Name& "] has a type name ["&attr.Type&"] that is not corresponding to its linked type name ["&datatype.Name&"].")
		end if
	else
		call reqUmlProfile(attr)
	end if
	
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------



' -----------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: reqUmlProfile
' Author: Kent Jonsrud
' Date: 2016-08-08, 2017-05-13, 2018-03-27
' Purpose: 
    'iso19109:2015 /req/uml/profile , includes iso109103:2015 requirement 25 and requirement 22.


sub reqUmlProfile(attr)
	
	'dim attr as EA.Attribute
	'navigate through all attributes 
	'for each attr in theElement.Attributes
		'if attr.ClassifierID = 0 then
			'Attribute not connected to a datatype class, check if the attribute has a iso TC 211 well known type
			if ProfileTypes.IndexOf(attr.Type,0) = -1 then	
				if ExtensionTypes.IndexOf(attr.Type,0) = -1 then	
					if CoreTypes.IndexOf(attr.Type,0) = -1 then	
						'Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has unknown type for attribute ["&attr.Name&" : "&attr.Type&"]. [/req/uml/profile] & krav/25 & krav/22")
						Session.Output("Error: unknown type for attribute ["&attr.Name&" : "&attr.Type&"]. [/req/uml/profile] & krav/25 & krav/22")
						globalErrorCounter = globalErrorCounter + 1 
					end if
				end if
			end if
		'end if 
	'next

end sub


sub reqUmlProfileLoad()
	
	'iso 19103:2015 Core types
	CoreTypes.Add "Date"
	CoreTypes.Add "Time"
	CoreTypes.Add "DateTime"
	CoreTypes.Add "CharacterString"
	CoreTypes.Add "Number"
	CoreTypes.Add "Decimal"
	CoreTypes.Add "Integer"
	CoreTypes.Add "Real"
	CoreTypes.Add "Boolean"
	CoreTypes.Add "Vector"

	CoreTypes.Add "Bit"
	CoreTypes.Add "Digit"
	CoreTypes.Add "Sign"

	CoreTypes.Add "NameSpace"
	CoreTypes.Add "GenericName"
	CoreTypes.Add "LocalName"
	CoreTypes.Add "ScopedName"
	CoreTypes.Add "TypeName"
	CoreTypes.Add "MemberName"

	CoreTypes.Add "Any"

	CoreTypes.Add "Record"
	CoreTypes.Add "RecordType"
	CoreTypes.Add "Field"
	CoreTypes.Add "FieldType"
	
	'iso 19103:2015 Annex-C types
	ExtensionTypes.Add "LanguageString"
	
	ExtensionTypes.Add "Anchor"
	ExtensionTypes.Add "FileName"
	ExtensionTypes.Add "MediaType"
	ExtensionTypes.Add "URI"
	
	ExtensionTypes.Add "UnitOfMeasure"
	ExtensionTypes.Add "UomArea"
	ExtensionTypes.Add "UomLenght"
	ExtensionTypes.Add "UomAngle"
	ExtensionTypes.Add "UomAcceleration"
	ExtensionTypes.Add "UomAngularAcceleration"
	ExtensionTypes.Add "UomAngularSpeed"
	ExtensionTypes.Add "UomSpeed"
	ExtensionTypes.Add "UomCurrency"
	ExtensionTypes.Add "UomVolume"
	ExtensionTypes.Add "UomTime"
	ExtensionTypes.Add "UomScale"
	ExtensionTypes.Add "UomWeight"
	ExtensionTypes.Add "UomVelocity"

	ExtensionTypes.Add "Measure"
	ExtensionTypes.Add "Length"
	ExtensionTypes.Add "Distance"
	ExtensionTypes.Add "Speed"
	ExtensionTypes.Add "Angle"
	ExtensionTypes.Add "Scale"
	ExtensionTypes.Add "TimeMeasure"
	ExtensionTypes.Add "Area"
	ExtensionTypes.Add "Volume"
	ExtensionTypes.Add "Currency"
	ExtensionTypes.Add "Weight"
	ExtensionTypes.Add "AngularSpeed"
	ExtensionTypes.Add "DirectedMeasure"
	ExtensionTypes.Add "Velocity"
	ExtensionTypes.Add "AngularVelocity"
	ExtensionTypes.Add "Acceleration"
	ExtensionTypes.Add "AngularAcceleration"
	
	'well known and often used spatial types from iso 19107:2003
	ProfileTypes.Add "DirectPosition"
	ProfileTypes.Add "GM_Object"
	ProfileTypes.Add "GM_Primitive"
	ProfileTypes.Add "GM_Complex"
	ProfileTypes.Add "GM_Aggregate"
	ProfileTypes.Add "GM_Point"
	ProfileTypes.Add "GM_Curve"
	ProfileTypes.Add "GM_Surface"
	ProfileTypes.Add "GM_Solid"
	ProfileTypes.Add "GM_MultiPoint"
	ProfileTypes.Add "GM_MultiCurve"
	ProfileTypes.Add "GM_MultiSurface"
	ProfileTypes.Add "GM_MultiSolid"
	ProfileTypes.Add "GM_CompositePoint"
	ProfileTypes.Add "GM_CompositeCurve"
	ProfileTypes.Add "GM_CompositeSurface"
	ProfileTypes.Add "GM_CompositeSolid"
	ProfileTypes.Add "TP_Object"
	'ProfileTypes.Add "TP_Primitive"
	ProfileTypes.Add "TP_Complex"
	ProfileTypes.Add "TP_Node"
	ProfileTypes.Add "TP_Edge"
	ProfileTypes.Add "TP_Face"
	ProfileTypes.Add "TP_Solid"
	ProfileTypes.Add "TP_DirectedNode"
	ProfileTypes.Add "TP_DirectedEdge"
	ProfileTypes.Add "TP_DirectedFace"
	ProfileTypes.Add "TP_DirectedSolid"
	ProfileTypes.Add "GM_OrientableCurve"
	ProfileTypes.Add "GM_OrientableSurface"
	ProfileTypes.Add "GM_PolyhedralSurface"
	ProfileTypes.Add "GM_triangulatedSurface"
	ProfileTypes.Add "GM_Tin"

	'well known and often used coverage types from iso 19123:2007
	ProfileTypes.Add "CV_Coverage"
	ProfileTypes.Add "CV_DiscreteCoverage"
	ProfileTypes.Add "CV_DiscretePointCoverage"
	ProfileTypes.Add "CV_DiscreteGridPointCoverage"
	ProfileTypes.Add "CV_DiscreteCurveCoverage"
	ProfileTypes.Add "CV_DiscreteSurfaceCoverage"
	ProfileTypes.Add "CV_DiscreteSolidCoverage"
	ProfileTypes.Add "CV_ContinousCoverage"
	ProfileTypes.Add "CV_ThiessenPolygonCoverage"
	'ExtensionTypes.Add "CV_ContinousQuadrilateralGridCoverageCoverage"
	ProfileTypes.Add "CV_ContinousQuadrilateralGridCoverage"
	ProfileTypes.Add "CV_HexagonalGridCoverage"
	ProfileTypes.Add "CV_TINCoverage"
	ProfileTypes.Add "CV_SegmentedCurveCoverage"

	'well known and often used temporal types from iso 19108:2006/2002?
	ProfileTypes.Add "TM_Instant"
	ProfileTypes.Add "TM_Period"
	ProfileTypes.Add "TM_Node"
	ProfileTypes.Add "TM_Edge"
	ProfileTypes.Add "TM_TopologicalComplex"
	
	'well known and often used observation related types from OM_Observation in iso 19156:2011
	ProfileTypes.Add "TM_Object"
	ProfileTypes.Add "DQ_Element"
	ProfileTypes.Add "NamedValue"
	
	'well known and often used quality element types from iso 19157:2013
	ProfileTypes.Add "DQ_AbsoluteExternalPositionalAccurracy"
	ProfileTypes.Add "DQ_RelativeInternalPositionalAccuracy"
	ProfileTypes.Add "DQ_AccuracyOfATimeMeasurement"
	ProfileTypes.Add "DQ_TemporalConsistency"
	ProfileTypes.Add "DQ_TemporalValidity"
	ProfileTypes.Add "DQ_ThematicClassificationCorrectness"
	ProfileTypes.Add "DQ_NonQuantitativeAttributeCorrectness"
	ProfileTypes.Add "DQ_QuanatitativeAttributeAccuracy"

	'well known and often used metadata element types from iso 19115-1:200x and iso 19139:2x00x
	ProfileTypes.Add "PT_FreeText"
	ProfileTypes.Add "LocalisedCharacterString"
	ProfileTypes.Add "MD_Resolution"
	'ProfileTypes.Add "CI_Citation"
	'ProfileTypes.Add "CI_Date"

	'other less known Norwegian geometry types
	ProfileTypes.Add "Punkt"
	ProfileTypes.Add "Kurve"
	ProfileTypes.Add "Flate"
	ProfileTypes.Add "Sverm"


end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------



'------------------------------------------------------------START-------------------------------------------------------------------------------------------
'Function name: scriptBreakingStructuresInModel
'Author: 		Åsmund Tjora
'Date: 			20170511 
'Purpose: 		Check that the model does not contain structures that will break script operations (e.g. cause infinite loops)
'Parameter: 	the package where the script runs
'Return value:	false if no script-breaking structures in model are found, true if parts of the model may break the script.
'Sub functions and subs:	inHeritanceLoop, inheritanceLoopCheck
function scriptBreakingStructuresInModel(thePackage)
	dim retVal
	retVal=false
	dim currentElement as EA.Element
	dim elements as EA.Collection
	
	'Package Dependency Loop Check
	set currentElement = thePackage.Element
'	Note:  Dependency loops will not cause script to hang
'	retVal=retVal or dependencyLoop(currentElement)
	
	'Inheritance Loop Check
	set elements = thePackage.elements
	dim i
	for i=0 to elements.Count-1
		set currentElement = elements.GetAt(i)
		if(currentElement.Type="Class") then
			retVal=retVal or inheritanceLoop(currentElement)
		end if
	next
	scriptBreakingStructuresInModel = retVal
end function

'Function name: dependencyLoop
'Author: 		Åsmund Tjora
'Date: 			20170511 
'Purpose: 		Check that dependency structure does not form loops.  Return true if no loops are found, return false if loops are found
'Parameter: 	Package element where check originates
'Return value:	false if no loops are found, true if loops are found.
function dependencyLoop(thePackageElement)
	dim retVal
	dim checkedPackagesList
	set checkedPackagesList = CreateObject("System.Collections.ArrayList")
	retVal=dependencyLoopCheck(thePackageElement, checkedPackagesList)
	if retVal then
		Session.Output("Error:  The dependency structure originating in [«" & thePackageElement.StereoType & "» " & thePackageElement.name & "] contains dependency loops [/req/uml/integration]")
		Session.Output("          See the list above for the packages that are part of a loop.")
		Session.Output("          Ignore this error for dependencies between packages outside the control of the current project.")
		globalErrorCounter = globalErrorCounter+1
	end if
	dependencyLoop = retVal
end function

function dependencyLoopCheck(thePackageElement, dependantCheckedPackagesList)
	dim retVal
	dim localRetVal
	dim dependee as EA.Element
	dim connector as EA.Connector
	
	' Generate a copy of the input list.  
	' The operations done on the list should not be visible by the dependant in order to avoid false positive when there are common dependees.
	dim checkedPackagesList
	set checkedPackagesList = CreateObject("System.Collections.ArrayList")
	dim ElementID
	for each ElementID in dependantCheckedPackagesList
		checkedPackagesList.Add(ElementID)
	next
	
	retVal=false
	checkedPackagesList.Add(thePackageElement.ElementID)
	for each connector in thePackageElement.Connectors
		localRetVal=false
		if connector.Type="Usage" or connector.Type="Package" or connector.Type="Dependency" then
			if thePackageElement.ElementID = connector.ClientID then
				set dependee = Repository.GetElementByID(connector.SupplierID)
				dim checkedPackageID
				for each checkedPackageID in checkedPackagesList
					if checkedPackageID = dependee.ElementID then localRetVal=true
				next
				if localRetVal then 
					Session.Output("         Package [«" & dependee.Stereotype & "» " & dependee.Name & "] is part of a dependency loop")
				else
					localRetVal=dependencyLoopCheck(dependee, checkedPackagesList)
				end if
				retVal=retVal or localRetVal
			end if
		end if
	next
	
	dependencyLoopCheck=retVal
end function


'Function name: inheritanceLoop
'Author: 		Åsmund Tjora
'Date: 			20170221 
'Purpose: 		Check that inheritance structure does not form loops.  Return true if no loops are found, return false if loops are found
'Parameter: 	Class element where check originates
'Return value:	false if no loops are found, true if loops are found.
function inheritanceLoop(theClass)
	dim retVal
	dim checkedClassesList
	set checkedClassesList = CreateObject("System.Collections.ArrayList")
	retVal=inheritanceLoopCheck(theClass, checkedClassesList)	
	if retVal then
		Session.Output("Error: Class hierarchy originating in [«" & theClass.Stereotype & "» "& theClass.Name & "] contains inheritance loops.")
	end if
	inheritanceLoop = retVal
end function

'Function name:	inheritanceLoopCheck
'Author:		Åsmund Tjora
'Date:			20170221
'Purpose		Internal workings of function inhertianceLoop.  Register the class ID, compare list of ID's with superclass ID, recursively call itself for superclass.  
'				Return "true" if class already has been registered (i.e. is a superclass of itself) 

function inheritanceLoopCheck(theClass, subCheckedClassesList)
	dim retVal
	dim superClass as EA.Element
	dim connector as EA.Connector

	' Generate a copy of the input list.  
	'The operations done on the list should not be visible by the subclass in order to avoid false positive at multiple inheritance
	dim checkedClassesList
	set checkedClassesList = CreateObject("System.Collections.ArrayList")
	dim ElementID
	for each ElementID in subCheckedClassesList
		checkedClassesList.Add(ElementID)
	next

	retVal=false
	checkedClassesList.Add(theClass.ElementID)	
	for each connector in theClass.Connectors
		if connector.Type = "Generalization" then
			if theClass.ElementID = connector.ClientID then
				set superClass = Repository.GetElementByID(connector.SupplierID)
				dim checkedClassID
				for each checkedClassID in checkedClassesList
					if checkedClassID = superClass.ElementID then retVal = true
				next
				if retVal then 
					Session.Output("Error: Class [«" & superClass.Stereotype & "» " & superClass.Name & "] is a generalization of itself")
				else
					retVal=inheritanceLoopCheck(superClass, checkedClassesList)
				end if
			end if
		end if
	next
	
	inheritanceLoopCheck = retVal
end function

'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


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


'global variables 
dim globalLogLevelIsWarning 'boolean variable indicating if warning log level has been choosen or not
globalLogLevelIsWarning = true 'default setting for warning log level is true
 
dim startClass as EA.Element  'the class which is the starting point for searching for multiple inheritance in the findMultipleInheritance subroutine 
dim loopCounterMultipleInheritance 'integer value counting number of loops while searching for multiple inheritance
dim foundHoveddiagram 'boolean to check if a diagram named Hoveddiagram is found. If found, foundHoveddiagram = true  
foundHoveddiagram = false 
dim numberOfHoveddiagram 'number of diagrams named Hoveddiagram
numberOfHoveddiagram = 0
dim numberOfHoveddiagramWithAdditionalInformationInTheName 'number of diagrams with a name starting with Hoveddiagram and including additional characters  
numberOfHoveddiagramWithAdditionalInformationInTheName = 0
dim globalErrorCounter 'counter for number of errors 
globalErrorCounter = 0 
dim globalWarningCounter
globalWarningCounter = 0
dim startPackageName
dim debug
debug = false

'List of well known type names defined in iso 19109:2015
dim ProfileTypes
'List of well known extension type names defined in iso 19103:2015
dim ExtensionTypes
'List of well known core type names defined in iso 19103:2015
dim CoreTypes

OnProjectBrowserScript
