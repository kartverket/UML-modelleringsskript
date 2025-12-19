!INC listAdocFraModell

''  Include-setningen over identifiserer pakka, kjører ut en full modellrapport,
''  og tilbyr nødvendige subrutiner og funksjoner til denne modulen slik at 
''  den kan generere en utlisting av de realiserte objekttypene


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


dim visSosiFormatRealisering '' viser sosi-format i realiseringsvedlegget 

''visSosiFormatRealisering = true
  visSosiFormatRealisering = false

if visSosiFormatRealisering then

	dim avledSosiNavnFraEgenskap '' avleder SOSI-navn fra egenskapsnavn
	dim avledSosiTypeFraDatatype '' avleder SOSI-type fra basis egenskapstype
	dim avledSosiNavnFraDatatype '' hent SOSI-navn fra datatypens SOSI-navn

	visSosiFormatRealisering = true	
	avledSosiTypeFraDatatype = true
''	avledSosiNavnFraEgenskap = true  '' BRYTER Realisering i SOSI-format 5.0
''	avledSosiNavnFraDatatype = true  '' BRYTER Realisering i SOSI-format 5.0

	set utfil = tomTekstfil( "SOSIformatRealisering.adoc")
	
else

	set utfil = tomTekstfil( "RealiserteObjekttyper.adoc")

end if

Session.Output "// Realiserte objekttyper i " + valgtPakke.element.name
Session.Output "// Start of UML-model"		

call visPakkasRealiserteObjekttyper( valgtPakke)

Session.Output "// End of UML-model"



''  ----------------------------------------------------------------------------

Sub visPakkasRealiserteObjekttyper( pakke)

	dim pakkenivaa : pakkenivaa = 3
	dim pakketittel : pakketittel = "Pakke : " + pakke.Name

	dim elementnivaa : elementnivaa = pakkenivaa + 1
	Dim element As EA.Element
	For each element in pakke.Elements
		dim ster : ster = Ucase(element.Stereotype)
		If element.Type <> "Class" then  	'' hopp over
		elseif element.Abstract = 1 then	'' hopp over abstrakte klasser
		elseif (ster = "FEATURETYPE" or ster = "") Then

			dim elementTittel  '' navnet på featuretype
			elementTittel = "Objekttype: " + element.Name
			settInnTekst unummerertOverskrift( elementnivaa, elementTittel)

			if visSosiFormatRealisering then
				visSosiRealiseringAvObjekttype element, elementnivaa +1
			else
				visObjekttype element, elementnivaa +1
			end if	
		End if		
	Next

	''	Gå rekursivt gjennom alle underpakker
	dim delpakke as EA.Package
	for each delpakke in pakke.Packages
	
		Call visPakkasRealiserteObjekttyper( delpakke)
		
	next
	
end sub

''  ----------------------------------------------------------------------------

Sub visObjekttype( objektType, subnivaa) 

	dim tabellhode, tabellformat
	tabellformat = "20,20,10"
	tabellhode = array( "Navn", "Type", "Mult.")	
''	hode(0) = "Navn på egenskap/rolle"
	
	dim i
	for i = 0 to UBound(tabellhode)
		tabellhode(i) = bold(tabellhode(i))
	next

	dim attributtListe
	attributtListe = egenskaperOgRoller( objektType)
	if UBound(attributtListe) > 0 then	
		settInnTekst unummerertOverskrift( subnivaa, "Attributter og roller")	
		attributtListe(0) = tabellhode
		call SettInnSomTabell( attributtListe, tabellFormat, "")
	end if

end sub

''  ----------------------------------------------------------------------------

Sub visSosiRealiseringAvObjekttype( objektType, subnivaa) 

	dim sosiGeo : sosiGeo = listeAvSosiGeometrier( objektType)
	settInnTekst unummerertOverskrift( subnivaa, "Geometrityper")
	if isEmpty(sosiGeo) then sosiGeo = "Ingen SOSI-geometrier"
	settInnTekst sosiGeo


	dim avgrensesAv : avgrensesAv = avgrensingslinjer(objektType)
	if not isEmpty(avgrensesAv) then 
		settInnTekst unummerertOverskrift( subnivaa, "Avgrenses av")
		settInnTekst join( avgrensesAv, ", ")
	end if		

	dim avgrenser : avgrenser = flaterSomAvgrenses(objektType)
	if not isEmpty(avgrenser) then 
		settInnTekst unummerertOverskrift( subnivaa, "Avgrenser")
		settInnTekst join( avgrenser, ", ")
	end if	


	dim tabellhode, tabellformat, i
	tabellformat = "20,20,5,25,10"  
	tabellhode = array( "Navn", "Type", "Mult.", "SOSI-navn", "SOSI-type")	
	for i = 0 to UBound(tabellhode)
		tabellhode(i) = bold(tabellhode(i))
	next
	dim attributter : attributter = klassensAlleAttributter( objektType)
	if UBound(attributter) > 0 then	
		settInnTekst unummerertOverskrift( subnivaa, "Attributter")	
		attributter(0) = tabellhode
		call SettInnSomTabell( attributter, tabellFormat, "")
	end if


	tabellformat = "25,25,5,25"
	tabellhode = array( "Rollenavn", "Objekttype", "Mult.", "SOSI-navn")
	for i = 0 to UBound(tabellhode)
		tabellhode(i) = bold(tabellhode(i))
	next

	dim roller : roller = klassensAlleroller( objektType)
	if UBound(roller) > 0 then	
		settInnTekst unummerertOverskrift( subnivaa, "Roller")	
		roller(0) = tabellhode
		call SettInnSomTabell( roller, tabellFormat, "")
	end if

end sub

''  ----------------------------------------------------------------------------

function avgrensingslinjer( element)
''	Inneholder rekursivt kall til elementets supertype

	dim liste
	if antallSupertyper( element) = 1 then
	
		liste = avgrensingslinjer( superType(element))
		
	end if

	dim elementId : elementId = element.elementID
	dim con
	For Each con In element.Connectors
		dim targetID, target
		if elementId = con.SupplierID then
			targetID = con.ClientID
			set target = con.ClientEnd	
		elseif elementId = con.ClientID  then
			targetID = con.SupplierID
			set target = con.SupplierEnd
		end if
		
		'	Det er tre måter å avgøre om target er et avgrensingsobjekt:
		' SOSI 4.5:	conn.Stereotype = "Topo". Kan navigere til target
		' SOSI 5.0:	element.Constraints inneholder 'KanAvgrensesAv Target.Name'
		' FKB-praksis:	target.Role = "avgrensesAv"
		'
		'	Det er bare FKB-varianten som er lagt inn
		'
		if erSosiAvgrensingsrolle( target) and targetID > 0 then 
			liste = merge(liste, Repository.GetElementByID( targetID).Name )
		end if
	next
	
	avgrensingslinjer = liste
end function

''  ----------------------------------------------------------------------------

function flaterSomAvgrenses( element)
''	Inneholder rekursivt kall til elementets supertype

	dim liste
	
	if antallSupertyper( element) = 1 then

		liste = flaterSomAvgrenses( superType(element))
	
	end if
	
	dim elementId : elementId = element.elementID
	dim con
	For Each con In element.Connectors
		dim targetID, current
		if elementId = con.SupplierID then
			targetID = con.ClientID
			set current = con.SupplierEnd	
		elseif elementId = con.ClientID  then
			targetID = con.SupplierID
			set current = con.ClientEnd		
		end if
		
		'	Det er tre måter å avgøre om current er et avgrensingsobjekt:
		' SOSI 4.5:	conn.Stereotype = "Topo". Kan navigere til current
		' SOSI 5.0:	element.Constraints inneholder 'KanAvgrensesAv Current.Name'
		' FKB-praksis:	current.Role = "avgrensesAv"
		'
		'	Det er bare FKB-varianten som er lagt inn
		'

		if erSosiAvgrensingsrolle( current) and targetID > 0 then 
			liste = merge(liste, Repository.GetElementByID( targetID).Name)
		end if
	next
	
	flaterSomAvgrenses = liste
end function

''  ----------------------------------------------------------------------------

function erSosiAvgrensingsrolle( rolle)
	'	Det er tre måter å avgøre om rolle er et avgrensingsobjekt:
	' SOSI 4.5:	conn.Stereotype = "Topo". Kan navigere til rolle
	' SOSI 5.0:	element.Constraints inneholder 'KanAvgrensesAv Target.Name'
	' FKB-praksis:	rolle.Role = "avgrensesAv"
	'
	'	Det er bare FKB-varianten som er sjekkes her
	'
	dim rollenavn  
	if not isEmpty(rolle) then rollenavn = LCase(rolle.Role) 

	dim erAvgrensingsrolle
	erAvgrensingsrolle = ( inStr( rollenavn, "avgrensesav") <> 0 )
	
	erSosiAvgrensingsrolle = erAvgrensingsrolle and visSosiFormatRealisering

end function

		''  ----------------------------------------------------------------------------

function egenskaperOgRoller( featureType)

	''	finn først featuretypens egen egenskaper og roller
	''	deretter fra supertypen, som legges først i lista
	
	dim egensk, roller
	egensk = klassensEgneAttributter( featureType, "")
	roller = klassensEgneRoller( featureType)
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
		super = egenskaperOgRoller( superType(featureType))
		egenskaperOgRoller = merge( super, egensk)
	end if

end function

''  ----------------------------------------------------------------------------

function klassensAlleAttributter( featureType)
''	Inneholder rekursivt kall til elementets supertype

	dim egensk
	egensk = klassensEgneAttributter( featureType, "")

	if antallSupertyper( featureType) = 0 then
		''	Denne featuretypen har ingen supertyper
		dim plassTilTabellhode()
		redim plassTilTabellhode(0)   '' første linje i lista holdes tom
		klassensAlleAttributter = merge( plassTilTabellhode, egensk)  

''	elseif antallSupertyper(featureType) > 1 then
''		'' FEILSITUASJON

	else  '' harSupertype(featureType) = true
		'' egenskaper og roller fra supertypen legges først i lista
		dim super
		super = klassensAlleAttributter( superType(featureType))
		klassensAlleAttributter = merge( super, egensk)
	end if

end function

''  ----------------------------------------------------------------------------

function kortAttributtBeskrivelse(att, egenskapsgruppe, datatype)

	dim egNavn : egNavn = egenskapsgruppe & att.Name 
	dim sNavn :	sNavn = sosiNavn(att, egenskapsgruppe)
	
	if sNavn =  "" and avledSosiNavnFraDatatype then
		sNavn = sosiNavn(datatype, egenskapsgruppe)
		if sNavn <>  "" then sNavn = kursiv(sNavn)
	end if
	
	dim mult : mult = "[" & att.LowerBound & ".." & att.UpperBound & "]"
	
	dim dtyp, sosiDatatype
	if isnull(datatype) then
		dtyp = att.type
		sosiDatatype = SOSItypeFraBasistype(dtyp)
	else
		dtyp = stereotypeNavn( datatype)
		
		if isDataType( dataType) then  
			'' marker sammensatt datatype
			sosiDatatype = "*"
		elseif iscodelist( datatype) then 
			'' hent SOSIdatatype fra kodelista
			sosiDatatype = sosiDatatypeFraTagger(dataType)
			
			if sosiDatatype = "" then sosiDatatype = "T"
		end if
	end if

	dim sType :	stype = sosiDatatypeFraTagger( att)
	if sType = "" then stype = sosiDatatype
	
	if not visSosiFormatRealisering then
	
		kortAttributtBeskrivelse = array(egNavn, dtyp, mult)
		
	elseif sosiGeometritype( att) = "" then

		kortAttributtBeskrivelse = array( egNavn, dtyp, mult, sNavn, sType) 
	
	end if

end function

''  ----------------------------------------------------------------------------

function klassensEgneAttributter (element, byVal egenskapsgruppe)
''	Inneholder rekursivt kall til elementets sammensatte datatyper

	if egenskapsgruppe <> "" then egenskapsgruppe = egenskapsgruppe + "."

	dim liste
	Dim att As EA.Attribute
	
	for each att in element.Attributes
		dim datatype
		if att.ClassifierID <> 0 then
			set datatype = Repository.GetElementByID(att.ClassifierID)
		else
			datatype = null
		end if
		
		dim attrib
		attrib = kortAttributtBeskrivelse(att, egenskapsgruppe, datatype)
		
		liste = merge(liste, array(attrib))

		'' ta med underegenskapene dersom attrib er en sammensatt datatype
		if isNull(attrib) then
		elseif isnull(datatype) then
		else
			dim styp : styp =  LCase(datatype.Stereotype)
			if listeInneholderElement( array( "datatype", "union"), styp) then
				dim attributter, egNavn 
				egNavn = attrib(0)	
				attributter = klassensEgneAttributter( datatype, egNavn)
				
				liste = merge( liste, attributter)
			end if
		end if
	next
	
	klassensEgneAttributter = liste
	
end function

''  ----------------------------------------------------------------------------

function klassensAlleRoller( featureType)
''	Inneholder rekursivt kall til elementets supertype

	dim egensk
	egensk = klassensEgneRoller( featureType)

	if antallSupertyper( featureType) = 0 then
		''	Denne featuretypen har ingen supertyper
		dim plassTilTabellhode()
		redim plassTilTabellhode(0)   '' første linje i lista holdes tom
		klassensAlleRoller = merge( plassTilTabellhode, egensk)  

''	elseif antallSupertyper(featureType) > 1 then
''		'' FEILSITUASJON

	else  '' harSupertype(featureType) = true
		'' egenskaper og roller fra supertypen legges først i lista
		dim super
		super = klassensAlleRoller( superType(featureType))
		klassensAlleRoller = merge( super, egensk)
	end if

end function

''  ----------------------------------------------------------------------------

function klassensEgneRoller( featureType)

	dim rollesamling  	'' array av rolle-arryer
''	rollesamling = alleRoller( featureType)
	rollesamling = sorterteRoller( featureType)
	
	if isEmpty(rollesamling) then 	EXIT function
	
	dim liste(), radNr 
	radNr = 0
	redim liste(featureType.Connectors.Count)

	dim rolle   		
	for each rolle in rollesamling 
		dim target, targetID
		targetID = rolle(4)
		set target = rolle(3)
		
		dim rollenavn, dtyp, mult		
		rollenavn = target.Role
		if targetID <> 0 then
			dtyp = stereotypenavn( Repository.GetElementByID(targetID) )
		end if
		mult = "[" & target.Cardinality & "]"

		if not visSosiFormatRealisering then 
			rollenavn = bold("Rolle: ") + rollenavn
			liste(radNr) = array(rollenavn, dtyp, mult)
		elseif not erSosiAvgrensingsrolle( target) then
''			dim sNavn : sNavn = 						sosiRollenavn(target)
			liste(radNr) = array(rollenavn, dtyp, mult, sosiRollenavn(target))
		end if

		radNr = radNr +1
	next
	
	redim preserve liste(radNr-1)
	
	klassensEgneRoller = liste
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

function sosiRollenavn(rolle)  ''sosiRollenavn

	dim sNavn 
	sNavn = taggedValueFraRolle(rolle, "SOSI_navn") 
	
	if sNavn <> "" then sosiRollenavn = ".." & sNavn

end function

''  ----------------------------------------------------------------------------

function sosiNavn( element, egenskapsgruppe)

	dim sNavn : sNavn = taggedValueFraElement(element, "SOSI_navn")
	
	if sNavn = "" then 
		if avledSosiNavnFraEgenskap then
			sNavn = Ucase(element.Name)
		else
			EXIT function
		end if 
	end if
	
	dim antallPrikker : antallPrikker = 2
	if egenskapsgruppe <> "" then
		antallPrikker = antallPrikker + Ubound(split(egenskapsgruppe, ".") )
	end if

	sosiNavn = string( antallPrikker, ".") + sNavn
end function

''  ----------------------------------------------------------------------------

function SOSItypeFraBasistype(basistype)
	'' Avlede SOSI basistype ihht Tabell 8.1 i "Realisering i SOSI-format 5.0"
	
	dim sosiBasistype
	if     basistype = "Integer" then 
		sosiBasistype = "H"

	elseif basistype = "Real" then 
		sosiBasistype = "D"
		
	elseif basistype = "CharacterString" then 
		sosiBasistype = "T"
		
	elseif basistype = "DateTime" then 
		sosiBasistype = "DATOTID"
		
	elseif basistype = "Date" then 
		sosiBasistype = "DATO"
		
	elseif basistype = "Boolean" then 
		sosiBasistype = "BOOLSK"
		
	end if
	
	if sosiBasistype <> "" and avledSosiTypeFraDatatype then 
		SOSItypeFraBasistype = sosiBasistype
''	else
''		SOSItypeFraBasistype = basistype
	end if

end function

''  ----------------------------------------------------------------------------

function sosiDatatypeFraTagger( element)

	dim sosiType, sosiLengde
	sosiType   = taggedValueFraElement(element, "SOSI_datatype")
	sosiLengde = taggedValueFraElement(element, "SOSI_lengde")

	if sosiType <> "" then
		sosiDatatypeFraTagger = sosiType & sosiLengde
	end if

end function

''  ----------------------------------------------------------------------------

function listeAvSosiGeometrier( element)

	dim sosiGeo
	sosiGeo = sosiGeometrier( element)
				
	if isEmpty(sosiGeo) then
''		settInnTekst bold("Geometrityper:  I N G E N" ) 
	elseif UBound(sosiGeo) = 0 then 
		listeAvSosiGeometrier = sosiGeo(0)
	else
		listeAvSosiGeometrier = join( sosiGeo , ", ")
	end if
	
''	settInnTekst avsnittSkille( )
end function
			

''  ----------------------------------------------------------------------------

function sosiGeometrier(featureType)
''	Inneholder rekursivt kall til elementets supertype

	dim att, liste, sGeo
	for each att in featureType.Attributes

		sGeo = sosiGeometritype( att)

		if sGeo <> "" then 	liste = merge(liste, sGeo )

	next
	
	if antallSupertyper( featureType) = 1 then
		'' egenskaper fra supertypen legges først i lista
		dim super
		super = sosiGeometrier( superType(featureType))
		sosiGeometrier = merge( super, liste)
	else
		sosiGeometrier = liste
	end if

end function

''  ----------------------------------------------------------------------------

function sosiGeometritype(element)

	dim sgtype
	dim gtype : gtype = element.Type

	if gtype = "Punkt" or gtype = "Kurve" or gtype = "Flate" then
	'' if listeInneholder(array("Punkt", "Kurve", "Flate"), gtype) then
		sgtype = UCase( gtype)

	'fra Ralisering i SOSI-format versjon 5.0 tabell 8.2:
	elseif gtype = "GM_Point" then
		sgtype = "PUNKT"
		
	elseif gtype = "GM_MultiPoint" then
		sgtype = "SVERM"
		
	elseif gtype = "GM_Curve" or gtype = "GM_CompositeCurve" then
		sgtype = "KURVE"
		
	elseif gtype = "GM_Surface" or gtype = "GM_CompositeSurface" then
		sgtype = "FLATE"
		
	'fra "etablert praksis"
	elseif gtype = "GM_Object" or gtype = "GM_Primitive" then
		sgtype = "OBJEKT"
		
	end if
	
	if sgtype <> "" then sosiGeometritype = sgtype
	
end function
