Option Explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: listAdocFraRegistreringsinstruks
' Original Author: Tore Johnsen, Åsmund Tjora
' Purpose: Generate documentation in AsciiDoc syntax
' Date: 08.04.2021
' Version: 0.1ish
'
' Versjon 0.17 2021-10-13 Kent: la inn mer rekusrivitet så egenskaper i supertyper kan presiseres (AtributesEX virker ikke), resatte et flagg - linje 1035
' Versjon 0.16 2021-10-13 tilpasninger av Geir Myhr Øien, lagt inn lenke til kodelister som er angitt tagged valuse "defaultcodespace"#, tagget med "v0.16"
' Versjon 0.15 2021-10-12 tilpasninger av Geir Myhr Øien, utlisting av taggen "FKB_minstestørrelse_*", tagget med "v0.15"
' Versjon 0.14 2021-10-07 tilpasninger av Geir Myhr Øien, tagget med "v0.14"
' Version 0.13 2021-10-07 endra litt på rekkefølgen mellom blokkene, starta restrukturering av koden
' Version 0.12 2021-10-06 spesialhandtering av restriksjoner med navn som starter med _presiseringAvKoder_
' Version 0.11 2021-10-06 figurer på kodelistekoder kopiert fra skript listAdocFraModell
' Version 0.10 2021-10-05 feilretting rundt image:: og Alt=
' Version 0.9 2021-10-04 hovedpakka ikke ut, figurer som i prodspek, underline i tV ut som blank, tilleggsdefinisjon bold etc, roller
' Version 0.8 2021-09-20 smårettinger
' Version 0.7 2021-09-10 diagrammer med "utelat" i navnet skrives ikke ut
' Version 0.6 2021-09-09 rettet en feil slik at nå kun egenskaper fra supertyper skrives ut
' Version 0.5 2021-06-30 vise egenskapsnavn foran datatypeegenskapsnavn (informasjon.navnerom) vise stereotypenavn foran datatyper og kodelister
' Version 0.4 2021-06-14/23 Nøste utover i datatypene og vise alle egenskaper (og roller TBD) Endra navn
' Version 0.3 2021-05-25/31 Collects guidance parameters from both tagged values and from special attributes 
' Version 0.2x: 2021-06-01 Kent Jonsrud: retta bildesti til app_img
'
' Version: 0.2
' Date: 2021-04-16 Kent Jonsrud:
' Use Case: Guidance for photogrammetric data collection
' Actor is a person digitizing from aeriel photos
' Design Goal: compact summary of all guiding requirements
'
' TBD: navnekræsj mellom diagrammer?
' TBD: feature type: table of tags from supertypes as FKB-standard A/B/C/D
' - feature type: additional definition (Minstestørrelse etc.)
' TBD: feature type: registreringsmetode
' TBD: feature type: tilleggsbeskrivelse
' TBD: feature type: forhold til andre objekttyper
' TBD: feature type: ?
' - feature attributes: table of all inherited attributes - name, SOSI_navn, type, length(?), FKB-standard A/B/C/D
' - one picture (and related picture text?) from each supertype
' - several pictures and related picture text from subtype

' TBD: common tables with attributes ??? really?

' TBD: common table of quality requirements classes for all feature types
' TBD: common table of quality requirements on each feature type
'	

' TBD: common table 
' TBD: opprydding !!!
'
Dim imgfolder, imgparent, parentimg, files 'v0.14
Dim diagCounter
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
			diagCounter = 0
			Dim thePackage As EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			imgfolder = "app_img"
			Set imgFSO=CreateObject("Scripting.FileSystemObject")
			imgparent = imgFSO.GetParentFolderName(Repository.ConnectionString())  & "\" & imgfolder
			if false then				
				Session.Output(" DEBUG.")
				Session.Output(" imgfolder: " & imgfolder  & "...")
				Session.Output(" imgFSO.GetAbsolutePathName: " & imgFSO.GetAbsolutePathName("./")  & "...")
				Session.Output(" Repository.ConnectionString: " & Repository.ConnectionString() & "...")
				Session.Output(" imgFSO.GetBaseName: " & imgFSO.GetBaseName(Repository.ConnectionString())  & "...")
				Session.Output(" imgFSO.GetParentFolderName: " & imgFSO.GetParentFolderName(Repository.ConnectionString())  & "...")
				Session.Output(" imgparent: " & imgparent  & "...")
			end if
			if not imgFSO.FolderExists(imgparent) then
				imgFSO.CreateFolder imgparent
			end if

			Call ListAsciiDoc(thePackage)
			Session.Output("// End of Registreringsinstruks UML-model")
			Set imgFSO = Nothing
        Case Else
            ' Error message
            Session.Prompt "This script does not support items of this type.", promptOK

    End Select

End Sub


Sub ListAsciiDoc(thePackage)

	Dim element As EA.Element
	dim tag as EA.TaggedValue
	Dim diag As EA.Diagram
	Dim projectclass As EA.Project
	set projectclass = Repository.GetProjectInterface()

	if InStr(LCase(thePackage.Name),"fotogrammetrisk") = 0 then
	
		Session.Output("=== "&thePackage.Name&"")
		if thePackage.Notes <> "" then Session.Output("*Pakkens definisjon:* "&thePackage.Notes&"")

		if thePackage.element.TaggedValues.Count > 0 then
			Session.Output(" ")	
			Session.Output("===== Tagged Values")
			Session.Output("[cols=""20,80""]")
			Session.Output("|===")
			for each tag in thePackage.element.TaggedValues
				if tag.Value <> "" then	
					if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" and LCase(tag.Name) <> "sosi_bildeavmodellelement" then
					'	Session.Output("|Tag: "&tag.Name&"")
					'	Session.Output("|Verdi: "&tag.Value&"")
						Session.Output("|"&tag.Name&"")
						Session.Output("|"&tag.Value&"")
						Session.Output(" ")			
					end if
				end if
			next

			Session.Output("|===")
		end if

		for each tag in thePackage.element.TaggedValues
			if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
				diagCounter = diagCounter + 1
			'	Session.Output("[caption=""Figur "&diagCounter&": "",title="&tag.Name&"]")
				Session.Output("[caption=""Figur  "",title="&tag.Name&"]")
			'	Session.Output("image::"&tag.Value&".png["&ThePackage.Name"."&tag.Name&"]")
				Session.Output("image::"&tag.Value&".png["&tag.Value&", Alt=""Bilde av pakke: "&thePackage.Name&"""]")
			end if
			'if EA-document then
			'	figure + figure text
			'end if
		next

	end if

'-----------------Diagram-----------------

	For Each diag In thePackage.Diagrams
		if InStr(LCase(diag.Name),"utelat") = 0 then
			diagCounter = diagCounter + 1
			Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, imgparent & "\" & diag.Name & ".png", 1)
		'	Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, "" & diag.Name&".png", 1)
			Repository.CloseDiagram(diag.DiagramID)
		'	Session.Output("[caption=""Figur "&diagCounter&": "",title="&diag.Name&"]")
			Session.Output("[caption=""Figur  "",title="&diag.Name&"]")
		'	Session.Output("image::"&imgfolder&"\"&diag.Name&".png["&diag.Name&"]")
			Session.Output("image::"&diag.Name&".png["&diag.Name&", Alt=""Diagramm: "&diag.Name&"""]")
		end if
	Next

	For each element in thePackage.Elements
		If element.Type = "Class" and (Ucase(element.Stereotype) = "FEATURETYPE" or Ucase(element.Stereotype) = "") Then
	'	If Ucase(element.Stereotype) = "FEATURETYPE" or Ucase(element.Stereotype) = "" Then
			Call ObjektOgDatatyper(element)
		End if
	Next
		
	For each element in thePackage.Elements
		If Ucase(element.Stereotype) = "DATATYPE" Then
			Call ObjektOgDatatyper(element)
		End if
	Next

	For each element in thePackage.Elements
		If Ucase(element.Stereotype) = "UNION" Then
			Call ObjektOgDatatyper(element)
		End if
	Next

if false then
	For each element in thePackage.Elements
		If Ucase(element.Stereotype) = "CODELIST" Then
			Call Kodelister(element)
		End if
		If Ucase(element.Stereotype) = "ENUMERATION" Then
			Call Kodelister(element)
		End if
		If element.Type = "Enumeration" Then
			Call Kodelister(element)
		End if
	Next
end if 'false
		
	dim pack as EA.Package
	for each pack in thePackage.Packages
		Call ListAsciiDoc(pack)
	next


end sub

'-----------------ObjektOgDatatyper-----------------
Sub ObjektOgDatatyper(element)
	Dim att As EA.Attribute
	dim tag as EA.TaggedValue
	Dim con As EA.Connector
	Dim fil As EA.File
	Dim supplier As EA.Element
	Dim client As EA.Element
	Dim association
	Dim aggregation
	association = False
	Dim generalizations
	Dim numberSpecializations ' tar også med antall realiseringer her
	Dim textVar, addnotes, punktum
	dim externalPackage

	parentimg = ""
	files = "" 'v0.14
	Session.Output(" ")
	Session.Output("<<<")
	Session.Output("'''")
	Session.Output(" ")
	
	Session.Output("[["&LCase(element.Name)&"]]")
	Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
	For Each con In element.Connectors
		set supplier = Repository.GetElementByID(con.SupplierID)
		If con.Type = "Generalization" And supplier.ElementID <> element.ElementID Then
			Session.Output("*Definisjon fra FKB produktspesifikasjon*: "&supplier.Notes&"")
			Session.Output(" ")
			numberSpecializations = numberSpecializations + 1
			
			parentimg = getTaggedValue(supplier,"SOSI_bildeAvModellelement")
			
		End If
	Next
	
	call klassebilder(element,parentimg,files)  'v0.14
	
	if element.Notes <> "" then
'		Session.Output("*Tilleggsinformasjon for fotogrammetrisk registrering:* "&element.Notes&"")
		Session.Output("===== Tilleggsinformasjon for fotogrammetrisk registrering")
		Session.Output(""&element.Notes&"")
		Session.Output(" ")
	end if
	
	files = "Ja"   'v0.14
	call klassebilder(element,parentimg,files)  'v0.14
	files = ""  'v0.14

	numberSpecializations = 0
	For Each con In element.Connectors
		set supplier = Repository.GetElementByID(con.SupplierID)
		If con.Type = "Generalization" And supplier.ElementID <> element.ElementID Then
	'		Session.Output("*Supertype:* «" & supplier.Stereotype&"» "&supplier.Name&"")
	'		Session.Output(" ")
			numberSpecializations = numberSpecializations + 1
		End If
	Next

	For Each con In element.Connectors  
''realiseringer.  
''Må forbedres i framtidige versjoner dersom denne skal med 
''- full sti (opp til applicationSchema eller øverste pakke under "Model") til pakke som inneholder klassen som realiseres
		set supplier = Repository.GetElementByID(con.SupplierID)
		If con.Type = "Realisation" And supplier.ElementID <> element.ElementID Then
			set externalPackage = Repository.GetPackageByID(supplier.PackageID)
			textVar=getPath(externalPackage)
	'		Session.Output("*Realisering av:* " & textVar &"::«" & supplier.Stereotype&"» "&supplier.Name)
	'		Session.Output(" ")
			numberSpecializations = numberSpecializations + 1
		end if
	next

'if element.Attributes.Count > 0 then
		Session.Output("===== Føringer")
		Session.Output("[cols=""25,75""]")
		Session.Output("|===")
		for each att in element.AttributesEx
		if att.name = "Registreringsmetode" or att.name = "Tilleggsbeskrivelse" or att.name = "Grunnrissreferanse" or att.name = "Høydereferanse" then
			Session.Output("|"&att.name&"")
			Session.Output("|"&getCleanDefinition(att.Notes)&"")
			Session.Output(" ")
		end if
		next
		for each tag in element.TaggedValues								
			if tag.Value <> "" then	
				if tag.Name = "Registreringsmetode" or tag.Name = "Tilleggsbeskrivelse" or tag.Name = "Grunnrissreferanse" or tag.Name = "Høydereferanse" or Mid(tag.Name,1,3) = "FKB" then
					call listFkbTag(tag.Name, tag.Value)
	if false then
					if tag.Name = "FKB-A" or tag.Name = "FKB_A" then
						call listFkbTag(tag.Name, tag.Value)
					else
						if tag.Name = "FKB-B" or tag.Name = "FKB_B" then
							call listFkbTag(tag.Name, tag.Value)
						else
							if tag.Name = "FKB-C" or tag.Name = "FKB_C" then
								call listFkbTag(tag.Name, tag.Value)
							else
								if tag.Name = "FKB-D" or tag.Name = "FKB_D" then
									call listFkbTag(tag.Name, tag.Value)
								else
									Session.Output("|"&tag.Name&"")
									Session.Output("|"&getCleanDefinition(tag.Value)&"")
									Session.Output(" ")	
								end if
							end if
						end if					
					end if					
	end if 'false
				end if
			end if
		next		
		
'		if element.Notes <> "" then
'			Session.Output("|Tilleggsdefinisjon")
'			Session.Output("|"&getCleanDefinition(element.Notes)&"")
'			Session.Output(" ")			
'		end if
		
		Session.Output("|===")
		
'		Session.Output("""P"" => Påkrevd registrering, ""O"" => Opsjonell registrering, ""-"" => Registreres ikke") 
		Session.Output(" ")
'end if

if element.AttributesEx.Count > 0 then
														if false then ' not repeat attributes line wise (*unless same name attributes shall have an additional description (somewhere))
															Session.Output("===== Egenskaper")
															for each att in element.AttributesEx
															if not getAttribute(element,att.Name) then ' redefines-ish
																Session.Output("[cols=""20,80""]")
																Session.Output("|===")
																Session.Output("|*Navn:* ")
																Session.Output("|*"&att.name&"*")
																Session.Output(" ")
																Session.Output("|Definisjon: ")
																Session.Output("|"&getCleanDefinition(att.Notes)&"")
																addnotes = getAttributeNotes(element,att.Name)
																if addnotes <> "" then
																	if addnotes <> att.Notes then
																		Session.Output(" Merknad for registreringsinstruks: "&getCleanDefinition(addnotes)&"")
																	end if
																end if
																Session.Output(" ")
																Session.Output("|Multiplisitet: ")
																Session.Output("|["&att.LowerBound&".."&att.UpperBound&"]")
																Session.Output(" ")
																if not att.Default = "" then
																	Session.Output("|Initialverdi: ")
																	Session.Output("|"&att.Default&"")
																	Session.Output(" ")
																end if
																Session.Output("|Type: ")
																if att.ClassifierID <> 0 then
																	Session.Output("|«" & Repository.GetElementByID(att.ClassifierID).Stereotype & "» "&att.Type&"")		
																else
																	Session.Output("|"&att.Type&"")
																end if

																if att.TaggedValues.Count > 0 then
																	Session.Output("|Tagged Values: ")
																	Session.Output("|")
																	for each tag in att.TaggedValues
																		Session.Output(""&tag.Name& ": "&tag.Value&" + ")
																	next
																end if
																Session.Output("|===")
															end if
															next
														end if ' false
	Session.Output(" ")  'v0.14
	Session.Output("<<<")  'v0.14
	Session.Output(" ")  'v0.14
	Session.Output("===== Egenskapstabell for objekttype: "&element.Name&"") 'type, length(?), FKB-standard A/B/C/   'v0.14
'	Session.Output("[cols=""15,15,15,7,7,7,7,7""]")
	Session.Output("[cols=""20,20,20,10""]")
	Session.Output("|===")
	Session.Output("|*Navn:* ")
	Session.Output("|*Type:* ")
	Session.Output("|*SOSI_navn:* ")
	Session.Output("|*Mult.:* ")
'	Session.Output("|*FKB-A:* ")
'	Session.Output("|*FKB-B:* ")
'	Session.Output("|*FKB-C:* ")
'	Session.Output("|*FKB-D:* ")
	Session.Output(" ")
											if false then
											for each att in element.AttributesEx
												if att.name = "Registreringsmetode" or att.name = "Tilleggsbeskrivelse" or att.name = "Grunnrissreferanse" or att.name = "Høydereferanse" then
												else

													if att.Type = "Punkt" or att.Type = "Kurve" or att.Type = "Flate" then
													' GM_Curve etc. TBD
														Session.Output("|"&att.name&"")
														Session.Output("|"&att.Type&"")
														if getTaggedValue(att,"SOSI_navn") = "" then
															Session.Output("|."&UCase(att.Type)&"")
														else
															Session.Output("|."&UCase(getTaggedValue(att,"SOSI_navn"))&"")
														end if
														Session.Output("|["&att.LowerBound&".."&att.UpperBound&"]"&"")
													else
														Session.Output("|"&att.name&"")
														Session.Output("|"&att.Type&"")
														Session.Output("|.."&getTaggedValue(att,"SOSI_navn")&"")
														Session.Output("|["&att.LowerBound&".."&att.UpperBound&"]"&"")
											'			Session.Output("|"&getTaggedValue(att,"SOSI_lengde")&"")
													' new tags in the PS?
										'			Session.Output("|"&getTaggedValue(att,"FKB-A")&"")
										'			Session.Output("|"&getTaggedValue(att,"FKB-B")&"")
										'			Session.Output("|"&getTaggedValue(att,"FKB-C")&"")
										'			Session.Output("|"&getTaggedValue(att,"FKB-D")&"")
													end if
													Session.Output(" ")	

													'nøste seg ut i datatypen?
													if att.ClassifierID <> 0 then
														punktum = "..."
														call listDatatype("", punktum, att)
													end if
													
												end if
											next
											end if 'false	
	call listDatatype("", "..", element)	
	Session.Output("|===")
end if

If element.Connectors.Count > numberSpecializations Then
	Relasjoner(element)
End If

	if element.Constraints.Count > 0 then
		FKBRestriksjoner(element)
	end if

End sub
'-----------------ObjektOgDatatyper End-----------------


' ------------------ Klassebilder Start ---------
Sub klassebilder(element, parentimg,files) 'v0.14
	Dim tag AS EA.TaggedValue
	Dim fil As EA.File
	
	if element.TaggedValues.Count > 0 and files = "" then 'v0.14

		for each tag in element.TaggedValues								
			if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
			'	diagCounter = diagCounter + 1
				Session.Output(" ")
				'Session.Output("'''") 'v0.14
				Session.Output(".Illustrasjon av objekttype "&element.Name&"")
				Session.Output("image::"&tag.Value&"[link="&tag.Value&", Alt=""Illustrasjon av objekttype: "&element.Name&"""]")
				Session.Output(" ")
			end if
		next
		if getTaggedValue(element,"SOSI_bildetekstTilModellelement") <> "" then
			Session.Output("Bildebeskrivelse: "& getCleanDefinition(getTaggedValue(element,"SOSI_bildetekstTilModellelement")) & "")
			Session.Output(" ")
		end if
	end if

	if parentimg <> "" and files = "" then  'v0.14
	'	diagCounter = diagCounter + 1
		Session.Output(" ")
		' Session.Output("'''") 'v0.14
		Session.Output(".Illustrasjon fra produktspesifikasjon av "&element.Name&"")
		Session.Output("image::"&parentimg&"[link="&parentimg&", Alt=""Illustrasjon fra produktspesifikasjon: "&element.Name&"""]")
		Session.Output(" ")
	end if

	if element.Files.Count > 0 and files <> "" then  'v0.14
		For Each fil In element.Files
		'		Session.Output("Filbeskrivelse Name: "& fil.Name & "")
		'		Session.Output("Filbeskrivelse Type: "& fil.Type & "")
		'		Session.Output("Filbeskrivelse Size: "& fil.Size & "")
		'		Session.Output("Filbeskrivelse ObjectType: "& fil.ObjectType & "")
		'		Session.Output("Filbeskrivelse FileDate: "& fil.FileDate & "")
		'		Session.Output("Filbeskrivelse Notes: "& fil.Notes & "")
		'		Session.Output(" ")
		'		diagCounter = diagCounter + 1
			Session.Output(" ")
			Session.Output("'''")
			Session.Output("."&getCleanDefinition(fil.Notes)&"")
			Session.Output("image::"&fil.Name&"[link="&fil.Name&", Alt="""&getCleanDefinition(fil.Notes)&"""]")
			Session.Output(" ")

		next
		Session.Output(" ")
	end if
	Session.Output(" ")
end sub

' ------------------ Klassebilder End ---------


'-----------------CodeList-----------------
Sub Kodelister(element)
Dim att As EA.Attribute
dim tag as EA.TaggedValue
dim utvekslingsalias
	Session.Output("|===")
	Session.Output("|===")
Session.Output(" ")
Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
Session.Output("Definisjon: "&getCleanDefinition(element.Notes)&"")
Session.Output(" ")

if element.TaggedValues.Count > 0 then
	Session.Output("===== Tagged Values")
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	for each tag in element.TaggedValues								
		if tag.Value <> "" then	
			if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" then
			'	Session.Output("|Tag: "&tag.Name&"")
			'	Session.Output("|Verdi: "&tag.Value&"")
				Session.Output("|"&tag.Name&"")
				Session.Output("|"&tag.Value&"")
				Session.Output(" ")			
			end if	
		end if
	next
	Session.Output("|===")
		
	for each tag in element.TaggedValues								
		if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
			diagCounter = diagCounter + 1
		'	Session.Output("[caption=""Figur "&diagCounter&": "",title="&tag.Name&"]")
			Session.Output("[caption=""Figur  "",title="&tag.Name&"]")
		'	Session.Output("image::"&tag.Value&".png["&ThePackage.Name"."&tag.Name&"]")
			Session.Output("image::"&tag.Value&"["&tag.Value&", Alt=""Bilde av kodeliste: "&element.Name&"""]")
		end if
	next
end if
if element.Attributes.Count > 0 then
Session.Output("===== Koder")
end if
utvekslingsalias = false
for each att in element.Attributes
	if att.Default <> "" then
		utvekslingsalias = true
	end if
next
if element.Attributes.Count > 1 then
if utvekslingsalias then
	Session.Output("[cols=""15,25,60""]")
	Session.Output("|===")
	Session.Output("|*Utvekslingsalias:* ")
	Session.Output("|*Kodenavn:* ")
	Session.Output("|*Definisjon:* ")
	Session.Output(" ")
	for each att in element.Attributes
		if att.Default <> "" then
			Session.Output("|"&att.Default&"")
		else
			Session.Output("|")
		end if
		Session.Output("|"&att.Name&"")
		Session.Output("|"&att.Notes&"")
		call attrbilde(att,"kodelistekode")
	next
	Session.Output("|===")
else
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	Session.Output("|*Navn:* ")
	Session.Output("|*Definisjon:* ")
	Session.Output(" ")
	for each att in element.Attributes
		Session.Output("|"&att.Name&"")
		Session.Output("|"&att.Notes&"")
		call attrbilde(att,"kodelistekode")
	next
	Session.Output("|===")
end if
else
for each att in element.Attributes
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	Session.Output("|Navn: ")
	Session.Output("|"&att.name&"")
	Session.Output(" ")
	Session.Output("|Definisjon: ")
	Session.Output("|"&att.Notes&"")
	if not att.Default = "" then
		Session.Output(" ")
		Session.Output("|Utvekslingsalias?: ")
		Session.Output("|"&att.Default&"")
	end if
	call attrbilde(att,"kodelistekode")
	Session.Output("|===")
next
end if
End sub
'-----------------CodeList End-----------------


'-----------------Relasjoner-----------------
sub Relasjoner(element)
Dim generalizations
Dim con
Dim supplier
Dim client
Dim textVar

Session.Output("===== Roller")


'assosiasjoner
For Each con In element.Connectors
	If con.Type = "Association" or con.Type = "Aggregation" Then
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		set supplier = Repository.GetElementByID(con.SupplierID)
		set client = Repository.GetElementByID(con.ClientID)
	'	Session.Output("|Type: ")
	'	Session.Output("|Assosiasjon ")
	'	Session.Output(" ")
		If supplier.elementID = element.elementID Then 'dette elementet er suppliersiden - implisitt at fraklasse er denne klassen
			textVar="|Til klasse"
			If con.ClientEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
			'	textVar=textVar+" _(navigerbar)_:"
			ElseIf con.ClientEnd.Navigable = "Non-Navigable" Then 
				textVar=textVar+" _(ikke navigerbar)_:"
			Else 
				textVar=textVar+":" 
			End If
		'	Session.Output(textVar)
		'	Session.Output("|«" & client.Stereotype&"» "&client.Name)
		'	Session.Output(" ")
			If con.ClientEnd.Role <> "" Then
				Session.Output("|*Rollenavn:* ")
				Session.Output("|*" & con.ClientEnd.Role & "*")
				Session.Output(" ")
			'End If
				If con.ClientEnd.RoleNote <> "" Then
					Session.Output("|Definisjon: ")
					Session.Output("|" & con.ClientEnd.RoleNote)
					Session.Output(" ")
				End If
				If con.ClientEnd.Cardinality <> "" Then
					Session.Output("|Multiplisitet: ")
					Session.Output("|[" & con.ClientEnd.Cardinality&"]")
					Session.Output(" ")
				End If
				Session.Output(textVar)
				Session.Output("|«" & client.Stereotype&"» "&client.Name)
				if false then
				If con.SupplierEnd.Role <> "" Then
					Session.Output("|Fra rolle: ")
					Session.Output("|" & con.SupplierEnd.Role)
					Session.Output(" ")
				End If
				If con.SupplierEnd.RoleNote <> "" Then
					Session.Output("|Fra rolle definisjon: ")
					Session.Output("|" & con.SupplierEnd.RoleNote)
					Session.Output(" ")
				End If
				If con.SupplierEnd.Cardinality <> "" Then
					Session.Output("|Fra multiplisitet: ")
					Session.Output("|[" & con.SupplierEnd.Cardinality&"]")
					Session.Output(" ")
				End If
			End If
			end if
		Else 'dette elementet er clientsiden, (rollen er på target)
			textVar="|Til klasse"
			If con.SupplierEnd.Navigable = "Navigable" Then
			'	textVar=textVar+" _(navigerbar)_:"
			ElseIf con.SupplierEnd.Navigable = "Non-Navigable" Then
				textVar=textVar+" _(ikke-navigerbar)_:"
			Else
				textVar=textVar+":"
			End If
		'	Session.Output(textVar)
		'	Session.Output("|«" & supplier.Stereotype&"» "&supplier.Name)
			If con.SupplierEnd.Role <> "" Then
				Session.Output("|*Rollenavn:* ")
				Session.Output("|*" & con.SupplierEnd.Role & "*")
				Session.Output(" ")
			'	End If
				If con.SupplierEnd.RoleNote <> "" Then
					Session.Output("|Definisjon:")
					Session.Output("|" & con.SupplierEnd.RoleNote)
					Session.Output(" ")
				End If
				If con.SupplierEnd.Cardinality <> "" Then
					Session.Output("|Multiplisitet: ")
					Session.Output("|[" & con.SupplierEnd.Cardinality&"]")
					Session.Output(" ")
				End If
				Session.Output(textVar)
				Session.Output("|«" & supplier.Stereotype&"» "&supplier.Name)
				if false then
				If con.ClientEnd.Role <> "" Then
					Session.Output("|Fra rolle: ")
					Session.Output("|" & con.ClientEnd.Role)
					Session.Output(" ")
				End If
				If con.ClientEnd.RoleNote <> "" Then
					Session.Output("|Fra rolle definisjon: ")
					Session.Output("|" & con.ClientEnd.RoleNote)
					Session.Output(" ")
				End If
				If con.ClientEnd.Cardinality <> "" Then
					Session.Output("|Fra multiplisitet: ")
					Session.Output("|[" & con.ClientEnd.Cardinality&"]")
					Session.Output(" ")
				End If
			End If
			end if
		End If
		Session.Output("|===")
	End If
Next
if false then
'aggregeringer
For Each con In element.Connectors
	If con.Type = "Aggregation" Then
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		set supplier = Repository.GetElementByID(con.SupplierID)
		set client = Repository.GetElementByID(con.ClientID)
		Session.Output("|Type: ")
		If con.clientend.aggregation = 1 Or con.supplierend.aggregation = 1 Then
			Session.Output("|Aggregering")
		ElseIf con.clientend.aggregation = 2 Or con.supplierend.aggregation = 2 Then
			Session.Output("|Komposisjon")
		End If
		Session.Output(" ")
		If supplier.elementID = element.elementID Then 'dette elementet er suppliersiden - implisitt at fraklasse er denne klassen
			textVar="|Til klasse"
			If con.clientend.aggregation = 0 Then 'motsatt side er komponent i denne klassen
				textVar=textVar+" _(del"
			Else
				textVar=textVar+" _(helhet"
			End If
			If con.ClientEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
				textVar=textVar+", navigerbar)_:"
			ElseIf con.ClientEnd.Navigable = "Non-Navigable" Then 
				textVar=textVar+", ikke navigerbar)_:"
			Else 
				textVar=textVar+")_:" 
			End If
			Session.Output(textVar)
			Session.Output("|«" & client.Stereotype&"» "&client.Name)
			Session.Output(" ")
			If con.ClientEnd.Role <> "" Then
				Session.Output("|Til rolle: ")
				Session.Output("|" & con.ClientEnd.Role)
				Session.Output(" ")
			End If
			If con.ClientEnd.RoleNote <> "" Then
				Session.Output("|Til rolle definisjon: ")
				Session.Output("|" & con.ClientEnd.RoleNote)
				Session.Output(" ")
			End If
			If con.ClientEnd.Cardinality <> "" Then
				Session.Output("|Til multiplisitet: ")
				Session.Output("|[" & con.ClientEnd.Cardinality&"]")
				Session.Output(" ")
			End If
			If con.SupplierEnd.Role <> "" Then
				Session.Output("|Fra rolle: ")
				Session.Output("|" & con.SupplierEnd.Role)
				Session.Output(" ")
			End If
			If con.SupplierEnd.RoleNote <> "" Then
				Session.Output("|Fra rolle definisjon: ")
				Session.Output("|" & con.SupplierEnd.RoleNote)
				Session.Output(" ")
			End If
			If con.SupplierEnd.Cardinality <> "" Then
				Session.Output("|Fra multiplisitet: ")
				Session.Output("|[" & con.SupplierEnd.Cardinality&"]")
				Session.Output(" ")
			End If
		Else 'dette elementet er clientsiden
			textVar="|Til klasse"
			If con.supplierEnd.aggregation = 0 Then 'motsatt side er komponent i denne klassen
				textVar=textVar+" _(del"
			Else
				textVar=textVar+" _(helhet"
			End If
			If con.SupplierEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
				textVar=textVar+", navigerbar)_:"
			ElseIf con.SupplierEnd.Navigable = "Non-Navigable" Then 
				textVar=textVar+", ikke navigerbar)_:"
			Else 
				textVar=textVar+")_:" 
			End If
			Session.Output(textVar)
			Session.Output("|«" & supplier.Stereotype&"» "&supplier.Name)
			If con.SupplierEnd.Role <> "" Then
				Session.Output("|Til rolle: ")
				Session.Output("|" & con.SupplierEnd.Role)
				Session.Output(" ")
			End If
			If con.SupplierEnd.RoleNote <> "" Then
				Session.Output("|Til rolle definisjon: ")
				Session.Output("|" & con.SupplierEnd.RoleNote)
				Session.Output(" ")
			End If
			If con.SupplierEnd.Cardinality <> "" Then
				Session.Output("|Til multiplisitet: ")
				Session.Output("|[" & con.SupplierEnd.Cardinality&"]")
				Session.Output(" ")
			End If
			If con.ClientEnd.Role <> "" Then
				Session.Output("|Fra rolle: ")
				Session.Output("|" & con.ClientEnd.Role)
				Session.Output(" ")
			End If
			If con.ClientEnd.RoleNote <> "" Then
				Session.Output("|Fra rolle definisjon: ")
				Session.Output("|" & con.ClientEnd.RoleNote)
				Session.Output(" ")
			End If
			If con.ClientEnd.Cardinality <> "" Then
				Session.Output("|Fra multiplisitet: ")
				Session.Output("|[" & con.ClientEnd.Cardinality&"]")
				Session.Output(" ")
			End If
		End If
		Session.Output("|===")
	End If
Next
end if

' Generaliseringer av pakken
generalizations = False
For Each con In element.Connectors
	If con.Type = "Generalization" Then
		set supplier = Repository.GetElementByID(con.SupplierID)
		set client = Repository.GetElementByID(con.ClientID)
		If supplier.ElementID=element.ElementID then 'dette er en generalisering
			If Not generalizations Then
				Session.Output("[cols=""20,80""]")
				Session.Output("|===")
				Session.Output("|Generalisering av:")
				textVar = "|«" + client.Stereotype + "» " + client.Name
				generalizations = True
			Else
				textVar = textVar + " +" + vbLF + "«" + client.Stereotype + "» " + client.Name
			End If
		End If
	End If
Next
If generalizations then
	Session.Output(textVar)
	Session.Output("|===")
End If

end sub
'-----------------Relasjoner End-----------------



'--------------------Start Sub-------------
sub listDatatype(egenskap, punktum, element)
Dim pktum, eskap, stereo, codeListUrl, teller  'v0.16
'Dim element As EA.Element
Dim datatype As EA.Element
Dim super As EA.Element
Dim att As EA.Attribute
dim tag as EA.TaggedValue  'v0.16
dim conn as EA.Collection
'			Session.Output("|attr.ClassifierID="&attr.ClassifierID&"")
	'		Session.Output("DEBUG: (egenskap, punktum, element.Name: " & egenskap & " , " & punktum & " , " & element.Name )
'set element = Repository.GetElementByID(attr.ClassifierID)
''if element.AttributesEx.Count > 0 then

	for each conn in element.Connectors
	'		if debug then Repository.WriteOutput "Script", "Debug: conn.Type [" & conn.Type & "] conn.ClientID [" & conn.ClientID & "] conn.SupplierID [" & conn.SupplierID & "].",0
		if conn.Type = "Generalization" then
			if element.ElementID = conn.ClientID then
	'				if debug then Repository.WriteOutput "Script", "Debug: supertype [" & Repository.GetElementByID(conn.SupplierID).Name & "].",0
'					superlist = getSupertypes(ftname, conn.SupplierID, indent)
				set super = Repository.GetElementByID(conn.SupplierID)
				call listDatatype(egenskap, punktum, super)
			end if
		end if
	next


	for each att in element.Attributes
		if att.name = "Registreringsmetode" or att.name = "Tilleggsbeskrivelse" or att.name = "Grunnrissreferanse" or att.name = "Høydereferanse" then
		else
			stereo = ""
			if att.Type = "Punkt" or att.Type = "Sverm" or att.Type = "Kurve" or att.Type = "Flate" or Mid(att.Type,1,3) = "GM_" then
				Session.Output("|"&egenskap&att.name&"")
				Session.Output("|"&att.Type&"")
				if getTaggedValue(att,"SOSI_navn") = "" then
					Session.Output("|."&UCase(att.Type)&"")
				else
					Session.Output("|."&UCase(getTaggedValue(att,"SOSI_navn"))&"")
				end if
				Session.Output("|["&att.LowerBound&".."&att.UpperBound&"]"&"")
			else
				Session.Output("|"&egenskap&att.name&"")
				if att.ClassifierID <> 0 then
					set datatype = Repository.GetElementByID(att.ClassifierID)
					stereo = "«" & datatype.Stereotype & "» "
				end if
				if att.ClassifierID <> 0 and LCase(stereo) = "«codelist» " then  'v0.16
					'gjør ingen ting  'v0.16
					codeListUrl = ""  'v0.16
					teller = 0  'v0.16
					for each tag in att.TaggedValues  'v0.16
						if LCase(tag.Name) = "defaultcodespace" and tag.Value <> "" then  'v0.16
							codeListUrl = tag.Value  'v0.16
							Session.Output("| "&stereo&" "&codeListUrl&"["&att.Type&", window = _blank]")  'v0.16
							teller = teller + 1  'v0.16
						end if  'v0.16
					next  'v0.16
					if teller < 1 then  'v0.16
						Session.Output("|"&stereo&att.Type&"")  'v0.16
					end if  'v0.16
				else  'v0.16
					Session.Output("|"&stereo&att.Type&"")	  'v0.16
				end if  'v0.16
				' Session.Output("|"&stereo&att.Type&"")  'v0.16 - denne linjen kan slettes	
				Session.Output("|"&punktum&getTaggedValue(att,"SOSI_navn")&"")
				Session.Output("|["&att.LowerBound&".."&att.UpperBound&"]"&"")
			end if
			Session.Output(" ")	
			
			'nøste seg utover i nye datatyper?
			if att.ClassifierID <> 0 and LCase(stereo) = "«datatype» " then
				pktum = punktum & "."
				eskap = egenskap & att.Name & "."
'				Session.Output("DEBUG2: (eskap, pktum, datatype.Name: " & eskap & " , " & pktum & " , " & datatype.Name )
				call listDatatype(eskap, pktum, datatype)
			end if
		end if
	next

	
' skriv ut roller - sortert etter tagged value sequenceNumber TBD

	for each conn in element.Connectors
		stereo = ""
		if conn.Type = "Association" then
			if element.ElementID = conn.ClientID then
				if conn.SupplierEnd.Role <> "" and conn.SupplierEnd.Navigable = "Navigable" then
					if InStr(LCase(conn.SupplierEnd.Role),"avgrens") = 0 then
		'			if getConnectorEndTaggedValue(conn.SupplierEnd,"xsdEncodingRule") <> "notEncoded" then
						Session.Output("|"&conn.SupplierEnd.Role&"")
						if conn.SupplierID <> 0 then
							set datatype = Repository.GetElementByID(conn.SupplierID)
							stereo = "«" & datatype.Stereotype & "» "
						end if
						Session.Output("|"&stereo&datatype.Name&"")			
						Session.Output("|"&punktum&getConnectorEndTaggedValue(conn.SupplierEnd,"SOSI_navn")&"")
						Session.Output("|["&conn.SupplierEnd.Cardinality&"]"&"")
					end if
				end if
			else
				if conn.ClientEnd.Role <> "" and conn.ClientEnd.Navigable = "Navigable" then
					if InStr(LCase(conn.ClientEnd.Role),"avgrens") = 0 then
						Session.Output("|"&conn.ClientEnd.Role&"")
						if conn.ClientID <> 0 then
							set datatype = Repository.GetElementByID(conn.ClientID)
							stereo = "«" & datatype.Stereotype & "» "
						end if
						Session.Output("|"&stereo&datatype.Name&"")			
						Session.Output("|"&punktum&getConnectorEndTaggedValue(conn.ClientEnd,"SOSI_navn")&"")
						Session.Output("|["&conn.ClientEnd.Cardinality&"]"&"")
					end if
				end if
			end if
			if stereo <> "" then	
				if LCase(datatype.Stereotype) <> "featuretype" then
					supereg = egenskap
					superpktum = punktum
					call listDatatype(supereg,superpktum,datatype)
				end if
			end if
		end if
	next

	'

''end if
end sub

'--------------------End Sub-------------



'-----------------FKBRestriksjoner-----------------
sub FKBRestriksjoner(element)
	Dim constr as EA.Constraint
	Dim datatype as EA.Element
	Dim att as EA.Attribute
	Dim tag as EA.TaggedValue 'v0.14
	dim restriksjon, presisering, egenskapsnavn, datatypeID, subtypeID, objekttypenavn 'v0.14
	restriksjon = 0
	presisering = 0
	objekttypenavn = 0  'v0.14
						
	For Each constr In element.Constraints
		if LCase(Mid(constr.Name,1,20)) <> "_presiseringavkoder_" then
			if restriksjon = 0 then
				Session.Output("")
				Session.Output("===== Restriksjoner")
				restriksjon = 1
				Session.Output("[cols=""20,80""]") 'v0.14
				Session.Output(" ")  'v0.14
				Session.Output("|===")  'v0.14
				Session.Output("|*Navn:* ")  'v0.14
				Session.Output("|*Beskrivelse:* ")  'v0.14
				Session.Output(" ")  'v0.14
			end if 
			if restriksjon > 0 then  'v0.14
				Session.Output("|" & constr.Name & "") 'v0.14
				Session.Output("|" & constr.Notes & "")  'v0.14
				Session.Output(" ")  'v0.14
			end if  'v0.14
		end if
	Next
	
	if restriksjon > 0 then  'v0.14
		Session.Output("|===")  'v0.14
	end if  'v0.14

	For Each constr In element.Constraints
		objekttypenavn = 0  'v0.17
		if LCase(Mid(constr.Name,1,20)) = "_presiseringavkoder_" then
			if presisering = 0 then
				Session.Output("")
				Session.Output("===== Presiseringer til beskrivelsen av kodelistekoder")
				Session.Output("Figurer og skisser knyttet til bruk av bestemte kodelister og koder.")
				Session.Output(" ")
				presisering = 1
			end if
			' finn rett kode via egenskapsnavnet
			egenskapsnavn = Mid(constr.Name,21,Len(constr.Name)-20)
	'		Session.Output("===== Gjelder egenskap: " & element.Name & "." & egenskapsnavn & "")
			' finn datatypen til egenskapen
			datatypeID = getDTID(egenskapsnavn,element)
			if datatypeID <> 0 then
				set datatype = Repository.GetElementByID(datatypeID)
	'			Session.Output("===== Gjelder egenskapens datatype: " & element.Name & "." & egenskapsnavn & " : " & datatype.Name & "")
				' finn subtypeklasse med samme navn
				subtypeID = getSTID(datatype)
				if subtypeID <> 0 then
					set datatype = Repository.GetElementByID(subtypeID)
	'				Session.Output("===== Gjelder tilsvarende datatype: " & element.Name & "." & egenskapsnavn & " : " & datatype.Name & "")
					' list ut alle koder med presiseringer
					if datatype.Attributes.Count > 0 then  'v0.14
						for each att in datatype.Attributes
							if att.TaggedValues.Count > 0 then   'v0.14
								for each tag in att.TaggedValues   'v0.14
									if LCase(tag.Name) = "fkb_objekttype" and LCase(tag.Value) = LCase(element.Name) then 'v0.14
										Session.Output("")
										'Session.Output("===== " & element.Name & "." & egenskapsnavn & " : " & datatype.Name & " - Kode : " & att.Name & "")  v0.14-erstattes av linjene under
										'Session.Output("*Definisjon :* " & getCleanDefinition(att.Notes) & "")  v0.14-erstattes av linjene under
										Session.Output("===== " & datatype.Name & " - Kodenavn: " & att.Name & "")  'v0.14
										Session.Output("*Definisjon:* " & getCleanDefinition(att.Notes) & "")  'v0.14
										call kodebilde(att)
										Session.Output("")
										objekttypenavn = 1  'v0.14
									end if  'v0.14
								next  'v0.14
							end if  'v0.14
							if objekttypenavn = 0 then  'v0.14
								Session.Output("")  'v0.14
								'Session.Output("===== " & element.Name & "." & egenskapsnavn & " : " & datatype.Name & " - Kode : " & att.Name & "")  v0.14-erstattes av linjene under
								'Session.Output("*Definisjon :* " & getCleanDefinition(att.Notes) & "")  v0.14-erstattes av linjene under
								Session.Output("===== " & datatype.Name & " - Kodenavn: " & att.Name & "")  'v0.14
								Session.Output("*Definisjon:* " & getCleanDefinition(att.Notes) & "")  'v0.14
								call kodebilde(att)  'v0.14
								Session.Output("")  'v0.14
							end if  'v0.14
						next
					end if  'v0.14
				else
					Session.Output("===== Fant ingen klasse med presiseringer i restriksjonen: " & constr.Name & " - " & egenskapsnavn & " : " & datatype.Name & "")
				end if
			end if
		end if
	Next

end sub
'-----------------FKBRestriksjoner End-----------------



'----------------- Start Funksjon getDTID-----------------
' finn ID til datatypen til egenskapen, eventuelt i en supertype
Function getDTID(egenskapsnavn,element)
	Dim supertype as EA.Element
	Dim con As EA.Connector	
	Dim att as EA.Attribute
	Dim t
	getDTID = 0
	for each att in element.Attributes		
		if att.Name = egenskapsnavn then
			getDTID = att.ClassifierID
		end if
	next
	if getDTID = 0 then
		For Each con In element.Connectors
			If con.Type = "Generalization" and con.ClientID = element.ElementID Then
				set supertype = Repository.GetElementByID(con.SupplierID)
				getDTID = getDTID(egenskapsnavn,supertype)
			End If
		Next	
	end if
end function
'----------------- End Funksjon getDTID-----------------


'----------------- Start Funksjon getSTID-----------------
' finn ID til en subtypeklasse med samme navn
Function getSTID(element)
	Dim subtype as EA.Element
	Dim con As EA.Connector	
	getSTID = 0
	For Each con In element.Connectors
		set subtype = Repository.GetElementByID(con.ClientID)
		If con.Type = "Generalization" And subtype.ElementID <> element.ElementID Then
			if subtype.Name = element.Name then
				getSTID = subtype.ElementID
			end if
		End If
	Next
end function
'----------------- End Funksjon getDTID-----------------

'--------------------Start Sub-------------
sub	listFkbTag(tagName, tagValue)
'		Session.Output("""P"" => Påkrevd registrering, ""O"" => Opsjonell registrering, ""-"" => Registreres ikke") 

	if LCase(Mid(tagName,1,20)) = "fkb_minstestørrelse_" then  'v0.15
		Session.Output("|FKB-" &Mid(tagName,21,21)& " minstestørrelse")  'v0.15
		Session.Output("|"&getCleanDefinition(tagValue)&"")  'v0.15
	else   'v0.15
		Session.Output("|"&underscore2space(tagName)&"")  'v0.15
		if tagValue = "P" then  'v0.15
			Session.Output("|Påkrevd registrering")  'v0.15
		else  'v0.15
			if tagValue = "O" then  'v0.15
				Session.Output("|Opsjonell registrering")  'v0.15
			else  'v0.15
				if tagValue = "-" then  'v0.15
					Session.Output("|Registreres ikke")  'v0.15
				else  'v0.15
					Session.Output("|"&getCleanDefinition(tagValue)&"")  'v0.15
				end if  'v0.15
			end if  'v0.15
		end if  'v0.15
	end if  'v0.15
	Session.Output(" ")	
end sub
'--------------------End Sub-------------

'-----------------Funksjon for full path-----------------
function getPath(package)
	dim path
	dim parent
	if package.parentID <> 0 then
'		Session.Output(" -----DEBUG getPath=" & getPath & " package.Name = " & package.Name & " package.ParentID = " & package.ParentID & " package.Element.Stereotype = " & package.Element.Stereotype & " ----- ")
		if package.Element.Stereotype = "" then
			path = package.Name
		else
			path = "«" + package.Element.Stereotype + "» " + package.Name
		end if

		if ucase(package.Element.Stereotype) <> "APPLICATIONSCHEMA" then
			set parent = Repository.GetPackageByID(package.ParentID)
			path = getPath(parent) + "/" + path
		end if
	end if
	getPath = path
end function
'-----------------Funksjon for full path End-----------------



'-----------------Funksjon for å sjekke om egenskap finnes i denne klassen Start-----------------
function getAttribute(element,attributeName)
	Dim att As EA.Attribute
	getAttribute = false
	if element.Attributes.Count > 0 then
		for each att in element.Attributes
			if att.Name = attributeName then
				getAttribute = true
			end if
		next
	end if
end function
'-----------------Funksjon End-----------------



'-----------------Funksjon for å hente notefelt fra navngitt egenskap i en klasaseStart-----------------
function getAttributeNotes(element,attributeName)
	Dim att As EA.Attribute
	getAttributeNotes = ""
	if element.Attributes.Count > 0 then
		for each att in element.Attributes
			if att.Name = attributeName then
				getAttributeNotes = att.Notes
			end if
		next
	end if
end function
'-----------------Funksjon for å hente notefelt fra egenskap End-----------------



'-----------------Function underscore2space Start-----------------
function underscore2space(txt)
	'replaces underscores with spaces
    Dim res, tegn, i
	underscore2space = ""

	res = ""
	' loop gjennom alle tegn
	For i = 1 To Len(txt)
		tegn = Mid(txt,i,1)
		If tegn = "_" Then
			res = res + " "
		Else
			res = res + Mid(txt,i,1)
		end if
		  
	Next
		
	underscore2space = res

end function
'-----------------Function underscore2space End-----------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Func Name: attrbilde(att)
' Author: Kent Jonsrud
' Date: 2021-09-16
' Date: 2021-10-06 FKB-utvidelser
' Purpose: skriver ut lenke og tekst til bilde av element ved siden av elementet

sub attrbilde(att,typ)
	dim tag as EA.TaggedValue
	dim bildetekst
	bildetekst = "Illustrasjon av " & typ & " "&att.Name
	for each tag in att.TaggedValues								
		if LCase(tag.Name) = "fkb_bildetekst" and tag.Value <> "" then
			bildetekst = getCleanDefinition(tag.Value)
		end if
	next
	for each tag in att.TaggedValues								
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			Session.Output(" +")
			Session.Output(""&bildetekst&"")
			Session.Output("image:"&tag.Value&"[link="&tag.Value&",width=100,height=100, Alt=""" & bildetekst & """]")
		end if
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Func Name: kodebilde(att)
' Author: Kent Jonsrud
' Date: 2021-10-06
' Purpose: skriver ut lenke og tekst til bilde av element under elementet

sub kodebilde(att)
	dim tag as EA.TaggedValue
	dim bildetekst
	bildetekst = "Illustrasjon av "&att.Name
	for each tag in att.TaggedValues								
		if LCase(tag.Name) = "fkb_bildetekst" and tag.Value <> "" then
			bildetekst = getCleanDefinition(tag.Value)
		end if
	next
	for each tag in att.TaggedValues								
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			Session.Output(" ")
			Session.Output("."&bildetekst&"")
			Session.Output("image::"&tag.Value&"[link="&tag.Value&", Alt=""" & bildetekst & """]")
			Session.Output(" ")
		end if
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'-----------------Function getCleanDefinition Start-----------------
function getCleanDefinition(txt)
	'removes all formatting in notes fields, except crlf
    Dim res, tegn, i, u
    u=0
	getCleanDefinition = ""

		res = ""
		txt = Trim(txt)
		For i = 1 To Len(txt)
		  tegn = Mid(txt,i,1)
			if tegn = "," then tegn = " " 'for adoc
			If tegn = "<" Then
				u = 1
				tegn = " "
			end if 
			If tegn = ">" Then
				u = 0
				tegn = " "
			end if
			If tegn < " " Then
				res = res + " "
			Else
				if u = 0 then
					res = res + Mid(txt,i,1)
				end if
			End If
		  
		Next
		
	getCleanDefinition = res

end function
'-----------------Function getCleanDefinition End-----------------



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


OnProjectBrowserScript
