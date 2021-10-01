Option Explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: listAdocFraModell
' Author: Tore Johnsen, Åsmund Tjora
' Purpose: Generate documentation in AsciiDoc syntax
' Original Date: 08.04.2021
'
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
Dim diagCounter,figurcounter
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
			figurcounter = 0
			Dim thePackage As EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			imgfolder = "diagrammer"
			Set imgFSO=CreateObject("Scripting.FileSystemObject")
			imgparent = imgFSO.GetParentFolderName(Repository.ConnectionString())  & "\" & imgfolder
			if not imgFSO.FolderExists(imgparent) then
				imgFSO.CreateFolder imgparent
			end if

			Call ListAsciiDoc(thePackage)
			Session.Output("// End of UML-model")
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
	Dim listTags
		
	if thePackage.Element.Stereotype <> "" then
		Session.Output("=== Pakke «"&thePackage.Element.Stereotype&"» "&thePackage.Name&"")
	else
		Session.Output("")
		Session.Output("<<<")
		Session.Output("'''")

		Session.Output("=== Pakke: "&thePackage.Name&"")
	end if
	Session.Output("*Definisjon:* "&getCleandefinition(thePackage.Notes)&"")

	if thePackage.element.TaggedValues.Count > 0 then
		listTags = false
		for each tag in thePackage.element.TaggedValues
			if tag.Value <> "" then	
				if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" then
					if listTags = false then
						Session.Output(" ")	
						Session.Output("===== Profilparametre i tagged values")
						Session.Output("[cols=""20,80""]")
						Session.Output("|===")
						listTags = true
					end if
					Session.Output("|"&tag.Name&"")
					Session.Output("|"&tag.Value&"")
					Session.Output(" ")			
				end if
			end if
		next
		if listTags = true then
			Session.Output("|===")
		end if
	end if

'-----------------Diagram-----------------

	for each tag in thePackage.element.TaggedValues
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			diagCounter = diagCounter + 1
			Session.Output(".Illustrasjon av pakke "&thePackage.Name&"")
			Session.Output("image::"&tag.Value&"[link="&tag.Value&",""Illustrasjon av pakke: "&thePackage.Name&"""]")
		end if
	next

	For Each diag In thePackage.Diagrams
		diagCounter = diagCounter + 1
		Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, imgparent & "\" & diag.Name & ".png", 1)
		Repository.CloseDiagram(diag.DiagramID)
		Session.Output(" ")
		Session.Output("'''")
		Session.Output(" ")
		Session.Output("."&diag.Name&" ")
		Session.Output("image::diagrammer/"&diag.Name&".png[link=diagrammer/"&diag.Name&".png,""Diagramm: "&diag.Name&"""]")
	Next

	For each element in thePackage.Elements
		If Ucase(element.Stereotype) = "FEATURETYPE" Then
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
		
	dim pack as EA.Package
	for each pack in thePackage.Packages
		Call ListAsciiDoc(pack)
	next

	Set imgFSO = Nothing
end sub

'-----------------ObjektOgDatatyper-----------------
	Sub ObjektOgDatatyper(element)
	Dim att As EA.Attribute
	dim tag as EA.TaggedValue
	Dim con As EA.Connector
	Dim supplier As EA.Element
	Dim client As EA.Element
	Dim association
	Dim aggregation
	association = False
 
	Dim numberSpecializations, numberGeneralizations, numberRealisations

	Dim textVar
	dim externalPackage
	Dim listTags

	Session.Output(" ")
	Session.Output("'''")

	Session.Output(" ")

	Session.Output("[["&LCase(element.Name)&"]]")
	if element.Abstract = 1 then
		Session.Output("==== «"&element.Stereotype&"» "&element.Name&" (abstrakt)")
	else
		Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
	end if
	Session.Output("*Definisjon:* "&getCleanDefinition(element.Notes)&"")
	Session.Output(" ")


	if element.TaggedValues.Count > 0 then
		for each tag in element.TaggedValues								
			if tag.Value <> "" then	
				if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" and LCase(tag.Name) <> "sosi_bildeavmodellelement" then
					if listTags = false then
						Session.Output("===== Profilparametre i tagged values")
						Session.Output("[cols=""20,80""]")
						Session.Output("|===")
						listTags = true
					end if
					Session.Output("|"&tag.Name&"")
					Session.Output("|"&tag.Value&"")
					Session.Output(" ")			
				end if
			end if
		next
		if listTags = true then
			Session.Output("|===")
		end if
		
		for each tag in element.TaggedValues								
			if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
				diagCounter = diagCounter + 1
				Session.Output(" ")
				Session.Output("'''")
				Session.Output(".Illustrasjon av objekttype "&element.Name&"")
				Session.Output("image::"&tag.Value&"[link="&tag.Value&",""Illustrasjon av objekttype: "&element.Name&"""]")
			end if
		next
	end if

	if element.Attributes.Count > 0 then
		Session.Output("===== Egenskaper")
		for each att in element.Attributes
			Session.Output("[cols=""20,80""]")
			Session.Output("|===")
			Session.Output("|*Navn:* ")
			Session.Output("|*"&att.name&"*")
			Session.Output(" ")
			Session.Output("|Definisjon: ")
			Session.Output("|"&getCleanDefinition(att.Notes)&"")
			Session.Output(" ")
			Session.Output("|Multiplisitet: ")
			Session.Output("|["&att.LowerBound&".."&att.UpperBound&"]")
			Session.Output(" ")
			if not att.Default = "" then
				Session.Output("|Initialverdi: ")
				Session.Output("|"&att.Default&"")
				Session.Output(" ")
			end if
			if not att.Visibility = "Public" then
				Session.Output("|Visibilitet: ")
				Session.Output("|"&att.Visibility&"")
				Session.Output(" ")
			end if
			Session.Output("|Type: ")
			if att.ClassifierID <> 0 then
				if isElement(att.ClassifierID) then
					dim stereo
					stereo = Repository.GetElementByID(att.ClassifierID).Stereotype
					if stereo = "" then
						Session.Output("|<<"&LCase(att.Type)&","&att.Type&">>")
					else
						Session.Output("|<<"&LCase(att.Type)&",«" & stereo & "» "&att.Type&">>")
					end if
				else
					Session.Output("|"&att.Type&"")
				end if
			else
				Session.Output("|http://skjema.geonorge.no/SOSI/basistype/"&att.Type&"["&att.Type&"]")		
			end if

			if att.TaggedValues.Count > 0 then
				Session.Output("|Profilparametre i tagged values: ")
				Session.Output("|")
				for each tag in att.TaggedValues
					Session.Output(""&tag.Name& ": "&tag.Value&" + ")
				next
			end if
			Session.Output("|===")
		next
	end if


	Relasjoner(element)

	if element.Methods.Count > 0 then
		Operasjoner(element)
	end if

	if element.Constraints.Count > 0 then
		Restriksjoner(element)
	end if

' Supertype
	numberSpecializations = 0
	For Each con In element.Connectors
		set supplier = Repository.GetElementByID(con.SupplierID)
		If con.Type = "Generalization" And supplier.ElementID <> element.ElementID Then
			if numberSpecializations = 0 then
				Session.Output("===== Arv og realiseringer")
				Session.Output("[cols=""20,80""]")
				Session.Output("|===")
			end if
			numberSpecializations = numberSpecializations + 1
			Session.Output("|Supertype: ")
			Session.Output("|<<"&LCase(supplier.Name)&",«" & supplier.Stereotype&"» "&supplier.Name&">>")
			Session.Output(" ")
		End If
	Next



' Spesialiseringer av klassen
	numberGeneralizations = 0
	For Each con In element.Connectors
		If con.Type = "Generalization" Then
			set supplier = Repository.GetElementByID(con.SupplierID)
			set client = Repository.GetElementByID(con.ClientID)
			If supplier.ElementID = element.ElementID then 'dette er en generalisering
				if numberSpecializations = 0 and numberGeneralizations = 0 then
					Session.Output("===== Arv og realiseringer")
					Session.Output("[cols=""20,80""]")
					Session.Output("|===")
				end if		
				If numberGeneralizations = 0 Then
					Session.Output("|Subtyper:")
					Session.Output("|<<"&LCase(client.Name)&",«" & client.Stereotype & "» " & client.Name & ">> +")
				Else
					Session.Output("<<"&LCase(client.Name)&",«" & client.Stereotype & "» " & client.Name & ">> +")
				End If
				numberGeneralizations = numberGeneralizations + 1
			End If
		End If
	Next

	For Each con In element.Connectors  
		numberRealisations = 0
'Må forbedres i framtidige versjoner dersom denne skal med 
'- full sti (opp til applicationSchema eller øverste pakke under "Model") til pakke som inneholder klassen som realiseres
		set supplier = Repository.GetElementByID(con.SupplierID)
		If con.Type = "Realisation" And supplier.ElementID <> element.ElementID Then
			if numberSpecializations = 0 and numberGeneralizations = 0 and numberRealisations = 0 then
				Session.Output("===== Arv og realiseringer")
				Session.Output("[cols=""20,80""]")
				Session.Output("|===")
			end if		
			set externalPackage = Repository.GetPackageByID(supplier.PackageID)
			textVar=getPath(externalPackage)
			if numberRealisations = 0 Then
				Session.Output("|Realisering av: ")
				Session.Output("|" & textVar &"::«" & supplier.Stereotype&"» "&supplier.Name&" +")
				numberRealisations = numberRealisations + 1
			else
				Session.Output("" & textVar &"::«" & supplier.Stereotype&"» "&supplier.Name&" +")
				Session.Output(" ")
			end if
			numberRealisations = numberRealisations + 1
		end if
	next

	If numberSpecializations + numberGeneralizations + numberRealisations > 0 then
		Session.Output("|===")
	End If

End sub
'-----------------ObjektOgDatatyper End-----------------


'-----------------CodeList-----------------
Sub Kodelister(element)
	Dim att As EA.Attribute
	dim tag as EA.TaggedValue
	dim utvekslingsalias, codeListUrl, asdict
	asdict = false
	Session.Output(" ")
	Session.Output("'''")
	 

	Session.Output(" ")
	Session.Output("[["&LCase(element.Name)&"]]")
	Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
	Session.Output("*Definisjon:* "&getCleanDefinition(element.Notes)&"")
	Session.Output(" ")

	if element.TaggedValues.Count > 0 then
		Session.Output("===== Profilparametre i tagged values")
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		for each tag in element.TaggedValues								
			if tag.Value <> "" then	
				if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" and LCase(tag.Name) <> "sosi_bildeavmodellelement" then
					Session.Output("|"&tag.Name&"")
					Session.Output("|"&tag.Value&"")
					Session.Output(" ")			
				end if	
			end if
		next
		Session.Output("|===")
			
		codeListUrl = ""	
		for each tag in element.TaggedValues								
			if LCase(tag.Name) = "asdictionary" and tag.Value = "true" then asdict = true
			if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
				diagCounter = diagCounter + 1
				Session.Output("image::"&tag.Value&"["&tag.Value&"]")
			end if
			if LCase(tag.Name) = "codelist" and tag.Value <> "" then
				codeListUrl = tag.Value
			end if
		next
	end if

	if codeListUrl <> "" and asdict then
		Session.Output("Koder fra ekstern kodeliste kan hentes fra register: "&codeListUrl&"")	
		Session.Output(" ")
	end if


	if element.Attributes.Count > 0 then
		Session.Output("===== Koder i modellen")
	end if
	utvekslingsalias = false
	for each att in element.Attributes
		if att.Default <> "" then
			utvekslingsalias = true
		end if
	next
	if element.Attributes.Count > 0 then
		if utvekslingsalias then
			Session.Output("[cols=""25,60,15""]")
			Session.Output("|===")
			Session.Output("|*Kodenavn:* ")
			Session.Output("|*Definisjon:* ")
			Session.Output("|*Utvekslingsalias:* ")
			Session.Output(" ")
			for each att in element.Attributes

				Session.Output("|"&att.Name&"")
				Session.Output("|"&getCleanDefinition(att.Notes)&"")
				if att.Default <> "" then
					Session.Output("|"&att.Default&"")
				else
					Session.Output("|")
				end if
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
				Session.Output("|"&getCleanDefinition(att.Notes)&"")
				call attrbilde(att,"kodelistekode")
			next
			Session.Output("|===")
		end if

	end if
End sub
'-----------------CodeList End-----------------


'-----------------Relasjoner-----------------
sub Relasjoner(element)
	Dim generalizations
	Dim con
	Dim supplier
	Dim client
	Dim textVar, skrivRoller

	skrivRoller = false


'assosiasjoner
' skriv ut roller - sortert etter tagged value sequenceNumber TBD

	For Each con In element.Connectors
		If con.Type = "Association" or con.Type = "Aggregation" Then
			set supplier = Repository.GetElementByID(con.SupplierID)
			set client = Repository.GetElementByID(con.ClientID)
			If supplier.elementID = element.elementID Then 'dette elementet er suppliersiden - implisitt at fraklasse er denne klassen
				textVar="|Til klasse"
				If con.ClientEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
				ElseIf con.ClientEnd.Navigable = "Non-Navigable" Then 
					textVar=textVar+" _(ikke navigerbar)_:"
				Else 
					textVar=textVar+":" 
				End If
				If con.ClientEnd.Role <> "" Then
					if skrivRoller = false then
						Session.Output("")
						Session.Output("===== Roller")
						Session.Output("[cols=""20,80""]")
						Session.Output("|===")
						skrivRoller = true
					else
						Session.Output("[cols=""20,80""]")
						Session.Output("|===")
					end if
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
					If con.SupplierEnd.Aggregation <> 0 Then
						Session.Output("|Assosiasjonstype: ")
						if con.SupplierEnd.Aggregation = 2 then
							Session.Output("|Komposisjon " & con.Type)
						else
							Session.Output("|Aggregering " & con.Type)
						end if
						Session.Output(" ")
					End If
					If con.Name <> "" Then
						Session.Output("|Assosiasjonsnavn: ")
						Session.Output("|" & con.Name)
						Session.Output(" ")
					End If

					Session.Output(textVar)
					Session.Output("|<<"&LCase(client.Name)&","&"«" & client.Stereotype&"» "&client.Name&">>")
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
				ElseIf con.SupplierEnd.Navigable = "Non-Navigable" Then
					textVar=textVar+" _(ikke-navigerbar)_:"
				Else
					textVar=textVar+":"
				End If
				If con.SupplierEnd.Role <> "" Then
					if skrivRoller = false then
						Session.Output("")
						Session.Output("===== Roller")
						Session.Output("[cols=""20,80""]")
						Session.Output("|===")
						skrivRoller = true
					else
						Session.Output("[cols=""20,80""]")
						Session.Output("|===")
						
					end if
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
					If con.ClientEnd.Aggregation <> 0 Then
						Session.Output("|Assosiasjonstype: ")
						if con.ClientEnd.Aggregation = 2 then
							Session.Output("|Komposisjon " & con.Type)
						else
							Session.Output("|Aggregering " & con.Type)
						end if
						Session.Output(" ")
					End If
					If con.Name <> "" Then
						Session.Output("|Assosiasjonsnavn: ")
						Session.Output("|" & con.Name)
						Session.Output(" ")
					End If

					Session.Output(textVar)
					Session.Output("|<<"&LCase(supplier.Name)&","&"«" & supplier.Stereotype&"» "&supplier.Name&">>")
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
			if skrivRoller = true then
				Session.Output("|===")
			end if
		End If
	Next



end sub
'-----------------Relasjoner End-----------------



'-----------------Operasjoner-----------------
sub Operasjoner(element)
	Dim meth as EA.Method

	Session.Output("")
	Session.Output("===== Operasjoner")

						
	For Each meth In element.Methods
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		Session.Output("|*Navn:* ")
		Session.Output("|*" & meth.Name & "*")
		Session.Output(" ")
		Session.Output("|Beskrivelse: ")
		Session.Output("|" & meth.Notes & "")
		Session.Output(" ")
	'	Session.Output("|Stereotype: ")
	'	Session.Output("|" & meth.Stereotype & "")
	'	Session.Output(" ")
	'	Session.Output("|Retur type: ")
	'	Session.Output("|" & meth.ReturnType & "")
	'	Session.Output(" ")
	'	Session.Output("|Oppførsel: ")
	'	Session.Output("|" & meth.Behaviour & "")
	'	Session.Output(" ")
		Session.Output("|===")
	Next

end sub
'-----------------Operasjoner End-----------------


'-----------------Restriksjoner-----------------
sub Restriksjoner(element)
	Dim constr as EA.Constraint

	Session.Output("")
	Session.Output("===== Restriksjoner")

						
	For Each constr In element.Constraints
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		Session.Output("|*Navn:* ")
		Session.Output("|*" & constr.Name & "*")
		Session.Output(" ")
		Session.Output("|Beskrivelse: ")
		Session.Output("|" & constr.Notes & "")
		Session.Output(" ")
	'	Session.Output("|Type: ")
	'	Session.Output("|" & constr.Type & "")
	'	Session.Output(" ")
	'	Session.Output("|Status: ")
	'	Session.Output("|" & constr.Status & "")
	'	Session.Output(" ")
	'	Session.Output("|Vekt: ")
	'	Session.Output("|" & constr.Weight & "")
	'	Session.Output(" ")
		Session.Output("|===")
	Next

end sub
'-----------------Restriksjoner End-----------------



'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Func Name: attrbilde(att)
' Author: Kent Jonsrud
' Date: 2021-09-16
' Purpose: skriver ut lenke til bilde av element ved siden av elementet

sub attrbilde(att,typ)
	dim tag as EA.TaggedValue
	for each tag in att.TaggedValues								
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			Session.Output(" +")
			Session.Output("Illustrasjon av " & typ & " "&att.Name&"")
			Session.Output("image:"&tag.Value&"[link="&tag.Value&",width=100,height=100,""Illustrasjon av " & typ & ": "&att.Name&"""]")
		end if
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------



'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Func Name: isElement
' Author: Kent Jonsrud
' Date: 2021-07-13
' Purpose: tester om det finnes et element med denne ID-en.

function isElement(ID)
	isElement = false
	if 	Mid(Repository.SQLQuery("select count(*) from t_object where Object_ID = " & ID & ";"), 113, 1) <> 0 then
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


'-----------------Function getCleanDefinition Start-----------------
function getCleanDefinition(txt)
	'removes all formatting in notes fields, except crlf
    Dim res, tegn, i, u
    u=0
	getCleanDefinition = ""

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
				  If tegn < " " and Asc(tegn) <> 10 and Asc(tegn) <> 13 Then
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
		
	getCleanDefinition = res

end function
'-----------------Function getCleanDefinition End-----------------


'-----------------Function nao Start-----------------
function nao()
					' I just want a correct xml timestamp to document when the script was run
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
'-----------------Function nao End-----------------

OnProjectBrowserScript
