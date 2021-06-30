Option Explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: listAdocFraModell
' Author: Tore Johnsen, Åsmund Tjora
' Purpose: Generate documentation in AsciiDoc syntax
' Date: 08.04.2021
'
' Version: 0.6 Date: 2021-06-30 Kent Jonsrud: leser kodelister fra levende register
'
' Version: 0.5 Date: 2021-06-29 Kent Jonsrud: error if role list is not shown
'
' Version: 0.4
' Date: 2021-06-14 Kent Jonsrud: case-insensitiv test på navnet på tagged value SOSI_bildeAvModellelement
' Date: 2021-06-24 Kent Jonsrud: endra navn
'
' Version: 0.3
' Date: 2021-06-01 Kent Jonsrud:
' - retta bildesti til app_img
' Version: 0.2
' Date: 2021-04-16 Kent Jonsrud:
' - tagged value SOSI_bildeAvModellelement på pakker og klasser: verdien vises som ekstern sti til bilde
' Date: 2021-04-15 Kent Jonsrud:
' - diagrammer legges i underkatalog med navn enten verdien i tagged value SOSI_kortnavn eller img
' TBD: sette inn blanke i diagrammnavn for enklere_filnavn.png ?
' Date: 2021-04-09/14 Kent Jonsrud:
' - tagged value lists come right after definition, on packages and classes
' - "Spesialisering av" changed to Supertype, no list of subtypes shown
' - removed formatting in notes, except CRLF
' - show stereotype on attribute "Type" if present
' - roles shall have same simple look and structure as attributes
' - Relasjoner changed to Roller, show only ends with role names (and navigable ?)
' - tagged values on CodeList classes, empty tags suppressed (suppress only those from the standard profile?), heading?
' - simpler list for codelists with more than 1 code, three-column list when Defaults are used (Utvekslingsalias)
'
' TBD: show stereotype on Type 
' TBD: show navigable 
' TBD: show association type 
' TBD: output operations and constraints right after attributes and roles
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
listTags = false
	
if thePackage.Element.Stereotype <> "" then
	Session.Output("=== «"&thePackage.Element.Stereotype&"» "&thePackage.Name&"")
else
	Session.Output("=== Pakke: "&thePackage.Name&"")
end if
Session.Output("Definisjon: "&getCleandefinition(thePackage.Notes)&"")

if thePackage.element.TaggedValues.Count > 0 then

	for each tag in thePackage.element.TaggedValues
		if tag.Value <> "" then	
			if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" then
				if listTags = false then
					Session.Output(" ")	
					Session.Output("===== Tagged Values")
					Session.Output("[cols=""20,80""]")
					Session.Output("|===")
					listTags = true
				end if
	'	Session.Output("|Tag: "&tag.Name&"")
			'	Session.Output("|Verdi: "&tag.Value&"")
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
'		if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			diagCounter = diagCounter + 1
			Session.Output("[caption=""Figur "&diagCounter&": "",title="&tag.Name&"]")
		'	Session.Output("image::"&tag.Value&".png["&ThePackage.Name"."&tag.Name&"]")
			Session.Output("image::"&tag.Value&".png["&tag.Value&"]")
		end if
	next

For Each diag In thePackage.Diagrams
	diagCounter = diagCounter + 1
	Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, imgparent & "\" & diag.Name & ".png", 1)
'	Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, "" & diag.Name&".png", 1)
	Repository.CloseDiagram(diag.DiagramID)
	Session.Output("[caption=""Figur "&diagCounter&": "",title="&diag.Name&"]")
	Session.Output("image::"&diag.Name&".png["&diag.Name&"]")
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
Dim generalizations
Dim numberSpecializations ' tar også med antall realiseringer her
Dim textVar
dim externalPackage
Dim listTags

Session.Output(" ")
Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
Session.Output("Definisjon: "&getCleanDefinition(element.Notes)&"")
Session.Output(" ")
numberSpecializations = 0
For Each con In element.Connectors
	set supplier = Repository.GetElementByID(con.SupplierID)
	If con.Type = "Generalization" And supplier.ElementID <> element.ElementID Then
		Session.Output("*Supertype:* «" & supplier.Stereotype&"» "&supplier.Name&"")
		Session.Output(" ")
		numberSpecializations = numberSpecializations + 1
	End If
Next
For Each con In element.Connectors  
'realiseringer.  
'Må forbedres i framtidige versjoner dersom denne skal med 
'- full sti (opp til applicationSchema eller øverste pakke under "Model") til pakke som inneholder klassen som realiseres
	set supplier = Repository.GetElementByID(con.SupplierID)
	If con.Type = "Realisation" And supplier.ElementID <> element.ElementID Then
		set externalPackage = Repository.GetPackageByID(supplier.PackageID)
		textVar=getPath(externalPackage)
		Session.Output("*Realisering av:* " & textVar &"::«" & supplier.Stereotype&"» "&supplier.Name)
		Session.Output(" ")
		numberSpecializations = numberSpecializations + 1
	end if
next

if element.TaggedValues.Count > 0 then
	for each tag in element.TaggedValues								
		if tag.Value <> "" then	
			if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" then
				if listTags = false then
					Session.Output("===== Tagged Values")
					Session.Output("[cols=""20,80""]")
					Session.Output("|===")
					listTags = true
				end if
			'	Session.Output("|Tag: "&tag.Name&"")
			'	Session.Output("|Verdi: "&tag.Value&"")
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
'		if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			diagCounter = diagCounter + 1
			Session.Output("[caption=""Figur "&diagCounter&": "",title="&tag.Name&"]")
		'	Session.Output("image::"&tag.Value&".png["&ThePackage.Name"."&tag.Name&"]")
			Session.Output("image::"&tag.Value&"["&tag.Value&"]")
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
		Session.Output("|Type: ")
	'	if att.ClassifierID <> 0 then
	'		Session.Output("|«" & Repository.GetElementByID(att.ClassifierID).Stereotype & "» "&att.Type&"")		
	'	else
			Session.Output("|"&att.Type&"")
	'	end if

		if att.TaggedValues.Count > 0 then
			Session.Output("|Tagged Values: ")
			Session.Output("|")
			for each tag in att.TaggedValues
				Session.Output(""&tag.Name& ": "&tag.Value&" + ")
			next
		end if
		Session.Output("|===")
	next
end if

If element.Connectors.Count > numberSpecializations Then
	Relasjoner(element)
End If
End sub
'-----------------ObjektOgDatatyper End-----------------


'-----------------CodeList-----------------
Sub Kodelister(element)
Dim att As EA.Attribute
dim tag as EA.TaggedValue
dim utvekslingsalias, codeListUrl
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
		
	codeListUrl = ""	
	for each tag in element.TaggedValues								
'		if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
		if LCase(tag.Name) = "sosi_bildeavmodellelement" and tag.Value <> "" then
			diagCounter = diagCounter + 1
			Session.Output("[caption=""Figur "&diagCounter&": "",title="&tag.Name&"]")
		'	Session.Output("image::"&tag.Value&".png["&ThePackage.Name"."&tag.Name&"]")
			Session.Output("image::"&tag.Value&"["&tag.Value&"]")
		end if
		if LCase(tag.Name) = "codelist" and tag.Value <> "" then
			codeListUrl = tag.Value
		end if
	next
end if
' testing http get
if codeListUrl <> "" then
'	Session.Output("DEBUG codeListUrl: " & codeListUrl & "")
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
			
			if mid(line,1,16) = "/gml:identifier>" and kodelistenavn = "" then
			' trigger utskrift
				'	første element
					Session.Output("Kodeliste hentet fra register: "&codeListUrl&"")	
					Session.Output(" ")	
					Session.Output("Kodeliste hentet på tidspunkt: "&nao()&"")
					Session.Output(" ")	
					Session.Output("Kodelistens navn i registeret: "&ualias&"")	
					Session.Output(" ")	


					Session.Output("===== Koder")
					Session.Output("[cols=""25,60,15""]")
					Session.Output("|===")
					Session.Output("|*Kodenavn:* ")
					Session.Output("|*Definisjon:* ")
					Session.Output("|*Utvekslingsalias:* ")
					Session.Output(" ")				
				
					kodelistenavn = ualias
			end if
			if mid(line,1,16) = "/gml:Definition>" and kodelistenavn <> "" then
			' trigger utskrift
					'koder
					Session.Output("|"&kodenavn&"")
					Session.Output("|"&kodedef&"")
					Session.Output("|"&ualias&"")					
			end if

		next
		Session.Output("|===")
	else
		Session.Output("Kodeliste kunne ikke hentes fra register: "&codeListUrl&"")	
		Session.Output(" ")		
'		Session.Output("DEBUG feil ved lesing av kodeliste: "&httpObject.status&"")
	end if
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
		Session.Output("|"&getCleanDefinition(att.Notes)&"")
		if not att.Default = "" then
			Session.Output(" ")
			Session.Output("|Utvekslingsalias: ")
			Session.Output("|"&att.Default&"")
		end if
		Session.Output("|===")
		' Hva med tagged values på koder? TBD
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
Dim textVar, skrivRoller

skrivRoller = false
'Session.Output("===== Roller")


'assosiasjoner
For Each con In element.Connectors
	If con.Type = "Association" or con.Type = "Aggregation" Then
'		Session.Output("[cols=""20,80""]")
'		Session.Output("|===")
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
				if skrivRoller = false then
					Session.Output("[cols=""20,80""]")
					Session.Output("|===")
					Session.Output("===== Roller")
					skrivRoller = true
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
				if skrivRoller = false then
					Session.Output("[cols=""20,80""]")
					Session.Output("|===")
					Session.Output("===== Roller")
					skrivRoller = true
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
		if skrivRoller = true then
			Session.Output("|===")
		end if
	End If
Next


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
				Session.Output("|*Subtyper:*")
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

'-----------------Funksjon for full path-----------------
function getPath(package)
	dim path
	dim parent
	if package.Element.Stereotype = "" then
		path = package.Name
	else
		path = "«" + package.Element.Stereotype + "» " + package.Name
	end if
	if not (ucase(package.Element.Stereotype)="APPLICATIONSCHEMA" or package.parentID = 0) then
		set parent = Repository.GetPackageByID(package.ParentID)
		path = getPath(parent) + "/" + path
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
