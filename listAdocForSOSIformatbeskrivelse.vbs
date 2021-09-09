Option Explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: listAdocForSOSIformatbeskrivelse
' Purpose: Genererer SOSI-formatbeskrivelse i AsciiDoc syntaks
'
' Version 0.3 2021-09-09 skriver ikke ut abstrakte klasser
' Version 0.2 2021-09-06 feilretting
' Version 0.1 2021-07-05 vise egenskapsnavn foran datatypeegenskapsnavn (informasjon.navnerom)
'						vise stereotypenavn foran datatyper og kodelister
'
'
' TBD: 
' TBD: opprydding !!!
'
Dim imgfolder, imgparent, parentimg
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
			Dim thePackage As EA.Package
			set thePackage = Repository.GetTreeSelectedObject()

			Call ListAsciiDoc(thePackage)
			Session.Output("// End of SOSI-format")
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


Session.Output("=== Pakke: "&thePackage.Name&"")
'stereotype TBD


For each element in thePackage.Elements
	If Ucase(element.Stereotype) = "FEATURETYPE" or Ucase(element.Stereotype) = "" Then
		if element.Abstract <> 1 then
			Call ObjektOgDatatyper(element)
		end if
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
Dim fil As EA.File
Dim supplier As EA.Element
Dim client As EA.Element

Dim textVar, addnotes, punktum
dim externalPackage

if element.Name <> "" then
	Session.Output(" ")
	Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
end if

if element.AttributesEx.Count > 0 then
	Session.Output("===== Modellelementnavn og SOSI-formatnavn") 
	Session.Output("[cols=""20,20,20,10""]")
	Session.Output("|===")
	Session.Output("|*Navn:* ")
	Session.Output("|*Type:* ")
	Session.Output("|*SOSI_navn:* ")
	Session.Output("|*Mult.:* ")
	Session.Output(" ")

	call listDatatype("", "..", element)	
	
	' kun roller? (vises ikke i Ex.Count)
	
	Session.Output("|===")

end if


End sub
'-----------------ObjektOgDatatyper End-----------------



'--------------------Start Sub-------------
sub listDatatype(egenskap, punktum, element)
Dim pktum, eskap, stereo, supereg, superpktum
Dim datatype As EA.Element
Dim att As EA.Attribute
dim super as EA.Element
dim conn as EA.Collection

for each conn in element.Connectors
	if conn.Type = "Generalization" then
		if element.ElementID = conn.ClientID then
			set super = Repository.GetElementByID(conn.SupplierID)
			supereg = egenskap
			superpktum = punktum
			call listDatatype(supereg,superpktum,super)
		end if
	end if
next
		
if element.Attributes.Count > 0 then
	for each att in element.Attributes
			stereo = ""
			if att.Type = "Punkt" or att.Type = "Kurve" or att.Type = "Flate" then
			' GM_Curve etc. TBD
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
				Session.Output("|"&stereo&att.Type&"")			
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

	next

end if


' skriv ut roller - sortert etter tagged value sequenceNumber TBD

for each conn in element.Connectors
		stereo = ""
		if conn.Type = "Association" then
		if element.ElementID = conn.ClientID then
			if conn.SupplierEnd.Role <> "" and conn.SupplierEnd.Navigable = "Navigable" then
				'
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
		else
			if conn.ClientEnd.Role <> "" and conn.ClientEnd.Navigable = "Navigable" then
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
		if stereo <> "" then	
			if LCase(datatype.Stereotype) <> "featuretype" then
				supereg = egenskap
				superpktum = punktum
				call listDatatype(supereg,superpktum,datatype)
			end if
		end if
	end if
next

end sub

'--------------------End Sub-------------




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
