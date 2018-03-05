option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		listSOSIKontrollfiler
' purpose:		Generate files for SOSI validator. Lager filer for SOSI-Kontroll fra SOSI-5.0 modeller.
' author:		Kent
' version:		2018-03-01

'	Skriver til fire filer i egen folder:
'		Liste med navnene på filer som skal inkluderes
'		Objekttypedefinisjoner inkludert nøstet innhold i alle datatyper, ikke arvede (?)
'		Objektutvalg som kobler objekttypenavnet til utvalgsgruppen
'		Elementbeskrivelser som angir datatype og lengde for enkeltelementene
		DIM sosFSO
		DIM sosFolder
		DIM defFile
		DIM objFile
		DIM utvFile
		DIM eleFile
		DIM DefTypes
		DIM def
		DIM obj
		DIM utv
		DIM ele
		DIM debug
		debug = false

sub listFeatureTypesForEnValgtPakke()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	DIM i
	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		if Repository.GetTreeSelectedItemType() = otPackage then
			'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
					dim message
			dim box
			box = Msgbox ("Skript listSOSIKontrollfiler" & vbCrLf & vbCrLf & "Skriptversjon 2018-03-01" & vbCrLf & "Starter listing til SOSIKontrollfiler for pakke : [" & theElement.Name & "].",1)
			select case box
			case vbOK
				dim kortnavn
				kortnavn = getPackageTaggedValue(theElement,"SOSI_kortnavn")
				if kortnavn = "" then
					kortnavn = theElement.Name
					Repository.WriteOutput "Script", "Pakken mangler tagged value SOSI_kortnavn! Kjører midlertidig videre med pakkenavnet som kortnavn: " & vbCrLf & kortnavn, 0
				end if
				kortnavn = InputBox("Velg produktets kortnavn.", "kortnavn", kortnavn)
				Repository.ClearOutput "Script"
				Repository.CreateOutputTab "Error"
				Repository.ClearOutput "Error"

				Set sosFSO=CreateObject("Scripting.FileSystemObject")
				if not sosFSO.FolderExists(kortnavn) then
					sosFSO.CreateFolder kortnavn
				end if
				'TBD to be version agnostic we must replace 50 here with the value in SOSI_versjon (except the dots)
				if not sosFSO.FolderExists(kortnavn & "\kap50") then
					sosFSO.CreateFolder kortnavn & "\kap50"
				end if
				defFile = kortnavn & "\" & "Def_" & getNCNameX(kortnavn) & ".50"
				objFile = kortnavn & "\kap50\" & getNCNameX(kortnavn) & "_o.50"
				utvFile = kortnavn & "\kap50\" & getNCNameX(kortnavn) & "_u.50"
				eleFile = kortnavn & "\kap50\" & getNCNameX(kortnavn) & "_d.50"
				Repository.WriteOutput "Script", Now & " sosFolder: " & kortnavn & " objFile: " & objFile & " utvFile: " & utvFile & " eleFile: " & eleFile, 0
				Set def = sosFSO.CreateTextFile(defFile,True,False)
				Set obj = sosFSO.CreateTextFile(objFile,True,False)
				Set utv = sosFSO.CreateTextFile(utvFile,True,False)
				Set ele = sosFSO.CreateTextFile(eleFile,True,False)
				'obj.Write".HODE"  & vbCrLf
				'obj.Write"..TEGNSETT UTF-8"  & vbCrLf & vbCrLf
				obj.Write"! *   " & utf8(kortnavn) & "   Objektdefinisjoner generert fra SOSI UML modell   " & now & "   *!"  & vbCrLf & vbCrLf
				'utv.Write".HODE"  & vbCrLf
				'utv.Write"..TEGNSETT UTF-8"  & vbCrLf & vbCrLf
				utv.Write"! *   " & utf8(kortnavn) & "   Utvalgsdefinisjoner generert fra SOSI UML modell   " & now & "   *!"  & vbCrLf & vbCrLf
				'ele.Write".HODE"  & vbCrLf
				'ele.Write"..TEGNSETT UTF-8"  & vbCrLf & vbCrLf
				ele.Write"! *   " & utf8(kortnavn) & "   Elementdefinisjoner generert fra SOSI UML modell   " & now & "   *!"  & vbCrLf & vbCrLf
				'call setBasicSOSITypes()
				'create global sosinametypelist
				Set DefTypes = CreateObject("System.Collections.ArrayList")


				call listFeatureTypes(theElement,kortnavn)
				
				utv.Write".GRUPPE-UTVALG Flateavgrensning" & vbCrLf
				utv.Write"..VELG ""..OBJTYPE"" = Flateavgrensning" & vbCrLf
				utv.Write"..BRUK-REGEL Flateavgrensning" & vbCrLf

				
				def.Write "[SyntaksDefinisjoner]" & vbCrLf
				def.Write "kap50\" & kortnavn & "_d.50" & vbCrLf
				def.Write "STD\SOSISTD.50" & vbCrLf & vbCrLf
				def.Write "[KodeForklaringer]" & vbCrLf
				def.Write "STD\KODER.50" & vbCrLf & vbCrLf
				def.Write "[UtvalgsRegler]" & vbCrLf
				def.Write "kap50\" & kortnavn & "_u.50" & vbCrLf & vbCrLf
				def.Write "[ObjektDefinisjoner]" & vbCrLf
				def.Write "kap50\" & kortnavn & "_o.50" & vbCrLf
				def.Write "STD\Flateavgrensning_o.50" & vbCrLf & vbCrLf
				'Til slutt språkelementdefinisjoner for alle enkeltelementer, med basistype: (...MÅLEMETODE T40")

				ele.Write vbCrLf & ".DEF" & vbCrLf & "..OBJTYPE T32" & vbCrLf

				for i = 0 To DefTypes.Count - 1
					'Repository.WriteOutput "Script", " DefTypes: " & DefTypes(i) & " index: " & i, 0
					ele.Write vbCrLf & ".DEF" & vbCrLf
					ele.Write DefTypes(i) & vbCrLf 
				next

				def.Close
				obj.Close
				utv.Close
				ele.Close
				' Release the file system object
				Set sosFSO = Nothing
				Repository.WriteOutput "Script", Now & " Filer skrevet til: " & kortnavn & ".",0

			case VBcancel

			end select
	

		Else
		  'Other than CodeList selected in the tree
		  MsgBox( "This script requires a package to be selected in the Project Browser." & vbCrLf & _
			"Please select a package in the Project Browser and try again." )
		end If
		'Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires a package to be selected in the Project Browser." & vbCrLf & _
	  "Please select a package in the Project Browser and try again." )
	end if
end sub


sub listFeatureTypes(pkg,kortnavn)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim super as EA.Element
	dim datatype as EA.Element
	dim conn as EA.Collection
 	set elements = pkg.Elements 
	dim i, sosinavn, sositype, sosilengde, sosimin, sosimax, koder, prikkniv, sosierlik
	for i = 0 to elements.Count - 1 
		dim currentElement as EA.Element 
		set currentElement = elements.GetAt( i ) 
				
		if currentElement.Type = "Class" and LCase(currentElement.Stereotype) = "featuretype" then
			
			if debug then Repository.WriteOutput "Script", kortnavn &";"&pkg.Name &";"&currentElement.Name &";"&getDefinitionText(currentElement),0

			if currentElement.ParentID <> 0 then
				if debug then Repository.WriteOutput "Script", " ParentID,ParentName :" & currentElement.ParentID & " - " & Repository.GetElementByID(currentElement.ParentID).Name,0
			end if
			utv.Write".GRUPPE-UTVALG " & utf8(currentElement.Name) & vbCrLf
			utv.Write"..VELG ""..OBJTYPE"" = " & utf8(currentElement.Name) & vbCrLf
			utv.Write"..BRUK-REGEL " & utf8(currentElement.Name) & vbCrLf & vbCrLf
			
			obj.Write vbCrLf & ".OBJEKTTYPE" & vbCrLf
			obj.Write"..TYPENAVN " & utf8(currentElement.Name) & vbCrLf
			if getSosiGeometrityper(currentElement) <> "" then
				obj.Write"..GEOMETRITYPE " & getSosiGeometrityper(currentElement) & vbCrLf
			else
				obj.Write"..GEOMETRITYPE PUNKT,SVERM,KURVE,FLATE,OBJEKT" & vbCrLf
			end if
			' restriksjon? -> ..AVGRENSES_AV KantUtsnitt,TakoverbyggKant,FiktivBygningsavgrensning(,Flateavgrensning?)
			for each conn in currentElement.Connectors
				if conn.Type = "Generalization" then
					if currentElement.ElementID = conn.ClientID then
						set super = Repository.GetElementByID(conn.SupplierID)
						obj.Write"..INKLUDER " & utf8(super.Name) & vbCrLf
						' må vi nøste helt opp og ta med alle? (..INKLUDER Ngis KommMå)
					end if
				end if
			next

			'obj.Write"..PRODUKTSPEK " & utf8(kortnavn) & " " & utf8(getPackageTaggedValue(pkg,"SOSI_versjon")) & vbCrLf
			obj.Write"..EGENSKAP """" * ""..OBJTYPE""    T32  1  1  = (" & utf8(currentElement.Name) & ")" & vbCrLf

			ele.Write vbCrLf & "! " & utf8(currentElement.Name) & vbCrLf

			prikkniv = ".."
			call listDatatypes(currentElement,prikkniv)

		end if
	
	next

	dim subP as EA.Package
	for each subP in pkg.packages
	    call listFeatureTypes(subP,kortnavn)
	next


end sub


sub listDatatypes(element,prikkniv)
	dim presentasjonsnavn
 	dim elements as EA.Collection 
	dim super as EA.Element
	dim datatype as EA.Element
	dim conn as EA.Collection
	dim connEnd as EA.ConnectorEnd
	dim i, umlnavn, sosinavn, sositype, sosilengde, sosimin, sosimax, sosierlik, koder, prikkniv1, roleEndElementID, sosidef
				
	if element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "featuretype" then

		dim attr as EA.Attribute
		for each attr in element.Attributes
			sosinavn = ""
				sosinavn = getTaggedValue(attr,"SOSI_navn")
			if attr.ClassifierID <> 0 then 
				if debug then Repository.WriteOutput "Script", "Debug: attr.Name [" & attr.Name & "] SOSI_navn [" & getTaggedValue(attr,"SOSI_navn") & "].",0
			else
				if debug then Repository.WriteOutput "Script", "Debug: attr.Name [" & attr.Name & "] no ClassifierID.",0
			end if
			if getSosiGeometritype(attr) = "" then
				if debug then Repository.WriteOutput "Script", "Debug: attr.Name [" & attr.Name & "] not geometry.",0
				sosimin = "0"
				sosimax = "N"
				sosierlik = "><"
				if attr.LowerBound = "1" then
					sosimin = "1"
				end if
				if attr.UpperBound = "1" then
					sosimax = "1"
				end if
				sositype = UCase(getTaggedValue(attr,"SOSI_type")) 
				if sositype = "" then
					sositype = "*"
					sositype = getBasicSOSIType(attr.Type)
				end if
				sosilengde = getTaggedValue(attr,"SOSI_lengde")
				if sositype <> "*" and sosilengde <> "" then
					sositype = sositype & sosilengde
				end if
				koder = ""
				'Initialverdi+frozen på basistyper?
				'Kodelisteegenskap
				if attr.ClassifierID <> 0 then
					set datatype = Repository.GetElementByID(attr.ClassifierID)
					if sosinavn = "" then
						sosinavn = getTaggedValue(datatype,"SOSI_navn")
					end if
					if sosinavn = "" then
						sosinavn = attr.Name
					end if
					if sositype = "*" and getTaggedValue(datatype,"SOSI_type") <> "" then
						' Er denne riktig dersom gamle kodelister har egne sosityper eller er alle kodelister av type T? (samme med sosilengde?) TBD
						sositype = getTaggedValue(datatype,"SOSI_type")
					end if
					if datatype.Type = "Class" and LCase(datatype.Stereotype) = "codelist" or LCase(datatype.Stereotype) = "enumeration" then
						'Repository.WriteOutput "Script", "Debug: Repository.GetElementByID(attr.ClassifierID).Name [" & Repository.GetElementByID(attr.ClassifierID).Name & "] kodelistens SOSI_navn [" & getTaggedValue(Repository.GetElementByID(attr.ClassifierID),"SOSI_navn") & "].",0
						if debug then Repository.WriteOutput "Script", "Debug: kodeliste.Name [" & datatype.Name & "] kodelistens SOSI_navn [" & getTaggedValue(datatype,"SOSI_navn") & "].",0
						if sositype = "*" then
							sositype = "T"
						end if
						koder = getKoder(datatype)
						if koder <> "" then
							sosierlik = "="
						end if
					end if
				else
					'ukjent basistype ?
					if sositype = "*" then
						sositype = "T666"
					end if
					
				end if
				obj.Write"..EGENSKAP """ & utf8(attr.Name) & """ " & utf8(attr.Type) & " """ & prikkniv & utf8(sosinavn) & """ " & utf8(sositype) & " " & sosimin & " " & sosimax & "  " & sosierlik & " (" & utf8(koder) & ")" & vbCrLf
			
				if prikkniv <> ".." or sositype = "*" then 
					'if  prikkniv = ".." and ! sositypeWritten(sositype) then
						if  prikkniv = ".." and sositype = "*" then 
							ele.Write vbCrLf & ".DEF" & vbCrLf
							' sositypeWritten.Add(sositype)
						end if
						ele.Write prikkniv & utf8(sosinavn) & " " & utf8(sositype) & vbCrLf
					'end if
				end if
				'Putt i liste over enkeltelementer med basistype som skal listes opp separate til slutt: TBD
				if attr.ClassifierID <> 0 then
					'Brukerdefinert datatype
					if datatype.Type = "Class" and LCase(datatype.Stereotype) = "datatype" or LCase(datatype.Stereotype) = "union" then
						if debug then Repository.WriteOutput "Script", "Debug: datatype.Name [" & datatype.Name & "] datatypens SOSI_navn [" & getTaggedValue(datatype,"SOSI_navn") & "].",0
						'             set datatype = Repository.GetElementByID(attr.ClassifierID)
						prikkniv1 = prikkniv & "."
						call listDatatypes(datatype,prikkniv1)
					end if
				end if
				'Kompaktifisering????
				'Liste over kjente elementer med kompaktifisering???
				
				sosidef = prikkniv & sosinavn & " " & sositype
				if sositype <> "*" then 
					if DefTypes.IndexOf(sosidef,0) = -1 then	
						' 	ikke i lista, legges inn
						DefTypes.Add sosidef
					end if
				end if
			end if
		
		next
			
		for each conn in element.Connectors
			if conn.Type = "Generalization" or conn.Type = "Realisation" or conn.Type = "NoteLink" then

			else
				'Repository.WriteOutput "Script", "Debug: Supplier Role.Name [" & conn.SupplierEnd.Role & "] datatypens SOSI_navn [" & getTaggedValue(Repository.GetElementByID(conn.ClientID).Name,"SOSI_navn") & "].",0
				'Repository.WriteOutput "Script", "Debug: Client Role.Name [" & conn.ClientEnd.Role & "] datatypens SOSI_navn [" & getTaggedValue(Repository.GetElementByID(conn.ClientID).Name,"SOSI_navn") & "].",0
				if debug then Repository.WriteOutput "Script", "Debug: Supplier Role.Name [" & conn.SupplierEnd.Role & "] datatypens SOSI_navn [" & Repository.GetElementByID(conn.SupplierID).Name & "].",0
				if debug then Repository.WriteOutput "Script", "Debug: Client Role.Name [" & conn.ClientEnd.Role & "] datatypens SOSI_navn [" & Repository.GetElementByID(conn.ClientID).Name & "].",0
				sositype = "REF"
				sosimin = "0"
				sosimax = "N"
				sosierlik = "><"
				koder = ""
				if conn.ClientID = element.ElementID then
					set datatype = Repository.GetElementByID(conn.SupplierID)
					umlnavn = conn.SupplierEnd.Role
					sosinavn = getConnectorEndTaggedValue(conn.SupplierEnd,"SOSI_navn")
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
					umlnavn = conn.ClientEnd.Role
					sosinavn = getConnectorEndTaggedValue(conn.ClientEnd,"SOSI_navn")
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
					if sosinavn = "" then
						sosinavn = umlnavn
					end if
					obj.Write"..EGENSKAP """ & utf8(umlnavn) & """ " & utf8(datatype.Name) & " """ & prikkniv & utf8(sosinavn) & """ " & utf8(sositype) & " " & sosimin & " " & sosimax & "  " & sosierlik & " (" & koder & ")" & vbCrLf

					sosidef = prikkniv & sosinavn & " REF"
					'if sositype <> "*" then 
						if DefTypes.IndexOf(sosidef,0) = -1 then	
							' 	ikke i lista, legges inn
							DefTypes.Add sosidef
						end if
					'end if
					
					'if composition2datatype then
						'Brukerdefinert datatype
						if datatype.Type = "Class" and LCase(datatype.Stereotype) = "datatype" or LCase(datatype.Stereotype) = "union" then
							if debug then Repository.WriteOutput "Script", "Debug: datatype.Name [" & datatype.Name & "] datatypens SOSI_navn [" & getTaggedValue(datatype,"SOSI_navn") & "].",0
							'             set datatype = Repository.GetElementByID(attr.ClassifierID)
							sositype = "*"
							if prikkniv <> ".." or sositype = "*" then 
								'if  prikkniv = ".." and ! sositypeWritten(sositype) then
									if  prikkniv = ".." and sositype = "*" then 
										ele.Write vbCrLf & ".DEF" & vbCrLf
										' sositypeWritten.Add(sositype)
									end if
									ele.Write prikkniv & utf8(sosinavn) & " " & utf8(sositype) & vbCrLf
								'end if
							end if
							
							prikkniv1 = prikkniv & "."
							call listDatatypes(datatype,prikkniv1)
						end if
				end if

				
			end if

		next

	end if

end sub

function getKoder(element)
		dim kode, koder
		koder = ""
		dim attr as EA.Attribute
		for each attr in element.Attributes
			'if debug then Repository.WriteOutput "Script", "Debug: code.Name [" & attr.Name & "] SOSI_navn [" & getTaggedValue(attr,"SOSI_navn") & "].",0
			kode = utf8(attr.Name)
			if getTaggedValue(attr,"SOSI_verdi") <> "" then
				kode = utf8(getTaggedValue(attr,"SOSI_verdi"))
			end if
			if attr.Default <> "" then
				kode = utf8(attr.Default)
			end if
			
			koder = koder & "," & kode
		next
		if Len(koder) < 2 then
			getKoder = ""
		else
			getKoder = Mid(koder,2,Len(koder))
		end if
end function

function getSosiGeometrityper(element)
		dim typer
		typer = ""
		dim attr as EA.Attribute
		for each attr in element.Attributes
			if getSosiGeometritype(attr) <> "" then
				typer = typer & "," & getSosiGeometritype(attr)
			end if
		next
		if Len(typer) < 2 then
			getSosiGeometrityper = ""
		else
			getSosiGeometrityper = Mid(typer,2,Len(typer))
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

function getNCNameX(str)
	' make name legal SOSI-Kontrolfil (NC+ingen punktum)
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
		  if tegn = "." Then
			  'Repository.WriteOutput "Script", "Bad0 in SOSI-kontrollfil: " & tegn,0
			  u=1
		  Else
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
		  End If
		Next
		' return res
		getNCNameX = res

End function

function getDefinitionText(currentElement)

    Dim txt, res, tegn, i, u
    u=0
	getDefinitionText = ""
		txt = Trim(currentElement.Notes)
		res = ""
		' loop gjennom alle tegn
		For i = 1 To Len(txt)
		  tegn = Mid(txt,i,1)
		  If tegn = ";" Then
			  res = res + " "
		  Else 
			If tegn = """" Then
			  res = res + "'"
			Else
			  If tegn < " " Then
			    res = res + " "
			  Else
			    res = res + Mid(txt,i,1)
			  End If
			End If
		  End If
		  
		Next
		
	getDefinitionText = res

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

listFeatureTypesForEnValgtPakke
