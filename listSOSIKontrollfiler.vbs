option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		listSOSIKontrollfiler
' purpose:		Generate files for SOSI validator. Lager filer for SOSI-Kontroll fra SOSI-5.0 modeller.
' author:		Kent
'
'
' version:		2022-01-25 legger ut koder på gammel måte, koder hentes fra sti til eksterne kodelister
'				NB må kjøres pånytt for hver endring i registeret ! og kanskje synkroniseres med når nye sosikontrollfiler distribueres
' version:		2022-01-19 sti til kodelister hentes alternativt fra kodelisteklassen hvis defaultCodeSpace mangler
' version:		2022-01-04 URI mappes til lovlig SOSI-type T, uppercase på SOSI-navn der tV SOSI_navn mangler (roller)
' version:		2021-12-23 bruker stien i defalutCodeSpace uansett innhold, fullsti, retting for egenskaper som arver geometrityper
' version:		2019-08-02 feilrettet og forbedret (?) støtte for arv mellom datatyper
' version:		2019-07-31 utelater SOSI-kontroll av elementer med tagged value xsdEncodingRule = notEncoded
'    			Objekttyper som manglende geometriegenskap legges under .OBJEKT
' version:		2019-04-03 retta feil vedr. SOSI_lengde og SOSI_datatype, datatyper arver fra sine supertyper
' version:		2019-03-20 for egenskaper med defaultCodeSpace ansees kodelisten som eksternt forvaltet og tom(!), og stien til den eksterne lista legges rett i o-fila .
' version:		2018-09-07 self associations and self datatypes are detected, not other types of circular usages or inheritance loops.
' version:		2018-03-01, 04.13
'
'	Kjente svakheter:
'	Hvis filene ikke dukker opp i en underkatalog under der .eap-fila ligger kan man lete på en tempsti ala:
'	C:\Users\jonken\AppData\Roaming\Sparx Systems\EA\Temp\Elveg\kap50
'
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
	DIM i, fullsti
	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		if Repository.GetTreeSelectedItemType() = otPackage then
			'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
					dim message
			dim box
			box = Msgbox ("Skript listSOSIKontrollfiler" & vbCrLf & vbCrLf & "Skriptversjon 2022-01-25" & vbCrLf & "Starter listing av modell som følger SOSI 5.0-regler til SOSIKontrollfiler for pakke : [" & theElement.Name & "].",1)
			select case box
			case vbOK
				dim kortnavn
				'tømmer System Output for lettere å fange opp eventuelle feilmeldinger der
				Repository.ClearOutput "Script"
				Repository.CreateOutputTab "Error"
				Repository.ClearOutput "Error"
				kortnavn = getPackageTaggedValue(theElement,"SOSI_kortnavn")
				if kortnavn = "" then
					kortnavn = theElement.Name
					Repository.WriteOutput "Script", "Pakken mangler tagged value SOSI_kortnavn! Kjører midlertidig videre med pakkenavnet som forslag til kortnavn: " & vbCrLf & kortnavn, 0
				end if
				kortnavn = InputBox("Velg produktets kortnavn.", "kortnavn", kortnavn)
				Repository.WriteOutput "Script", Now & "Starter listing til SOSIKontrollfiler for pakke : [" & theElement.Name & "], valgt SOSI_kortnavn: " & vbCrLf & kortnavn, 0

				Set sosFSO=CreateObject("Scripting.FileSystemObject")
				fullsti = sosFSO.GetParentFolderName(Repository.ConnectionString())				
				if not sosFSO.FolderExists(fullsti & "/" & kortnavn) then
					sosFSO.CreateFolder fullsti & "/" & kortnavn
				end if
				'TBD to be version agnostic we must replace 50 here with the value in SOSI_versjon (except the dots)
				if not sosFSO.FolderExists(fullsti & "/" & kortnavn & "\kap50") then
					sosFSO.CreateFolder fullsti & "/" & kortnavn & "\kap50"
				end if
				defFile = fullsti & "/" & kortnavn & "\" & "Def_" & getNCNameX(kortnavn) & ".50"
				objFile = fullsti & "/" & kortnavn & "\kap50\" & getNCNameX(kortnavn) & "_o.50"
				utvFile = fullsti & "/" & kortnavn & "\kap50\" & getNCNameX(kortnavn) & "_u.50"
				eleFile = fullsti & "/" & kortnavn & "\kap50\" & getNCNameX(kortnavn) & "_d.50"
				'Repository.WriteOutput "Script", Now & " sosFolder: " & kortnavn & " objFile: " & objFile & " utvFile: " & utvFile & " eleFile: " & eleFile, 0
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
				Repository.WriteOutput "Script", Now & " Filer skrevet til katalogen: " & kortnavn & ".",0

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
	dim i, sosinavn, sositype, sosilengde, sosimin, sosimax, koder, prikkniv, sosierlik, superlist
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
			' hardkodet overstyring av mappingregler ?
			if getTaggedValue(currentElement,"SOSI_geometri") <> "" then
				obj.Write"..GEOMETRITYPE " & getTaggedValue(currentElement,"SOSI_geometri") & vbCrLf
			else
				if getSosiGeometrityper(currentElement) <> "" then
					obj.Write"..GEOMETRITYPE " & getSosiGeometrityper(currentElement) & vbCrLf
				else
					'obj.Write"..GEOMETRITYPE PUNKT,SVERM,KURVE,FLATE,OBJEKT" & vbCrLf
					obj.Write"..GEOMETRITYPE OBJEKT" & vbCrLf
				end if
			end if
			' restriksjon? -> ..AVGRENSES_AV KantUtsnitt,TakoverbyggKant,FiktivBygningsavgrensning(,Flateavgrensning?)
			superlist = ""
			for each conn in currentElement.Connectors
				if conn.Type = "Generalization" then
					if currentElement.ElementID = conn.ClientID then
						' må nøste helt opp og ta med alle (..INKLUDER Ngis KommMå)
						superlist = getSupertypes(conn.SupplierID)
						obj.Write"..INKLUDER " & utf8(superlist) & vbCrLf
						'set super = Repository.GetElementByID(conn.SupplierID)
						'obj.Write"..INKLUDER " & utf8(super.Name) & vbCrLf
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
	dim sconn as EA.Collection
	dim connEnd as EA.ConnectorEnd
	dim i, umlnavn, sosinavn, sositype, sosilengde, sosimin, sosimax, sosierlik, koder, prikkniv1, roleEndElementID, sosidef, codelistUrl
				
	if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union" or LCase(element.Stereotype) = "featuretype") then


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
				if attr.LowerBound = "1" and LCase(element.Stereotype) <> "union" then
					sosimin = "1"
				end if
				if attr.UpperBound = "1" or LCase(element.Stereotype) = "union" then
					sosimax = "1"
				end if
				sositype = UCase(getTaggedValue(attr,"SOSI_datatype")) 
				'if sositype = "" then
					' skal SOSI_datatype kunne overkjøre basistype? (Nei)
					sositype = "*"
					sositype = getBasicSOSIType(attr.Type)
				'end if
				sosilengde = getTaggedValue(attr,"SOSI_lengde")
				if sositype <> "*" and sosilengde <> "" then
					if sositype = "T" or sositype = "H" or sositype = "D" then 
						'ignorer lengdeangivelse på noen basistype, legger inn på andre. Hva med kodelistekoder som er T ??
						sositype = sositype & sosilengde
					end if
				end if
				koder = ""
				'Initialverdi+frozen på basistyper?
				'Kodelisteegenskap
				if attr.ClassifierID <> 0 then
					set datatype = Repository.GetElementByID(attr.ClassifierID)
					'test om navnet etter egenskapen er det samme som navnet på den refererte klassen TODO
					
					if sosinavn = "" then
						sosinavn = getTaggedValue(datatype,"SOSI_navn")
					end if
					if sosinavn = "" then
						sosinavn = attr.Name
					end if
					if sositype = "*" and getTaggedValue(datatype,"SOSI_datatype") <> "" then
						' Er denne riktig dersom gamle kodelister har egne sosityper eller er alle kodelister av type T? (samme med sosilengde?) TBD
						'sositype = getTaggedValue(datatype,"SOSI_datatype")
					end if
					if datatype.Type = "Enumeration" or ( datatype.Type = "Class" and LCase(datatype.Stereotype) = "codelist" or LCase(datatype.Stereotype) = "enumeration" ) then
						'Repository.WriteOutput "Script", "Debug: Repository.GetElementByID(attr.ClassifierID).Name [" & Repository.GetElementByID(attr.ClassifierID).Name & "] kodelistens SOSI_navn [" & getTaggedValue(Repository.GetElementByID(attr.ClassifierID),"SOSI_navn") & "].",0
						if debug then Repository.WriteOutput "Script", "Debug: kodeliste.Name [" & datatype.Name & "] kodelistens SOSI_navn [" & getTaggedValue(datatype,"SOSI_navn") & "].",0
						if sositype = "*" then
							sositype = "T"
						end if
						sosierlik = "="
						if getTaggedValue(attr,"defaultCodeSpace") <> "" then
							if getTaggedValue(attr,"defaultCodeSpace") <> getTaggedValue(datatype,"codeList") then
								Repository.WriteOutput "Script", "Info: egenskapen [" & Element.Name & "." & attr.Name & "]  har kodelistesti [" & getTaggedValue(attr,"defaultCodeSpace") & "]  men datatypeklassen [" & datatype.Name & "] har ulik kodelistesti [" & getTaggedValue(datatype,"codeList") & "].",0
							end if
							if getTaggedValue(datatype,"asDictionary") <> "true" then
								Repository.WriteOutput "Script", "Info: kodelista [" & datatype.Name & "]  har ikke en tagged value asDictionary satt til true.",0
							end if

'							koder = getTaggedValue(attr,"defaultCodeSpace")
'							sosierlik = "><"
							codelistUrl = getTaggedValue(attr,"defaultCodeSpace")
							koder = getAlleKoder(codelistUrl)
						else
							if getTaggedValue(datatype,"codeList") <> "" then
'								koder = getTaggedValue(datatype,"codeList")
'								sosierlik = "><"
								codelistUrl = getTaggedValue(datatype,"codeList")
								koder = getAlleKoder(codelistUrl)
							else
								' legger ut kodene som vanlig
								koder = getKoder(datatype)
							end if
						end if
						if koder = "" then
							sosierlik = "><"
						end if
					end if
				else
					'ukjent basistype ?
					if sositype = "*" then
						sositype = "FIX TYPE"
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
					if datatype.Type = "Datatype" or (datatype.Type = "Class" and LCase(datatype.Stereotype) = "datatype" or LCase(datatype.Stereotype) = "union") then
						'set datatype = Repository.GetElementByID(attr.ClassifierID)
						'
						'rekurser og hent egenskaper og roller fra eventuelle supertyper til datatypen TODO
						for each sconn in datatype.Connectors
							if debug then Repository.WriteOutput "Script", "Debug: sconn.Type [" & sconn.Type & "] sconn.ClientID [" & sconn.ClientID & "] sconn.SupplierID [" & sconn.SupplierID & "].",0
							if sconn.Type = "Generalization" then
								if datatype.ElementID = sconn.ClientID then
									if debug then Repository.WriteOutput "Script", "Debug: datatype har supertype [" & Repository.GetElementByID(sconn.SupplierID).Name & "].",0
									set super = Repository.GetElementByID(sconn.SupplierID)
									prikkniv1 = prikkniv & "."
	'								call listDatatypes(super,prikkniv1)
								end if
							end if
						next
						
						
						
						if debug then Repository.WriteOutput "Script", "Debug: datatype.Name [" & datatype.Name & "] datatypens SOSI_navn [" & getTaggedValue(datatype,"SOSI_navn") & "].",0
						'             set datatype = Repository.GetElementByID(attr.ClassifierID)
						prikkniv1 = prikkniv & "."
						if datatype.Name = element.Name then
							Repository.WriteOutput "Script", "Error - circular self reference: datatype.Name [" & datatype.Name & "] from attribute name [" & element.Name & "." & attr.Name & "].",0
						else
							call listDatatypes(datatype,prikkniv1)
						end if
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
					if getConnectorEndTaggedValue(conn.SupplierEnd,"xsdEncodingRule") = "notEncoded" then
						umlnavn = ""
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
					if getConnectorEndTaggedValue(conn.ClientEnd,"xsdEncodingRule") = "notEncoded" then
						umlnavn = ""
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
							if datatype.Name = element.Name then
								Repository.WriteOutput "Script", "Error - circular self reference: datatype.Name [" & datatype.Name & "] from role name [" & element.Name & "." & umlnavn & "].",0
							else
								call listDatatypes(datatype,prikkniv1)
							end if
						else
							' association may be to a feature type
						end if
				end if

				
			end if

		next




		for each conn in element.Connectors
			if conn.Type = "Generalization" then
				if element.ElementID = conn.ClientID then
					' må nøste helt opp og ta med alle "inline"
					if debug then Repository.WriteOutput "Script", "Debug: supertype [" & Repository.GetElementByID(conn.SupplierID).Name & "].",0
'					superlist = getSupertypes(ftname, conn.SupplierID, indent)
					if element.Type = "Datatype" or (element.Type = "Class" and LCase(element.Stereotype) = "datatype" or LCase(element.Stereotype) = "union") then
						set super = Repository.GetElementByID(conn.SupplierID)
						'prikkniv = prikkniv & "."
						call listDatatypes(super,prikkniv)
					end if
				end if
			end if
		next



	end if

end sub

function getSupertypes(elementID)
	dim super as EA.Element
	dim conn as EA.Collection
	dim supername, supernames
	set super = Repository.GetElementByID(elementID)
	for each conn in super.Connectors
		if conn.Type = "Generalization" then
			if super.ElementID = conn.ClientID then
				supername = getSupertypes(conn.SupplierID)
			end if
		end if
	next
	if debug then Repository.WriteOutput "Script", "Debug: super.Name [" & super.Name & "]  supername [" & supername & "].",0
	getSupertypes = super.Name & " "  & supername
end function

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



function getAlleKoder(codeListUrl)
	getAlleKoder = ""
	Dim codelist
	codelist = ""
	dim code1
	dim code2 
	code1 = ""
	code2 = ""
	' testing http get
	if codeListUrl <> "" then
	'	Session.Output("<!-- DEBUG codeListUrl: " & codeListUrl & " -->")
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
					if code1 = "" then
						code1 = ualias
					else
						if code2 = "" then
							code2 = ualias
							getAlleKoder = getAlleKoder + ualias
						else
							getAlleKoder = getAlleKoder + "," + ualias		
						end if
					end if
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
				



				ualias = ""
								
			next
	'		Session.Output("|===")
		else
	'		Session.Output("Kodeliste kunne ikke hentes fra register: "&codeListUrl&"")	
	'		Session.Output(" ")		
			if debug then Session.Output("<!-- DEBUG feil ved lesing av kodeliste: ["&codeListUrl&"] status:["&httpObject.status&"]-->")
		end if
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
		if getSosiGeometrityper = "" then
		'	arver geometritype?
			if getParentId(element) <> 0 then
				dim super as EA.Element
				set super = Repository.GetElementByID(getParentId(element))
				getSosiGeometrityper = getSosiGeometrityper(super)
			end if
		end if
		
end function


function getParentId(element)
	getParentId = 0
	dim conn as EA.Collection
	for each conn in element.Connectors
		if conn.Type = "Generalization" then
			if element.ElementID = conn.ClientID then
				getParentId = conn.SupplierID
			end if
		end if
	next
end function


function getSosiGeometritype(attr)
		'fra Ralisering i SOSI-format versjon 5.0 tabell 8.2:
		getSosiGeometritype = ""
		if attr.Type = "Punkt" or attr.Type = "GM_Point" then
			getSosiGeometritype = "PUNKT"
		end if
		if attr.Type = "Sverm" or attr.Type = "GM_MultiPoint" then
			getSosiGeometritype = "SVERM"
		end if
		if attr.Type = "Kurve" or attr.Type = "GM_Curve" or attr.Type = "GM_CompositeCurve" then
			getSosiGeometritype = "KURVE,BUEP,KLOTOIDE"
		end if
		if attr.Type = "Flate" or attr.Type = "GM_Surface" or attr.Type = "GM_CompositeSurface" then
			getSosiGeometritype = "FLATE"
		end if
		'fra "etablert praksis"
		if attr.Type = "GM_Object" or attr.Type = "GM_Primitive" then
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
	if UCase(umltype) = "URI" then
		getBasicSOSIType = "T"
	end if
	if UCase(umltype) = "Any" then
		getBasicSOSIType = "T"
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
