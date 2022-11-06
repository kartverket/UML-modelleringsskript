option explicit

!INC Local Scripts.EAConstants-VBScript

 ' Script Name: listKoderTilCSV
 ' Author: Kent Jonsrud Kartverket
 ' Purpose: list code names in system output
 ' Date: 2022-11-06 nasjonale tegn i notefeltet
 ' Date: 2022-11-04 superforenklet, faker utvekslingsalias hvis de ikke finnes - TODO
 ' Date: 2022-11-03 la inn en ny kolonne på slutten med mulighet for skos:broader (Mer generelt begrep)
 ' Date: 2022-11-02 Smårettinger, hvis Utgått finnes i navnet settes status til Utgått
 ' Date: 2021-10-01 Navn;Kodeverdi;Eier;Status;Oppdatert;Versjons ID;Beskrivelse;Gyldig fra;Gyldig til;ID
 ' Date: 2021-09-29 Endra ledetekster 
 ' Date: 2020-11-21
'
 '
 ' TBD: Kodenavn;Utvekslingsalias;Eier;Status;oppdateringsdato;Versjon;Definisjon;Gyldig fra;Gyldig til;ID;BredereID

	Dim groupsList
	Dim maingroupsList
	Set groupsList = CreateObject( "System.Collections.Sortedlist" )
	Set maingroupsList = CreateObject( "System.Collections.Sortedlist" )

sub oppdaterKoderForEnValgtKodeliste()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
	Repository.CreateOutputTab "Error"
	Repository.ClearOutput "Error"

	'Get the currently selected CodeList in the tree to work on

	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()

	
	if not theElement is nothing  then
		if (theElement.ObjectType = otElement) then
			dim box, mess
			mess = 	"List codelist codes to CSV. Script version 2022-11-06." & vbCrLf
			mess = mess + "Element: "& vbCrLf & "[«" & theElement.Stereotype & "» " & theElement.Name & "]."

			box = Msgbox (mess, vbOKCancel)
			select case box
			case vbOK
				if theElement.Type="Class" and LCase(theElement.Stereotype) = "codelist" Or LCase(theElement.Stereotype) = "enumeration" then
					'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
					fixOldCodelists(theElement)
				Else
					'Other than CodeList selected in the tree
					MsgBox( "This script requires a class with stereotype «CodeList» to be selected in the Project Browser." & vbCrLf & _
					"Please select a  CodeList class in the Project Browser and try once more." )
				end If
			'	Repository.WriteOutput "Script", "-----------------------" & Now & " Finished, check the Error and Types tabs", 0
				Repository.EnsureOutputVisible "Script"
			case VBcancel
						
			end select 
		else
			MsgBox( "This script requires a CodeList class element to be selected in the Project Browser." & vbCrLf & _
			"Please select a  CodeList class in the Project Browser and try once more." )
		end if
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires some CodeList class to be selected in the Project Browser." & vbCrLf & _
	  "Please select a  CodeList class in the Project Browser and try again." )
	end if

end sub



sub fixOldCodelists(el)
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
'	Repository.WriteOutput "Script", "-----------------------" & Now & " Script listKoder running on CodeList: " & el.Name, 0
	'Repository.WriteOutput "Script", Now & " " & el.Stereotype & " " & el.Name, 0
	'Repository.WriteOutput "Script", "id, kode, definisjon, initialverdi, SOSI_verdi, SOSI_presentasjonsnavn",0
    
	dim id
	id = 1
	call fillGroups()
	call fillMaingroups()

	'''Repository.WriteOutput "Script",  "Navn;Kodeverdi;Eier;Status;Oppdatert;Versjons ID;Beskrivelse;Gyldig fra;Gyldig til;ID;Mer generelt begrep" ,0
'forenklet:	'''Repository.WriteOutput "Script",  "Navn;Beskrivelse;Kodeverdi;Status;Gyldig fra;Gyldig til" ,0
'forenklet2:	''Navn;Beskrivelse;Kodeverdi;Navn engelsk;beskrivelse engelsk;kodeverdi engelsk;Status;Gyldig fra;Gyldig til
'	Repository.WriteOutput "Script",  "Navn;Beskrivelse;Kodeverdi;Status;Gyldig fra;Gyldig til" ,0
	' forenklet
	Repository.WriteOutput "Script",  "Navn;Beskrivelse;Kodeverdi;Navn engelsk;beskrivelse engelsk;kodeverdi engelsk;Status;Gyldig fra;Gyldig til" ,0


dim attr as EA.Attribute
	for each attr in el.Attributes
	'	Repository.WriteOutput "Script", Now & " " & el.Name & "." & attr.Name, 0

	'If the old codes are in a form suitable as initial draft of proper definition:
	'kopierKodensNavnTilTomDefinisjon(attr)
	'remember to refrase these definitions according to the rules for developing definitions

	' kopierKodensNavnTilTagSOSI_presentasjonsnavn(attr)

    'flyttInitialverdiTilTagSOSI_verdi(attr)

    'settKodensNavnTilNCName(attr)
    'eller
    'settKodensNavnTilEgen_Navn(attr)

   Call listKodeliste(el, id, attr)
   'id = id + 1

	'listKodeliste(attr)

	next

end sub



Sub listKodeliste(codelist, id, attr)

	dim kodenavn, definisjon, SOSI_verdi, SOSI_presentasjonsnavn, navnerom, gyldig, broader, utvekslingsalias
	'dim codelist
	if getTaggedValue(codelist,"codeList") <> "" then
		navnerom = getTaggedValue(codelist,"codeList")
	else
		navnerom = "http://skjema.geonorge.no/sosi/basistype"
	end if
	kodenavn = attr.Name
	definisjon = attr.Notes
	definisjon = Replace(definisjon,"""","")
	'check if the element has a tagged value with the required name
		dim existingTaggedValue AS EA.TaggedValue
		'dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		SOSI_presentasjonsnavn = ""
		SOSI_verdi = ""
		utvekslingsalias = attr.Default
		for taggedValuesCounter = 0 to attr.TaggedValues.Count - 1
			set existingTaggedValue = attr.TaggedValues.GetAt(taggedValuesCounter)

			if existingTaggedValue.Name = "SOSI_verdi" then
				SOSI_verdi = existingTaggedValue.Value
			end if
			if existingTaggedValue.Name = "SOSI_presentasjonsnavn" then
				SOSI_presentasjonsnavn = existingTaggedValue.Value
			end if
		next
'   		Repository.WriteOutput "Script", "INSERT INTO databaseskjema." & codelist & " VALUES (" & id & ",'"& kodenavn & "','" & attr.Default & "','" & ikkelinjeskiftEllerEnkeltfnutt(definisjon) & "','" & SOSI_verdi & "','" & SOSI_presentasjonsnavn & "');",0
	if InStr(kodenavn,"Utgått") then
		gyldig="Utgått"
	else
		gyldig="Gyldig"
	end if

	broader = getBroader(codelist,attr)
	if utvekslingsalias = "" then
		utvekslingsalias = kodenavn
	end if
	'''Repository.WriteOutput "Script",  kodenavn &";"& attr.Default &";SOSI;"&gyldig&";"& date &";0.1;"& ikkelinjeskiftEllerEnkeltfnutt(attr.Notes) &";;;"& navnerom & "/" & attr.Name &";"&broader ,0
'	Repository.WriteOutput "Script",  kodenavn&";"&ikkelinjeskiftEllerEnkeltfnutt(attr.Notes)&";"&attr.Default&";"&gyldig&";"&date&";" ,0
'!	Repository.WriteOutput "Script",  kodenavn&";"&ikkelinjeskiftEllerEnkeltfnutt(attr.Notes)&";"&attr.Default&";;;;"&gyldig&";"&date&";" ,0
	Repository.WriteOutput "Script",  kodenavn&";"&ikkelinjeskiftEllerEnkeltfnutt(trimutf8(attr.Notes))&";"&utvekslingsalias&";;;;"&gyldig&";"&date&";" ,0

End Sub



Sub kopierKodensNavnTilTomDefinisjon(attr)
		if attr.Notes = "" then
			dim notestring
		  ' Move ALL (old) names to START of definition by commenting out the if/endif around
		  ' and use this "notestring =" instead
			' notestring = attr.Name & " " & attr.Notes
			notestring = attr.Name
			Repository.WriteOutput "Script", "New notestring: " & notestring,0
			attr.Notes = notestring
			attr.Update()
		end if

End Sub

Sub kopierKodensNavnTilTagSOSI_presentasjonsnavn(attr)
  		'Repository.WriteOutput "Script", "SOSI_presentasjonsnavn: " & attr.Name,0
      Call TVSetElementTaggedValue(attr, "SOSI_presentasjonsnavn", attr.Name)

End Sub


Sub flyttInitialverdiTilTagSOSI_verdi(attr)
		If attr.Default <> "" then
  		Repository.WriteOutput "Script", "Initial value moved: " & attr.Default,0

      Call TVSetElementTaggedValue(attr, "SOSI_verdi", attr.Default)

      attr.Default = ""
      attr.Update()

		End if
End Sub

sub TVSetElementTaggedValue( theElement, taggedValueName, taggedValue)
	'Repository.WriteOutput "Script", "  Checking if tagged value [" & taggedValueName & "] exists",0
	if not theElement is nothing and Len(taggedValueName) > 0 then
		dim newTaggedValue as EA.TaggedValue
		set newTaggedValue = nothing
		dim taggedValueExists
		taggedValueExists = False

		'check if the element has a tagged value with the provided name
		dim existingTaggedValue AS EA.TaggedValue
		dim currentExistingTaggedValue AS EA.TaggedValue
		dim taggedValuesCounter
		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set existingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			if existingTaggedValue.Name = taggedValueName then
				taggedValueExists = True
				set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			end if
		next

		'if the element does not contain a tagged value with the provided name, create a new one
		if not taggedValueExists = True then
			set newTaggedValue = theElement.TaggedValues.AddNew( taggedValueName, taggedValue )
			newTaggedValue.Update()
			'Repository.WriteOutput "Script", "    ADDED tagged value ["& taggedValueName & " " & taggedValue & "]",0
		Else
		  If currentExistingTaggedValue.Value = "" Then
		    currentExistingTaggedValue.Value = taggedValue
		    currentExistingTaggedValue.Update()
			  ' Repository.WriteOutput "Script", "    ADDED value ["& taggedValueName & " " & taggedValue& "]",0
		  End If
			'Repository.WriteOutput "Script", "    FOUND tagged value ["& taggedValueName & " " & currentExistingTaggedValue.Value & "]",0
		end if
	end if
end Sub

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


Sub settKodensNavnTilNCName(attr)
		' make name legal NCName
		' (alternatively replace each bad character with a "_", typically used for codelist with proper names.)
		' (Sub settBlankeIKodensNavnTil_(attr))
    Dim txt, txt1, txt2, res, tegn, i, u
    u=0
		'Repository.WriteOutput "Script", "Old code: " & attr.Name,0
		txt = Trim(attr.Name)
		res = ""
			'Repository.WriteOutput "Script", "New NCName: " & txt & " " & res,0

		' loop gjennom alle tegn
		For i = 1 To Len(txt)
		  ' blank, komma, !, ", #, $, %, &, ', (, ), *, +, /, :, ;, <, =, >, ?, @, [, \, ], ^, `, {, |, }, ~
		  ' (tatt med flere fnuttetyper, men hva med "."?)
		  tegn = Mid(txt,i,1)
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
				if res = "" then
					if tegn = "1" or tegn = "2" or tegn = "3" or tegn = "4" or tegn = "5" or tegn = "6" or tegn = "7" or tegn = "8" or tegn = "9" or tegn = "0" or tegn = "-" or tegn = "." Then
						' NCNames can not start with any of these characters, skip this
					else
						If u = 1 Then
							res = res + UCase(tegn)
							u=0
						else
							res = res + tegn
						end if
					end if
				else
					'Repository.WriteOutput "Script", "Good: " & tegn & "  " & i & " " & u,0
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
		txt1 = LCase( Mid(res,1,1) )
		i = Len(res) - 1
		if i < 0 then
			Repository.WriteOutput "Script", "Error: Unable to construct NCName for code: [" & attr.Name & "]",0
		else
			txt2 =Mid(res,2,i)
			txt = txt1 + txt2
			if txt <> attr.Name then
				Repository.WriteOutput "Script", "Change: Old code: [" & attr.Name & "] changed to new NCName: [" & txt & "]",0
				' return txt
				attr.Name = txt
				attr.Update()
			end if
		end if

End Sub

Sub settKodensNavnTilEgen_Navn(attr)
		' make name legal NCName by replacing each bad character with a "_", typically used for codelist with proper names.)

    Dim txt, res, tegn, i, u
    u=0
		txt = Trim(attr.Name)
		'res = LCase( Mid(txt,1,1) )
		res = Mid(txt,1,1)
			'Repository.WriteOutput "Script", "New NCName: " & txt & " " & res,0

		' loop gjennom alle tegn
		For i = 2 To Len(txt)
		  ' blank, komma, !, ", #, $, %, &, ', (, ), *, +, /, :, ;, <, =, >, ?, @, [, \, ], ^, `, {, |, }, ~
		  ' (tatt med flere fnuttetyper, men hva med "."?)
		  tegn = Mid(txt,i,1)
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
			        res = res + "_"
		          'res = res + UCase(tegn)
		          u=0
			      'else
		        End If
		        res = res + tegn
		      End If
		    End If
		  End If
		Next
		Repository.WriteOutput "Script", "New NCName: " & res,0
		' return res
		attr.Name = res
		attr.Update()

End Sub

Function ikkelinjeskiftEllerEnkeltfnutt(streng)
    Dim txt, res, tegn, i, u
        u = 0
		txt = ""
		' loop gjennom alle tegn og ta bort linjeskift og lignende lavverditegn
		For i = 1 To Len(streng)
            tegn = Mid(streng,i,1)
			if tegn >  " " then
			    if tegn =  "'" or tegn = "’" then
			       txt = txt + """"
			    else
					if tegn = ";" or tegn = "<" or tegn = ">" then
					   txt = txt + " "
					else
					   txt = txt + tegn
					end if
				end if
			    u = 0
			else
			    if u = 0 then
			      txt = txt + " "
			    end if
			    u = 1
			end if
        next
  ikkelinjeskiftEllerEnkeltfnutt = txt
End Function


'-----------------Function trimutf8 Start-----------------
function trimutf8(txt)
	'convert national characters back to utf8
    Dim inp
'	dim res, tegn, i, u, ÉéÄäÖöÜü-Áá &#233; forrige &#229;r i samme retning skal den h&#248; prim&#230;rt prim&#230;rt

	inp = Trim(txt)
	if InStr(1,inp,"&#230;",0) <> 0 then
		inp = Replace(inp,"&#230;","æ",1,-1,0)
	end if
	if InStr(1,inp,"&#248;",0) <> 0 then
		inp = Replace(inp,"&#248;","ø",1,-1,0)
	end if
	if InStr(1,inp,"&#229;",0) <> 0 then
		inp = Replace(inp,"&#229;","å",1,-1,0)
	end if
	if InStr(1,inp,"&#198;",0) <> 0 then
		inp = Replace(inp,"&#198;","Æ",1,-1,0)
	end if
	if InStr(1,inp,"&#216;",0) <> 0 then
		inp = Replace(inp,"&#216;","Ø",1,-1,0)
	end if
	if InStr(1,inp,"&#197;",0) <> 0 then
		inp = Replace(inp,"&#197;","Å",1,-1,0)
	end if
	if InStr(1,inp,"&#201;",0) <> 0 then
		inp = Replace(inp,"&#201;","É",1,-1,0)
	end if
	if InStr(1,inp,"&#233;",0) <> 0 then
		inp = Replace(inp,"&#233;","é",1,-1,0)
	end if
	' -Áá-ČčĐđŊŋŠšŦŧŽž-Ńń-Ïï
	' &#252;-&#193;&#225;-ČčĐđŊŋŠšŦŧŽž-Ńń-&#207;&#239;
	if InStr(1,inp,"&#196;",0) <> 0 then
		inp = Replace(inp,"&#196;","Ä",1,-1,0)
	end if
	if InStr(1,inp,"&#228;",0) <> 0 then
		inp = Replace(inp,"&#228;","ä",1,-1,0)
	end if
	if InStr(1,inp,"&#214;",0) <> 0 then
		inp = Replace(inp,"&#214;","Ö",1,-1,0)
	end if
	if InStr(1,inp,"&#246;",0) <> 0 then
		inp = Replace(inp,"&#246;","ö",1,-1,0)
	end if
	if InStr(1,inp,"&#220;",0) <> 0 then
		inp = Replace(inp,"&#220;","Ü",1,-1,0)
	end if
	if InStr(1,inp,"&#252;",0) <> 0 then
		inp = Replace(inp,"&#252;","ü",1,-1,0)
	end if
	if InStr(1,inp,"&#193;",0) <> 0 then
		inp = Replace(inp,"&#193;","Á",1,-1,0)
	end if
	if InStr(1,inp,"&#225;",0) <> 0 then
		inp = Replace(inp,"&#225;","á",1,-1,0)
	end if
	' ÆØÅæøåÉéÄäÖöÜü-Áá-ČčĐđŊŋŠšŦŧŽž-Ńń-Ïï
	' &#198;&#216;&#197;&#230;&#248;&#229;&#201;&#233;&#196;&#228;&#214;&#246;&#220;&#252;-&#193;&#225;-ČčĐđŊŋŠšŦŧŽž-Ńń-&#207;&#239;
	if InStr(1,inp,"&#207;",0) <> 0 then
		inp = Replace(inp,"&#207;","Ï",1,-1,0)
	end if
	if InStr(1,inp,"&#239;",0) <> 0 then
		inp = Replace(inp,"&#239;","ï",1,-1,0)
	end if

	trimutf8 = inp
end function
'-----------------Function trimutf8 End-----------------

function getBroader(codelist,attr)
	dim gruppe
	getBroader = ""
	if codelist.Name = "Navneobjekttype" then
		gruppe = getTaggedValue(attr,"SOSI_verdi")
		if gruppe <> "" then
			i1 = Int(CInt(gruppe) / 100) * 100
			if groupsList.IndexOfKey(CStr(i1)) <> -1 then
				getBroader = groupsList.getByIndex(groupsList.IndexOfKey(CStr(i1)))
			end if
		end if
	end if
	if codelist.Name = "Navneobjektgruppe" then
		gruppe = getTaggedValue(attr,"SOSI_verdi")
		if gruppe <> "" then
			i1 = Int(CInt(gruppe) / 1000) * 1000
			if maingroupsList.IndexOfKey(CStr(i1)) <> -1 then
				getBroader = maingroupsList.getByIndex(maingroupsList.IndexOfKey(CStr(i1)))
			end if
		end if
	end if
end function

sub fillGroups()
	groupsList.Add "1100", "terrengomrÃ¥der"
	groupsList.Add "1200", "hÃ¸yder"
	groupsList.Add "1300", "senkninger"
	groupsList.Add "1400", "flater"
	groupsList.Add "1500", "skrÃ¥ninger"
	groupsList.Add "1600", "terrengdetaljer"
	groupsList.Add "2100", "bartFjell"
	groupsList.Add "2200", "lÃ¸smasseavsetninger"
	groupsList.Add "2300", "vegetasjon"
	groupsList.Add "2400", "vÃ¥tmark"
	groupsList.Add "2500", "dyrkamark"
	groupsList.Add "2600", "isOgPermafrost"
	groupsList.Add "2700", "uttakOgDeponi"
	groupsList.Add "3100", "stillestÃ¥endeVann"
	groupsList.Add "3200", "ferskvannskontur"
	groupsList.Add "3300", "grunnerIFerskvann"
	groupsList.Add "3400", "rennendeVann"
	groupsList.Add "3500", "detaljerIFerskvann"
	groupsList.Add "4100", "farvann"
	groupsList.Add "4200", "kystkontur"
	groupsList.Add "4300", "grunnerISjÃ¸"
	groupsList.Add "4400", "sjÃ¸bunn"
	groupsList.Add "4500", "detaljISjÃ¸"
	groupsList.Add "5100", "bebyggelsesomrÃ¥der"
	groupsList.Add "5200", "gardsbebyggelse"
	groupsList.Add "5300", "bolighus"
	groupsList.Add "5400", "nÃ¦ring"
	groupsList.Add "5500", "institusjoner"
	groupsList.Add "5600", "fritidsanlegg"
	groupsList.Add "6100", "veg"
	groupsList.Add "6200", "bane"
	groupsList.Add "6300", "luftfart"
	groupsList.Add "6400", "sjÃ¸fart"
	groupsList.Add "6500", "navigasjon"
	groupsList.Add "6600", "samferdselsanlegg"
	groupsList.Add "6700", "energi"
	groupsList.Add "6800", "kommunikasjon"
	groupsList.Add "7100", "administrativeIndelinger"
	groupsList.Add "7200", "verne-OgBruksomrÃ¥der"
	groupsList.Add "8100", "kulturminner"
	groupsList.Add "8200", "kulturinstitusjoner"
	'			gruppe = maingroupsList.getByIndex(maingroupsList.IndexOfKey(i1))
'	Repository.WriteOutput "Script"," debug: groupsList.GetKey(1) - groupsList.getByIndex(1): "&groupsList.GetKey(1)&" "&groupsList.getByIndex(1), 0
'	Repository.WriteOutput "Script"," debug: groupsList.IndexOfKey(1100): "&groupsList.IndexOfKey(1100), 0

end sub


sub fillMaingroups()
	maingroupsList.Add "1000", "terreng"
	maingroupsList.Add "2000", "markslag"
	maingroupsList.Add "3000", "ferskvann"
	maingroupsList.Add "4000", "sjÃ¸"
	maingroupsList.Add "5000", "bebyggelse"
	maingroupsList.Add "6000", "infrastruktur"
	maingroupsList.Add "7000", "offentligAdministrasjon"
	maingroupsList.Add "8000", "kultur"
end sub


oppdaterKoderForEnValgtKodeliste
