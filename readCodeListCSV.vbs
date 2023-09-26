option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		readCodeListCSV
' purpose:		Start with a empty CodeList class, read a CSV-file and add the codes to the class
' formål:		Les inn koder fra en CSV-fil til en tom kodelisteklasse
' author:		Kent
' version:		2023-09-26	automatisk sette (Utgått) når gyldigTil er satt, ta vare på samisk i utf8 fra csv
'							datoformat til iso8601-format, legge in ID på hver kode
' version:		2019-05-23
'
'
'				TBD: utvekslingsalias til tagged value ?
'				TBD:

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

sub readCodeListCSV()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	DIM i
	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		'f Repository.GetTreeSelectedItemType() = otPackage then
		if Repository.GetTreeSelectedItemType() = otElement and UCASE(theElement.Stereotype) = "CODELIST" then
			'Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0
			dim message
			dim box
			box = Msgbox ("Skript readCodeListCSV" & vbCrLf & vbCrLf & "Skriptversjon 2023-09-26" & vbCrLf & "Leser koder fra en semikolonseparert tabellfil inn i kodelisteklassen : [" & theElement.Name & "].",1)
			select case box
			case vbOK
				dim kortnavn
				'tømmer System Output for lettere å fange opp eventuelle feilmeldinger der
				Repository.ClearOutput "Script"
				Repository.CreateOutputTab "Error"
				Repository.ClearOutput "Error"
				
				'get file name
				kortnavn = "C:\Kent\b\kommunenummer2024.csv"
				kortnavn = InputBox("Angi CSV-filas navn.", "filnavn", kortnavn)
				Dim fso
				Dim file
				Dim line
				Dim arr
				Dim txt
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set file = fso.OpenTextFile (kortnavn, 1, 0)
				'les forbi første linje
				line = file.Readline		
				Do Until file.AtEndOfStream
					line = file.Readline		
					'Repository.WriteOutput "Script", Now & "line [" & line & "]" & vbCrLf ,0
					arr = Split(line, ";")
					'Navn;Kodeverdi;Eier; Status;Oppdatert;Versjons ID;Beskrivelse;Gyldig fra;Gylig til; ID
					'Repository.WriteOutput "Script", "0-5[" & arr(0) & "]" &"[" & arr(1) & "]" &"[" & arr(2) & "]" &"[" & arr(3) & "]" &"[" & arr(4) & "]" &"[" & arr(5) & "]" & vbCrLf ,0
					'Repository.WriteOutput "Script", "6-9[" & arr(6) & "]" &"[" & arr(7) & "]" &"[" & arr(8) & "]" &"[" & arr(9) & "]" & vbCrLf ,0
					'add code as attribute
	'	if false then
					Dim newCode as EA.Attribute
					Dim newTag as EA.AttributeTag
					set newCode = theElement.attributes.AddNew(readutf8(arr(0)),"Attribute")
					theElement.attributes.Refresh()
					
					'add name as description, and (Utgått) if code starts with 01,02,06 etc.
					'newCode.Notes = arr(6)
					txt = readutf8(arr(6))
	'				newCode.Notes = txt
					newCode.Type = ""
					' ??
					newCode.Default = arr(1)
					newCode.Update()
	'	end if			
	'				txt = readutf8(arr(6))
					if arr(8) <> "" then
						txt = txt & ", kode " & arr(1) & ", (utgått " & iso8601date(arr(8)) & ")"
					else
						if arr(3) = "Tilbaketrukket" then
							txt = txt & ", kode " & arr(1) & ", (tilbaketrukket) "
						else
							txt = txt & ", kode " & arr(1)
						end if		
					end if
					Repository.WriteOutput "Script", "name[" & readutf8(arr(0)) & "]" & "  description[" & txt & "]" & "  [" & iso8601date(arr(7)) & "]" & "  [" & iso8601date(arr(8)) & "]" & vbCrLf ,0
					
	'	if false then
					newCode.Notes = txt
					'add tagged values oppdateringsdato, gyldigFra if set, and gyldigTil if set
					if arr(4) <> "" then
						Set newTag = newCode.TaggedValues.AddNew("oppdateringsdato","AttributeTag")
	'					newTag.Value = arr(4)
						newTag.Value = iso8601date(arr(4))
						newTag.Update()
						'newCode.Refresh()
					end if
					if arr(7) <> "" then
						Set newTag = newCode.TaggedValues.AddNew("gyldigFra","AttributeTag")
						newTag.Value = iso8601date(arr(7))
						newTag.Update()
						'newCode.Refresh()
					end if
					if arr(8) <> "" then
						Set newTag = newCode.TaggedValues.AddNew("gyldigTil","AttributeTag")
						newTag.Value = iso8601date(arr(8))
						newTag.Update()
						'newCode.Refresh()
					end if
					if arr(9) <> "" then
						Set newTag = newCode.TaggedValues.AddNew("IRI","AttributeTag")
	'					newTag.Value = arr(9)
						newTag.Value = "http://skjema.geonorge.no/SOSI/kodeliste/AdmEnheter/2024/Kommunenummer/" & arr(1)
						newTag.Update()
						'newCode.Refresh()
					end if
		'			if arr(1) <> "" then
		'				Set newTag = newCode.TaggedValues.AddNew("utvekslingsalias","AttributeTag")
		'				newTag.Value = arr(1)
		'				newTag.Update()
		'				'newCode.Refresh()
		'			end if
					'
					'
					newCode.Update()
					theElement.attributes.Refresh()
	'	end if			
				Loop				
				file.Close	
				

				Repository.WriteOutput "Script", Now & " Fil lest: " & kortnavn & ".",0

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


function iso8601date(str)
	iso8601date = ""
	if Len(str) >= 10 then
		iso8601date = Mid(str,7,4) & "-" & Mid(str,4,2) & "-" & Mid(str,1,2)
	end if
end function


function readutf8(str)
	' make string utf-8
	Dim txt, res, tegn, utegn, vtegn, wtegn, xtegn, i
	Dim c3, c4, c5, c6, c7, ca, e2
	
	readutf8 = ""
	'exit function
	
	'txt = Trim(str)
	txt = str
		c3 = 195 ' vanligste startbyte for utf8
			c4 = 196
			c5 = 197
			c6 = 198
			c7 = 199
			ca = 202
			e2 = 226
'		c3 = 195 b6 = 182 ö
'		c3 = 195 a1 = 161 á
'		c5 = 197 8b = 139 ŋ
		
	' loop gjennom alle tegn
	res = ""
	i = 0
	While i < Len(txt)
		i = i + 1
		tegn = Mid(txt,i,1)
'		Repository.WriteOutput "Script", "readutf8: i=" & i & " tegn=" & tegn,0
		
		Select case int(AscW(tegn))
			Case c3
				'Repository.WriteOutput "Script", "readutf8: c3A i=" & i & " tegn=" & tegn & "int(AscW(tegn))=" & int(AscW(tegn)) ,0
				'Repository.WriteOutput "Script", "readutf8: c3B int(AscW((Mid(txt,i+1,1)))=" & int(AscW(Mid(txt,i+1,1))) ,0
				
				i = i + 1
				Select case int(AscW(Mid(txt,i,1)))
					Case 134
						res=res+"Æ"
						
					Case 152
						res=res+"Ø"
					Case 732
						res=res+"Ø"

					Case 133
						res=res+"Å"
					Case 8230
						res=res+"Å"

					Case 166
						res=res+"æ"

					Case 184
						res=res+"ø"

					Case 165
						res=res+"å"
						
					Case 129
						res=res+"Á"
						
					Case 161
						res=res+"á"
						
					Case 182
						res=res+"ö"

					Case else
						utegn = int(AscW(tegn)) & 511
						vtegn = utegn * 64
						wtegn = int(AscW(Mid(txt,i+1,1))) & 1023
						xtegn = wtegn
						Repository.WriteOutput "Script", "readutf8: c3 i=" & i & " tegn=" & tegn & " " & int(AscW(Mid(txt,i,1))) & " -> " & utegn & " + " & vtegn & " + " & wtegn & " + " & xtegn & " + " & res & " " & AscW(vtegn)+AscW(xtegn),0
		'				res=res+Chr(AscW(vtegn)+AscW(xtegn))
						res=res+"Ã"+Mid(txt,i,1)
				End Select
			'Case c4
			
			Case c5
				'Repository.WriteOutput "Script", "readutf8: c5A i=" & i & " tegn=" & tegn & "int(AscW(tegn))=" & int(AscW(tegn)) ,0
				'Repository.WriteOutput "Script", "readutf8: c5B int(AscW((Mid(txt,i+1,1)))=" & int(AscW(Mid(txt,i+1,1))) ,0
			
				i = i + 1
				Select case int(AscW(Mid(txt,i,1)))

					Case 160
						res=res+"Š"
						
					Case 161
						res=res+"š"
						
					Case 138
						res=res+"Ŋ"
						
					Case 139
						res=res+"ŋ"
						
		'			Case 8249
		'				res=res+"ŋ"
						
					Case else
						utegn = int(AscW(tegn)) & 511
						vtegn = utegn * 64
						wtegn = int(AscW(Mid(txt,i+1,1))) & 1023
						xtegn = wtegn
						Repository.WriteOutput "Script", "readutf8: c5 i=" & i & " tegn=" & tegn & " " & Mid(txt,i,1) & " " & int(AscW(Mid(txt,i,1))) & " -> " & utegn & " + " & vtegn & " + " & wtegn & " + " & xtegn & " + " & res & " " & AscW(vtegn)+AscW(xtegn),0
		'				res=res+Chr(AscW(vtegn)+AscW(xtegn))
						res=res+"Ã"+Mid(txt,i,1)
						
					End Select
				
			'Case c6
			
			'Case c7
			
			'Case ca

			Case e2
				' + 80 + 93 = "en dash" = halv fonthøyde, ( + 80 + 94 = "em dash" = full fonthøyde)
				i = i + 2
				res=res+"–"
			Case else
				res=res+tegn
		
		End Select

	Wend
	readutf8 = res



End function


readCodeListCSV
