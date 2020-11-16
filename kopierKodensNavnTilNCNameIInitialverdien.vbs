option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		kopierKodensNavnTilNCNameIInitialverdien
' purpose:		S
' formål:		kopierer koders navn til NCNames i initialverdien dersom initialverdien er tom
' author:		Kent
' version:		2019-05-24 (flyttKodensNavnTilInitialverdi)
' version:		2020-11-16 lag NCnavn til initialverdiene
'
'
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

sub flyttKodensNavnTilInitialverdi()
	' Show and clear the script output window
	Repository.EnsureOutputVisible "Script"
	DIM i
	Dim theElement as EA.Element
	Set theElement = Repository.GetTreeSelectedObject()
	if not theElement is nothing  then
		'if theElement.Type="Package" and UCASE(theElement.Stereotype) = "APPLICATIONSCHEMA" then
		'f Repository.GetTreeSelectedItemType() = otPackage then
		if Repository.GetTreeSelectedItemType() = otElement and UCASE(theElement.Stereotype) = "CODELIST" then
	'	if Repository.GetTreeSelectedItemType() = otElement and UCASE(theElement.Stereotype) = "CODELIST" or ""otEnumeration"" then
			if debug then Repository.WriteOutput "Script", Now & " " & theElement.Stereotype & " " & theElement.Name, 0 end if
			dim message
			dim box
			box = Msgbox ("Skript kopierKodensNavnTilNCNameIInitialverdien" & vbCrLf & vbCrLf & "Skriptversjon 2020-11-16" & vbCrLf & "kopierer koders navn til NCNames i tomme initialverdier : [" & theElement.Name & "].",1)
			select case box
			case vbOK
				'tømmer System Output for lettere å fange opp eventuelle feilmeldinger der
				Repository.ClearOutput "Script"
				Repository.CreateOutputTab "Error"
				'Repository.ClearOutput "Error"
				
				dim attr as EA.Attribute
				for each attr in theElement.Attributes
				
					if attr.Default = "" then
						attr.Default = getNCName(attr.Name)
						'attr.Notes = getCleanDefinitionText(attr)
						attr.Update()
						Repository.WriteOutput "Script", "kopiert navn til NCName i initialverdi : ["  & attr.Name & "] [" & attr.Default & "]" & "[" & attr.Notes & "]" & vbCrLf ,0
					else
						Repository.WriteOutput "Script", "ingen kopiering for kode med navn : ["  & attr.Name & "] [" & attr.Default & "]" & "[" & attr.Notes & "]" & vbCrLf ,0
					end if
				
				next

				Repository.WriteOutput "Script", Now & " Endret koder for klasse: " & theElement.Name & ".",0

			case VBcancel

			end select
	

		Else
		  'Other than CodeList selected in the tree
		  MsgBox( "This script requires a codelist class to be selected in the Project Browser." & vbCrLf & _
			"Please select a codelist class in the Project Browser and try again." )
		end If
		'Repository.WriteOutput "Script", Now & " Finished, check the Error and Types tabs", 0
	else
		'No CodeList selected in the tree
		MsgBox( "This script requires a codelist class to be selected in the Project Browser." & vbCrLf & _
	  "Please select a codelist class in the Project Browser and try again." )
	end if
end sub

function toLabel(name)
	'expands tecnical NCNames to normal language names
    Dim txt, res, tegn, i, u
    u=0
	toLabel = ""
	txt = Trim(name)
		res = Mid(txt,1,1)
		' loop gjennom alle resterende tegn og sett inn blank og liten bokstav der det er stor bokstav
		For i = 2 To Len(txt)
			tegn = Mid(txt,i,1)
			If tegn <> LCase(tegn) Then
				res = res + " "
				res = res + LCase(tegn)
			Else 
				res = res + tegn
			End If
		Next
		
	toLabel = res

end function

function getNCName(str)
	' make name legal NCName
	Dim txt, res, tegn, i, u
    u=0
		txt = Trim(str)
		'res = LCase( Mid(txt,1,1) )
		'if Mid(txt,1,1) < ":" then
		'	res = "_" + Mid(txt,1,1)
		'else
			res = Mid(txt,1,1)
		'end if
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
		          res = res + UCase(tegn)
		          u=0
			      else
		          res = res + tegn
		        End If
		      End If
		    End If
		  End If
		Next
		' return res
		'getNCName = res
		getNCName = LCase(Mid(res,1,1)) + Mid(res,2,Len(res))

End function


function getCleanDefinitionText(currentElement)
	'removes all formatting in notes fields
    Dim txt, res, tegn, i, u
    u=0
	getCleanDefinitionText = ""
		'txt = Trim(currentElement.Notes)
		txt = currentElement.Notes
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
				  If tegn < " " Then
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
		
	getCleanDefinitionText = res

end function




flyttKodensNavnTilInitialverdi
