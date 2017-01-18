option explicit
!INC Local Scripts.EAConstants-VBScript
' Sub Name: listDiagrammerSomViserElementet (krav18-viseAlt)
' Author: Kent Jonsrud
' Date: 2016-09-29, 2017-01-05/09/17///


' -----------------------------------------------------------START---------------------------------------------------------------------------------------------------
'Global objects for testing whether a class is showing all its content in at least one diagram.  /krav/18
dim startPackage as EA.Package
dim diaoList
dim diagList
dim debug
' -----------------------------------------------------------END---------------------------------------------------------------------------------------------------
sub OnProjectBrowserScript()
	Repository.EnsureOutputVisible("Script")
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	select case treeSelectedType
		'
		case otElement
			Dim theElement as EA.Element
			Set theElement = Repository.GetTreeSelectedObject()
		'case otPackage
' -----------------------------------------------------------START---------------------------------------------------------------------------------------------------

			dim thePackage as EA.Package
			'set thePackage = getApplicationSchemaPackage(theElement)
			set thePackage = Repository.GetPackageByID(theElement.PackageID)
			'liste over koblinger mellom alle diaobjktID og elementID i applikasjonsskjemapakka der objektet ligger
			'if package is not not As then look for AS in all owner packages, start from found AS if found, otherwise start from package where element is found.
			'TBD find correct thePackage
			set startPackage = thePackage

			Set diaoList = CreateObject( "System.Collections.Sortedlist" )
			Set diagList = CreateObject( "System.Collections.Sortedlist" )
			debug = false
			'For testing whether a class is showing all its content in at least one diagram.  /krav/18
			recListDiagramObjects(thePackage)

' -----------------------------------------------------------END---------------------------------------------------------------------------------------------------
			'Debug
			Dim i
			For i = 0 To diaoList.Count - 1
				if debug then Session.Output("Debug: diaoList key [" & diaoList.GetKey(i) & "] value [" & diaoList.GetByIndex(i) & "]")
 			Next
			dim message
			dim box
			box = Msgbox ("Start searching for diagrams where element : [" & theElement.Name & "] is shown. The top owner package is: [" &thePackage.Name& "]. ",1)
			'box = Msgbox ("The owner package is: [" &thePackage.Name& "]. Start searching for diagrams where elements are shown.",1)
			select case box
			case vbOK
				'findElements(thePackage)
		 		if debug then Session.Output("Debug: ------------ Start class: [«" &theElement.Stereotype& "» " &theElement.Name& "] of type. [" &theElement.Type& "]. ")
				if theElement.Type = "Class" or theElement.Type = "DataType" or theElement.Type = "Enumeration" or theElement.Type = "Interface" then
					call krav18viseAlt(theElement)
				end if
			case VBcancel

			end select
		case else
			Session.Prompt "This script does not support items of this type.", promptOK
	end select
end sub


' -----------------------------------------------------------START---------------------------------------------------------------------------------------------------
' Sub Name: krav18-viseAlt
' Author: Kent Jonsrud
' Date: 2016-08-09..30, 2016-09-05, 2016-09-29
' Purpose: test whether a class is showing all its content in at least one class diagram.
    '/krav/18
    'Alle klasser og assosiasjoner skal i minst ett diagram vise alle arvede og alle egne egenskaper, roller, operasjoner og restriksjoner.
    'Sjekk at klasser i en pakke, deres egenskaper, operasjoner og restriksjoner samt assosiasjoner
    'der source er en klasse i pakka og assosiasjonsroller til disse assosiasjonene  vises i minst ett diagram (i samme pakke).
    'Det samme gjelder alle disse elementenes supertyper.
	'
	'
	'Sjekker om klasser vises i det hele tatt og om diagram-settings er ok.
	'Må sjekke assosiasjoner og arvede properties også vises i samme diagram.
	'Må prøve å se på om noen av klassene eller assosiasjonene ligger oppå hverandre eller er skjult av noe annet.
	'Intrikat: viser alle arvede egenskaper og operasjoner i klassen, og arver ingen assosiasjoner. Må da superklassen vises? (inherited constraints?)
	'
	'
    'kan vi lage test som finner datatyper eller kodelister som ikke er brukt som type til egenskaper?
    'krav 18 pri-1 som går på at alle elementer skal vise alt i minst ett diagram, og det bør hete Hoveddiagram (Ttttt).
    '
    '/krav/19 - seinere + visuell sjekk
    'Alle klasser og assosiasjoner skal ha en definisjon som beskriver mening og forståelse.
    'Aksepterer at assosiasjoner defineres tilstrekkelig gjennom sine assosiasjonsroller. (=må ha minst en navnet rolle med definisjon!)


sub krav18viseAlt(theElement)

	dim diagram1 as EA.Diagram
	dim diagram as EA.Diagram
	dim diagrams as EA.Collection
	dim diao1 as EA.DiagramObject
	dim diao as EA.DiagramObject
	dim conn as EA.Collection
	dim super as EA.Element
	dim base as EA.Collection
	dim child as EA.Collection
	dim embed as EA.Collection
	dim realiz as EA.Collection
	dim viserAlt
	viserAlt = false


	'Pseudocode:
	'-----------
	'for hver pakke (ikke med i /krav/18 men er mulig å tolkes i /krav/!)
	'	finn om pakken er vist i minst et diagram
	'	finn om de har lagret en avhengighetslink og om denne vises (enkel tolking)
	'	finn om de BURDE ha flere avhengighetslinker og om disse vises
	'		for alle klasser i pakka
	'			er det klasser som linker til eller bruker datatype fra ei ekstern pakke som det ikke er pakkeavhengighetslink til?
	'
	'
	'for hver klasse (og assosiasjon?  alle disse ivaretas vel via sin link til en klasse?)  (hva med noter?)
	' finn neste diagram den vises i (DiagramID)
	'  vises alle properties og constraints i diagrammet?
	'  		1-viser den alle egenskaper med type og mult, og operasjoner og constraints der?
	' 		2-viser den alle assosiasjonslinker med rollenavn, mult. og navigerbarhet og navnet på linket klasse der?
	'   	finn arved klasser
	'    	 viser disse arvede klassene alle sine properties og constraints i samme diagram?
	'			1-
	'			2-
	'		og er det varianter som oppnår det samme? (ShowInheritedAttributes vises på klassen og tomme supertyper vises med sine assosiasjoner?)
	'	ja -> OK/exit
	' ferdig med klassen og ingen funnet -> ERROR
	'neste klasse


	'Innledende debug      Repository.GetPackageByID(theElement.PackageID).Name
	if debug then
		Session.Output("Debug: theElement name [«" &theElement.Stereotype& "» " &theElement.Name& "]. ElementID "&theElement.ElementID)
		Session.Output("Debug: theElement ClassifierID [" &theElement.ClassifierID& "] Modified [" &theElement.Modified& "].  ")
		Session.Output("Debug: theElement ClassifierName [" &theElement.ClassifierName& "] ClassifierType [" &theElement.ClassifierType& "].  ")
		Session.Output("Debug: theElement ParentID [" &theElement.ParentID& "] StereotypeEx [" &theElement.StereotypeEx& "].  ")

		for each base in theElement.BaseClasses
			Session.Output("Debug: theElement BaseClass  [«" &base.Stereotype& "» " &base.Name& "] ClassifierType [" &base.ClassifierType& "].  ")
		next
		for each child in theElement.Elements
			Session.Output("Debug: theElement Child Element  [«" &child.Stereotype& "» " &child.Name& "] ClassifierType [" &child.ClassifierType& "].  ")
		next
		for each embed in theElement.EmbeddedElements
			Session.Output("Debug: theElement Embedded Element  [«" &embed.Stereotype& "» " &embed.Name& "] ClassifierType [" &embed.ClassifierType& "].  ")
		next
		for each realiz in theElement.Realizes
			Session.Output("Debug: theElement Realizes Element  [«" &realiz.Stereotype& "» " &realiz.Name& "] ClassifierType [" &realiz.ClassifierType& "].  ")
		next
	end if

	'navigate through all diagrams and find those the element knows
	Dim i, shownTimes
	shownTimes=0
	For i = 0 To diaoList.Count - 1
		'if debug then Session.Output("Debug: looking for element [" & theElement.Name & "] in diaoList element number [" & i & "] value [" & diaoList.GetKey(i) & "]")
		if theElement.ElementID = diaoList.GetByIndex(i) then
			if debug then Session.Output("Debug: element [" & theElement.Name & "] has value in diaoList element number [" & i & "] value [" & diaoList.GetKey(i) & "]")
			
			set diagram = Repository.GetDiagramByID(diagList.GetByIndex(i))
			if debug then Session.Output("Debug: diagram name [" & diagram.Name & "] of type ["&diagram.Type&"] ")

			shownTimes = shownTimes + 1
			for each diao in diagram.DiagramObjects
				if diao.ElementID = theElement.ElementID then
					exit for
				end if
			next

			if debug then
				Session.Output("Debug: ------------Diagram name ["&diagram.Name&"] of type ["&diagram.Type&"] showing class: [«" &theElement.Stereotype& "» " &theElement.Name& "].  ")
				Session.Output("Debug: DiagramID ["&diagram.DiagramID&"] with ShowPackageContents ["&diagram.ShowPackageContents&"] and ShowPublic (features) ["&diagram.ShowPublic&"] MetaType: [" &diagram.MetaType& "].  ")
				Session.Output("Debug: ExtendedStyle ["&diagram.ExtendedStyle&"].  ")
				Session.Output("Debug: StyleEX ["&diagram.StyleEX&"].  ")
				Session.Output("Debug: FilterElements ["&diagram.DiagramID&"]  MetaType: [" &diagram.MetaType& "].  ")
				Session.Output("Debug: diao DiagramID ["&diao.DiagramID&"] with ElementID ["&diao.ElementID&"] FeatureStereotypesToHide: " &diao.FeatureStereotypesToHide& " ShowNotes: " &diao.ShowNotes& ".  ")
				Session.Output("Debug: diao Public attributes compartment switch ["&diao.ShowPublicAttributes&"] with ElementID ["&diao.ElementID&"] FeatureStereotypesToHide: " &diao.FeatureStereotypesToHide& " theElement.Attributes.Count: " &theElement.Attributes.Count& ".  ")
			end if

			if theElement.Attributes.Count = 0 or InStr(1,diagram.ExtendedStyle,"HideAtts=1") = 0 then
				if theElement.Methods.Count = 0 or InStr(1,diagram.ExtendedStyle,"HideOps=1") = 0 then
					if InStr(1,diagram.ExtendedStyle,"HideEStereo=1") = 0 then
						if InStr(1,diagram.ExtendedStyle,"UseAlias=1") = 0 or theElement.Alias = "" then
							if debug then Session.Output("Debug: InStr(1,diagram.ExtendedStyle,'HideAtts=1') = " & InStr(1,diagram.ExtendedStyle,"HideAtts=1") )
							if debug then Session.Output("Debug: calls showAllProperties(theElement, diagram, diao)")
							if (showAllProperties(theElement, diagram, diao)) then
								'shows all OK in this diagram, how about inherited?
								if debug then Session.Output("Debug: showAllProperties OK in diagram ["&diagram.Name&"] for Element ["&theElement.Name&"].  ")
								viserAlt = true
							else
								if debug then Session.Output("Debug: showAllProperties FAIL in diagram ["&diagram.Name&"] for Element ["&theElement.Name&"].  ")
							end if
							if debug then Session.Output("Debug: Diagram ["&diagram.Name&"] shows Alias name ["&theElement.Alias&"] on Element ["&theElement.Name&"].")
							if debug then Session.Output("Debug: Diagram ["&diagram.Name&"] hides Stereotype ["&theElement.Stereotype&"] on Element ["&theElement.Name&"].")
							if debug then Session.Output("Debug: Diagram ExtendedStyle ["&diagram.ExtendedStyle&"].  ")
							if debug then Session.Output("Debug: Diagram StyleEX ["&diagram.StyleEX&"].  ")
						else
							if debug then Session.Output("Debug: +++++++++Diagram ExtendedStyle ["&diagram.ExtendedStyle&"].  ")
							Session.Output("Info: Diagram ["&diagram.Name&"] uses Alias name:  1,diagram.ExtendedStyle,'UseAlias=1') <> 0 and theElement.Alias <> ''")
						end if
					else
						if debug then Session.Output("Debug: +++++++++Diagram ExtendedStyle ["&diagram.ExtendedStyle&"].  ")
						Session.Output("Info: Diagram ["&diagram.Name&"] has turned off stereotype visibility.")					
						Session.Output("Info: Diagram ["&diagram.Name&"] has ....:  InStr(1,diagram.ExtendedStyle,'HideEStereo=1') <> 0")
						Session.Output("Info: Diagram ExtendedStyle ["&diagram.ExtendedStyle&"].  ")
					end if
				else
					if debug then Session.Output("Debug: +++++++++Diagram ExtendedStyle ["&diagram.ExtendedStyle&"].  ")
					Session.Output("Info: Diagram ["&diagram.Name&"] has turned off operation visibility.")					
					Session.Output("Info: Diagram ["&diagram.Name&"] has ....:  theElement.Methods.Count <> 0 or InStr(1,diagram.ExtendedStyle,HideOps=1) <> 0")				
				end if
			else
				if debug then Session.Output("Debug: +++++++++Diagram ExtendedStyle ["&diagram.ExtendedStyle&"].  ")
				Session.Output("Info: Diagram ["&diagram.Name&"] has turned off attribute visibility.")				
				Session.Output("Info: Diagram ["&diagram.Name&"] has ....:  theElement.Attributes.Count <> 0 and diagram.ExtendedStyle,HideAtts=1")				
			end if
		
		end if

	next

'TestEnd:	if debug then Session.Output("Debug: viserAlt: ["& viserAlt)
	if NOT viserAlt then
		'if debug then Session.Output("Error: Class not fully shown in at least one diagram: [«" &theElement.Stereotype& "» "&theElement.Name&"]   [/krav/18 ]")
		'Session.Output("Error: Class [«" &theElement.Stereotype& "» "&theElement.Name&"] not fully shown in at least one diagram.    [/krav/18 ]")
		if shownTimes = 0 then
			Session.Output("Error: Class [«" &theElement.Stereotype& "» "&theElement.Name&"] not shown in any diagram.    [/krav/18 ]")
		else
			Session.Output("Error: Class [«" &theElement.Stereotype& "» "&theElement.Name&"] not shown fully in at least one diagram.    [/krav/18 ]")
		end if
	else
		Session.Output("Ok: Element [«" &theElement.Stereotype& "» "&theElement.Name&"] is shown fully in at least one diagram.    [/krav/18 ]")

	end if



end sub

function showAllProperties(theElement, diagram, diagramObject)

	showAllProperties = false

	if debug then Session.Output("Debug: diagramObject DiagramID ["&diagramObject.DiagramID&"] with ElementID ["&diagramObject.ElementID&"] FeatureStereotypesToHide: " &diagramObject.FeatureStereotypesToHide& " ShowNotes: " &diagramObject.ShowNotes& ".  ")
	if debug then Session.Output("Debug: diagramObject Public attributes compartment switch ["&diagramObject.ShowPublicAttributes&"] with ElementID ["&diagramObject.ElementID&"] FeatureStereotypesToHide: " &diagramObject.FeatureStereotypesToHide& " theElement.Attributes.Count: " &theElement.Attributes.Count& ".  ")

	if debug then Session.Output("Debug: diagramObject Style ["&diagramObject.Style&"].  ")
	'Session.Output("Debug: diagramObject Style ["&diagramObject.Style&"].  ")
	'Session.Output("Debug: diagram StyleEX ["&diagram.StyleEX&"].  ")


'																										StyleEX
'																										SPL=S_BAB616=45F145

	'diagram.ExtendedStyle har en streng med diagrammsettings, diagramObject.Style har settings fra Featrue Visibility
	'if InStr(1,diagram.ExtendedStyle,"HideAtts=1") = 0 and diagramObject.ShowPublicAttributes or theElement.Attributes.Count = 0 then
	'if InStr(1,diagram.ExtendedStyle,"HideAtts=1") = 0 and diagramObject.ShowPublicAttributes and InStr(1,diagram.StyleEX,"SPL=") = 0 or theElement.Attributes.Count = 0 then
	if InStr(1,diagram.ExtendedStyle,"HideAtts=1") = 0 and diagramObject.ShowPublicAttributes and InStr(1,diagramObject.Style,"AttCustom=1" ) = 0 or theElement.Attributes.Count = 0 then
		if InStr(1,diagram.ExtendedStyle,"HideOps=1") = 0 and diagramObject.ShowPublicOperations or InStr(1,diagramObject.Style,"OpCustom=0" ) <> 0 or theElement.Methods.Count = 0 then
			if InStr(1,diagram.ExtendedStyle,"ShowCons=0") = 0 or diagramObject.ShowConstraints or InStr(1,diagramObject.Style,"Constraint=1" ) <> 0 or theElement.Constraints.Count = 0 then
				' all attribute parts really shown? ...
				if InStr(1,diagram.StyleEX,"VisibleAttributeDetail=1" ) = 0 or theElement.Attributes.Count = 0 then
					' if show all connections then
						showAllProperties = true
						Session.Output("Info: Diagram ["&diagram.Name&"] OK, shows all attributes and operations in class ["&theElement.Name&"]")				
					' else
						'Session.Output("Info 5 Diagram ["&diagram.Name&"] Roles.....=0 and diagramObject.ShowConstraints=false or InStr(1,diagramObject.Style,'Constraint=1' ) <> 0 or theElement.Constraints.Count > 0.  ")
						'showAllProperties = false
					' end if
				else
					Session.Output("Info: Diagram ["&diagram.Name&"] Fail to show all as Feaure ... Visibility is set to not show ???.")
					Session.Output("Info 4 diagram.StyleEX VisibleAttributeDetail=1.  ")
				end if
			else
				Session.Output("Info: Diagram ["&diagram.Name&"] Fail to show all as Diagram Properties are set to not show Constraints.")
				Session.Output("Info 3 ShowCons=0 and diagramObject.ShowConstraints=false or InStr(1,diagramObject.Style,'Constraint=1' ) <> 0 or theElement.Constraints.Count > 0.  ")
			end if
		else
			Session.Output("Info: Diagram ["&diagram.Name&"] Fail to show all as Diagram Properties are set to not show Operations.")
			Session.Output("Info 2 HideOps=1 and diagramObject.ShowPublicOperations=false or InStr(1,diagramObject.Style,'OpCustom=0' ) <> 0 or theElement.Methods.Count > 0.  ")
		end if
	else
		Session.Output("Info: Diagram ["&diagram.Name&"] Fail to show all as Diagram Properties are set to not show Attributes.")
		Session.Output("Info 1 HideAtts=1 and diagramObject.ShowPublicAttributes=false or InStr(1,diagramObject.Style,'AttCustom=0' ) <> 0  theElement.Attributes.Count > 0.  ")
	end if
end function




'Recursive loop through package p and its subpackages, creating a list of all model element showings (diagram objects) and their corresponding element
sub recListDiagramObjects(p)
	dim d as EA.Diagram
	dim Dobj as EA.DiagramObject
	if debug then Session.Output("Debug: Building list of diagram objects in package:  [" &p.Name& "]  PackageID: [" &p.PackageID&"]  ")
	for each d In p.diagrams
		for each Dobj in d.DiagramObjects
			if debug then Session.Output("Debug: Dobj in d.DiagramObjects, InstanceId:  [" &Dobj.InstanceID& "]  ElementId [" &Dobj.ElementID&"] diagram: [" &d.Name&"] element: [" &Repository.GetElementByID(Dobj.ElementID).Name& "].  ")
			If not diaoList.ContainsKey(Dobj.InstanceID) Then
				if debug then Session.Output("Debug: add to diaoList:  [" &Dobj.InstanceID& "] [" &Dobj.ElementID&"] diagram: [" &d.Name&"] element: [" &Repository.GetElementByID(Dobj.ElementID).Name& "].  ")
				diaoList.Add Dobj.InstanceID, Dobj.ElementID
				if debug then Session.Output("Debug: add to diagList:  [" &Dobj.InstanceID& "] [" &Dobj.DiagramID&"] diagram: [" &d.Name&"] element: [" &Repository.GetElementByID(Dobj.ElementID).Name& "].  ")
				diagList.Add Dobj.InstanceID, Dobj.DiagramID
			end if
		next
	next

	dim subP as EA.Package
	for each subP in p.packages
	    recListDiagramObjects(subP)
	next
end sub
' -----------------------------------------------END----------------------------------------------------------------------------------------------------------------


'
sub findElements(package)


	dim elements as EA.Element
	set elements = package.Elements

	dim packages as EA.Collection
	set packages = package.Packages

	dim diagrams as EA.Collection
	Dim diagram AS EA.Diagram
	Set diagrams = package.Diagrams
			'debug
			for each diagram in diagrams
				if debug then Session.Output("Debug: theElement Diagrams  [" &diagram.DiagramID& " " &diagram.Name& "].  ")
			next

			'Session.Output( " -Testing package: " & package.Name)
			' Navigate the elements collection, pick the classes
			'Session.Output( " number of elements in package: " & elements.Count)
			'debug
			dim i,j
			j = 0
			for i = 0 to elements.Count - 1
				dim currentElement as EA.Element
				set currentElement = elements.GetAt( i )

' -----------------------------------------------START--------------------------------------------------------------------------------------------------------------
		 		if debug then Session.Output("Debug: ------------ Start class: [«" &currentElement.Stereotype& "» " &currentElement.Name& "] of type. [" &currentElement.Type& "]. ")
				'if currentElement.Type = "Class" or currentElement.Type = "DataType" or currentElement.Type = "Enumeration" or currentElement.Type = "Interface" then
					'Iso 19103 Requirement 18 - each classifier must show all its (inherited) properties together in at least one diagram.
					call krav18viseAlt(currentElement)
					j = j + 1
				'end if
' -----------------------------------------------END----------------------------------------------------------------------------------------------------------------


			next
	dim p
' -----------------------------------------------START--------------------------------------------------------------------------------------------------------------
			if debug then Session.Output("Debug: End of testing theElement Diagrams  [" &diagram.DiagramID& " " &diagram.Name& "].  ")
			if false then
			dim currentPackage as EA.Package
			for p = 0 to packages.Count - 1
				set currentPackage = packages.GetAt( p ) 'getAT
				findElements(currentPackage) 'searching for other packages in the package
				'Set diagrams = currentPackage.Diagrams
				'for each diagram in diagrams
				'	if debug then Session.Output("Debug: Diagrams in subpackages [" &diagram.DiagramID& " " &diagram.Name& "].  ")
				'next
			next
			end if
' -----------------------------------------------END----------------------------------------------------------------------------------------------------------------
			'Session.Output( " -Number of elements tested: " & j & "/" & elements.Count & " in " & package.Name)

end sub

OnProjectBrowserScript
