option explicit
!INC Local Scripts.EAConstants-VBScript
' Sub Name: listDiagrammerSomViserElementet (krav18-viseAlt) 
' Author: Kent Jonsrud
' Date: 2016-09-29, 2017-01-05/09/17/02-21/22/05-12/13


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
			if debug then
			Dim i
			For i = 0 To diaoList.Count - 1
				'if debug then Session.Output("Debug: diaoList key [" & diaoList.GetKey(i) & "] value [" & diaoList.GetByIndex(i) & "]")
				Session.Output("Debug: Diagram: [" & Repository.GetDiagramByID(diagList.GetByIndex(i)).Name & "] Class: [" & Repository.GetElementByID(diaoList.GetByIndex(i)).Name & "]")
 			Next
			end if
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
	dim dial as EA.DiagramLink
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

			'test that the EA GUI has set all visibilityswithes correctly for showing classes
			if theElement.Attributes.Count = 0 or InStr(1,diagram.ExtendedStyle,"HideAtts=1") = 0 then
				if theElement.Methods.Count = 0 or InStr(1,diagram.ExtendedStyle,"HideOps=1") = 0 then
					if InStr(1,diagram.ExtendedStyle,"HideEStereo=1") = 0 then
						if InStr(1,diagram.ExtendedStyle,"UseAlias=1") = 0 or theElement.Alias = "" then
							if debug then Session.Output("Debug: InStr(1,diagram.ExtendedStyle,'HideAtts=1') = " & InStr(1,diagram.ExtendedStyle,"HideAtts=1") )
							if debug then Session.Output("Debug: calls PropertiesShown(theElement, diagram, diao)")
							if (PropertiesShown(theElement, diagram, diao)) then
								'shows all OK in this diagram, how about inherited?
								if debug then Session.Output("Debug: PropertiesShown OK in diagram ["&diagram.Name&"] for Element ["&theElement.Name&"].  ")
								viserAlt = true
							else
								if debug then Session.Output("Debug: PropertiesShown FAIL in diagram ["&diagram.Name&"] for Element ["&theElement.Name&"].  ")
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

function PropertiesShown(theElement, diagram, diagramObject)

	dim conn as EA.Connector
	dim super as EA.Element
	dim diaos as EA.DiagramObject
	dim SuperpropertiesShown, InheritanceHandled, supername
	PropertiesShown = false
	SuperpropertiesShown = true
	InheritanceHandled = true
	supername = ""

	if debug then Session.Output("Debug: PropertiesShown - theElement ElementID ["&theElement.ElementID&"] with Name ["&theElement.Name&"] FeatureStereotypesToHide: " &diagramObject.FeatureStereotypesToHide& " ShowNotes: " &diagramObject.ShowNotes& ".  ")
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
		'Diagram Properties are set to show Attributes, or no Attributes in the class
		if InStr(1,diagram.ExtendedStyle,"HideOps=1") = 0 and diagramObject.ShowPublicOperations or InStr(1,diagramObject.Style,"OpCustom=0" ) <> 0 or theElement.Methods.Count = 0 then
			'Diagram Properties are set to show Operations, or no Operations in the class
			if InStr(1,diagram.ExtendedStyle,"ShowCons=0") = 0 or diagramObject.ShowConstraints or InStr(1,diagramObject.Style,"Constraint=1" ) <> 0 or theElement.Constraints.Count = 0 then
				'Diagram Properties are set to show Constraints, or no Constraints in the class
				' all attribute parts really shown? ...
				if InStr(1,diagram.StyleEX,"VisibleAttributeDetail=1" ) = 0 or theElement.Attributes.Count = 0 then
					'Feaure Visibility is set to show all Attributes
					if InStr(1,diagram.ExtendedStyle,"HideRel=0") = 1 or theElement.Connectors.Count = 0 then
						'Diagram Properties set up to show all Associations, or no Associations in the class				
						if AssociationsShown(theElement, diagram, diagramObject) then
							if debug then Session.Output("Debug: All Associations shown ok in this diagram: ["&diagram.Name&"].  ")
							'All Associations shown ok in this diagram
							'Must now recurce and check that all inherited elements are also shown in this diagram
							'Any Supertype exist?
								for each conn in theElement.Connectors
									if debug then Session.Output("Debug: Connector 0 ConnectorID: [" & conn.ConnectorID & "]  Type: ["&conn.Type&"] ")
									if debug then Session.Output("Debug: Connector 1 SupplierID: [" & conn.SupplierID & "]  ClientID: ["&conn.ClientID&"] ")
									if conn.Type = "Generalization" then
										if theElement.ElementID = conn.ClientID then 
											InheritanceHandled = false
											supername = Repository.GetElementByID(conn.SupplierID).Name
										end if
										for each diaos in diagram.DiagramObjects
											if debug then Session.Output("Debug: Connector 2 diaos.ElementID: [" & diaos.ElementID & "]  diaos.Style: ["&diaos.Style&"] ")
											Set super = Repository.GetElementByID(diaos.ElementID)
											if debug then Session.Output("Debug: Connector 3 [" & super.ElementID & "]  Style: ["&super.Type&"] ")
											if debug then Session.Output("Debug: Connector 4 [" & conn.SupplierID & "]  conn.ClientID: ["&conn.ClientID&"] ")
											if super.ElementID <> theElement.ElementID and super.ElementID = conn.SupplierID then
												' Supertype found, recurce into it
												if (PropertiesShown(super, diagram, diaos) ) then
													Session.Output("Info: Diagram ["&diagram.Name&"] supertype ["&super.Name&"] of ["&theElement.Name&"] shown completely in this diagram.")				
													'This Supertype is shown ok
												else
													Session.Output("Info: Diagram ["&diagram.Name&"] supertype ["&super.Name&"] of ["&theElement.Name&"] not shown completely in this diagram.")				
													SuperpropertiesShown = false
												end if
												InheritanceHandled = true
												'exit for? or is it posible to test multiple inheritance sicessfully?
											else
												' Class has subtype, it is not tested
												if debug then Session.Output("Debug: Connector 5 [" & conn.SupplierID & "]  conn.ClientID: ["&conn.ClientID&"] Class has subtype, it is not tested")
											end if
										next
										if not InheritanceHandled then
											'Supertype may not be in this diagram at all
											Session.Output("Info: Diagram ["&diagram.Name&"] supertype ["&supername&"] of ["&theElement.Name&"] not shown completely in this diagram.")				
											SuperpropertiesShown = false
										end if
									else
									end if
								next
							'else
								'no supertypes
							'end if
							'all inherited attributes shown in the class? and no inherited associations?
							if SuperpropertiesShown then PropertiesShown = true
						else
							Session.Output("Info: Diagram ["&diagram.Name&"] not able to show all associations for class ["&theElement.Name&"]")				
						end if
					else
						Session.Output("Info: Diagram ["&diagram.Name&"] Diagram Properties not set up to show any associations for class ["&theElement.Name&"]")				
					end if

					' All model elements are shown in the diagram.
					' But are there any other classes in the diagram who are blocking full view of this element?
					'if ElementBlocked((theElement, diagram, dial) then
						'PropertiesShown = false
					'end if


					if PropertiesShown then
						Session.Output("Info: Diagram ["&diagram.Name&"] OK, shows all attributes and operations in class ["&theElement.Name&"]")				
					end if
					
					' else
						'Session.Output("Info 5 Diagram ["&diagram.Name&"] Roles.....=0 and diagramObject.ShowConstraints=false or InStr(1,diagramObject.Style,'Constraint=1' ) <> 0 or theElement.Constraints.Count > 0.  ")
						'PropertiesShown = false
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


function AssociationsShown(theElement, diagram, diagramObject)
	dim i, roleEndElementID, roleEndElementShown, GeneralizationsFound
	dim dial as EA.DiagramLink
	dim connEl as EA.Connector
	dim conn as EA.Connector
	dim diaoRole as EA.DiagramObject
	AssociationsShown = false
	GeneralizationsFound = 0
	
	for each connEl in theElement.Connectors
		if debug then 
			Session.Output("Debug: Diagram [" & diagram.Name & "]  Type: ["&diagram.Type&"] ")
			Session.Output("Debug: Element [" & theElement.Name & "]  Type: ["&theElement.Type&"] ")
			Session.Output("Debug: connEl [" & connEl.ConnectorID & "]  Type: ["&connEl.Type&"] ")
			Session.Output("Debug: connector SequenceNo [" & connEl.SequenceNo & "]  StateFlags: ["&connEl.StateFlags&"] ")
			Session.Output("Debug: connector Stereotype [" & connEl.Stereotype & "]  StyleEx: ["&connEl.StyleEx&"] ")
		end if
		'test only for Association, Aggregation (+Composition) - leave out Generalization and Realization and the rest
		'TODO connEl.Type
		if connEl.Type = "Generalization" or connEl.Type = "Realization" then
			GeneralizationsFound = GeneralizationsFound + 1
		else
			for each dial in diagram.DiagramLinks
				if debug then 
					Session.Output("Debug: diagram link ID [" & dial.ConnectorID & "] Path: ["&dial.Path&"] ")
					Session.Output("Debug: diagram link Geometry [" & dial.Geometry & "] , Style ["&dial.Style&"] ")
				end if
				Set conn = Repository.GetConnectorByID(dial.ConnectorID)
				if debug then 
					Session.Output("Debug: diagram link connector [" & conn.ConnectorID & "] , StyleEX ["&conn.StyleEX&"] ")
					Session.Output("Debug: diagram link supp,clie [" & conn.SupplierID & "] , ["&conn.ClientID&"] ")
					Session.Output("Debug: diagramID connEl,conn  [" & connEl.DiagramID & "] , ["&conn.DiagramID&"] ")
				end if
				'if connEl.ConnectorID = conn.ConnectorID and connEl.DiagramID = conn.DiagramID then
				if connEl.ConnectorID = conn.ConnectorID then
				'connector has diagramlink so it is shown in this diagram!
				
					'is the class at the other connector end actually shown in this diagram?
				'	roleEndElementShown = false
				'	if conn.ClientID = theElement.ElementID then
				'		roleEndElementID = conn.SupplierID
				'	else
				'		roleEndElementID = conn.ClientID
				'	end if
				'	for each diaoRole in diagram.DiagramObjects
				'		if diaoRole.ElementID = roleEndElementID then
				'				roleEndElementShown = true
				'			exit for
				'		end if
				'	next
				'		
						
						'this role property points to class at supplier end
				'		Session.Output("Debug: looking for element [" & conn.SupplierID & "] at supplier end")
				'		For i = 0 To diaoList.Count - 1
				'			if conn.SupplierID = diaoList.GetByIndex(i) then
				'				'shown at all?
				'				
				'				exit for
				'			end if
				'		next
	 
					if debug then Session.Output("Debug: connector is shown in this diagram - ok ")
					AssociationsShown = true
				else
					if debug then Session.Output("Debug: connector is not shown in this diagram ")
				end if

				'are the connector end elements (role name and multiplicity shown ok?
		'		if conn.ClientID = theElement.ElementID then
		'			if 
		'				AssociationsShown = true
		'				exit for
		'			end if
		'		end if
		'		if conn.SupplierID = theElement.ElementID then
		'			if 
		'				AssociationsShown = true
		'				exit for
		'			end if
		'		end if
		
			next
		end if
	next

	'are there any other connector end elements too close?

	if debug then Session.Output("Debug: connector AssociationsShown = "&AssociationsShown)
	if GeneralizationsFound > 0 and not AssociationsShown then
		if theElement.Connectors.Count = GeneralizationsFound then
			AssociationsShown = true
		end if
	else
		if theElement.Connectors.Count = 0 then
			AssociationsShown = true
		end if
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
			'If not diaoList.ContainsKey(Dobj.InstanceID) Then
				if debug then Session.Output("Debug: add to diaoList:  [" &Dobj.InstanceID& "] [" &Dobj.ElementID&"] diagram: [" &d.Name&"] element: [" &Repository.GetElementByID(Dobj.ElementID).Name& "].  ")
				diaoList.Add Dobj.InstanceID, Dobj.ElementID
				if debug then Session.Output("Debug: add to diagList:  [" &Dobj.InstanceID& "] [" &Dobj.DiagramID&"] diagram: [" &d.Name&"] element: [" &Repository.GetElementByID(Dobj.ElementID).Name& "].  ")
				diagList.Add Dobj.InstanceID, Dobj.DiagramID
			'end if
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