option explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name:
' Author: Tor Kjetil Nilsen, Arkitektum AS
' Purpose: Validate use of incorrect or disconnected types
' Date: 2015-12-30

sub OnProjectBrowserScript()
	
	
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	select case treeSelectedType				
		case otPackage
		' Code for when a package is selected
			dim thePackage as EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			Repository.WriteOutput "Script", "Start validating [" & thePackage.Name & "] " & Now,0
			ValidatePackage(thePackage)
			Repository.WriteOutput "Script", "End validating [" & thePackage.Name & "] " & Now,0
		case else
			' Error message
			Session.Prompt "This script does not support items of this type.", promptOK
			
	end select
	
end sub

OnProjectBrowserScript

sub ValidatePackage(p)
	dim stringTypes
	Set stringTypes = CreateObject("System.Collections.ArrayList")
	stringTypes.Add "char"
	stringTypes.Add "character"
	stringTypes.Add "string"
	stringTypes.Add "charcterstring"

	
	dim intTypes
	Set intTypes = CreateObject("System.Collections.ArrayList")
	intTypes.Add "int"
	intTypes.Add "short"
	intTypes.Add "long"
	
	dim realTypes
	Set realTypes = CreateObject("System.Collections.ArrayList")
	realTypes.Add "double"
	realTypes.Add "float"

	dim boolTypes
	Set boolTypes = CreateObject("System.Collections.ArrayList")
	boolTypes.Add "bolsk"
	boolTypes.Add "boolsk"
	
	dim BasicTypes
	Set BasicTypes = CreateObject("Scripting.Dictionary")
	with BasicTypes
	.Add "characterstring" , "CharacterString"
	.Add "integer", "Integer"
	.Add "real", "Real"
	.Add "decimal", "Decimal"
	.Add "date", "Date"
	.Add "datetime", "DateTime"
	.Add "boolean", "Boolean"
	.Add "number", "Number"
	.Add "time", "Time"
	.Add "vector", "Vector"
	.Add "genericname", "GenericName"
	.Add "localname", "LocalName"
	.Add "scopename", "ScopeName"
	.Add "length", "Length"
	.Add "distance", "Distance"
	.Add "angle", "Angle"
	.Add "speed", "Speed"
	.Add "scale", "Scale"
	.Add "area", "Area"
	.Add "volume", "Volume"
	.Add "measure", "Measure"
	.Add "sign", "Sign"
	.Add "unitofmeasure", "UnitOfMeasure"
	
	.Add "flate", "Flate"
	.Add "kurve", "Kurve"
	.Add "punkt", "Punkt"
	.Add "sverm", "Sverm"
	
	.Add "gm_object", "GM_Object"
	.Add "gm_primitive", "GM_Primitive"
	.Add "directposition", "DirectPosition"
	.Add "gm_position", "GM_Position"
	.Add "gm_pointarray", "GM_PointArray"
	.Add "gm_point", "GM_Point"
	.Add "gm_curve", "GM_Curve"
	.Add "gm_surface", "GM_Surface"
	.Add "gm_polyhedralsurface", "GM_PolyhedralSurface"
	.Add "gm_triangulatedsurface","GM_TriangulatedSurface"
	.Add "gm_tin","GM_Tin"
	.Add "gm_solid","GM_Solid"
	.Add "gm_orientablecurve","GM_OrientableCurve"
	.Add "gm_orientablesurface","GM_OrientableSurface"
	.Add "gm_ring","GM_Ring"
	.Add "gm_shell","GM_Shell"
	.Add "gm_compositepoint","GM_CompositePoint"
	.Add "gm_compositecurve","GM_CompositeCurve"
	.Add "gm_compositesurface","GM_CompositeSurface"
	.Add "gm_compositesolid","GM_CompositeSolid"
	.Add "gm_complex","GM_Complex"
	.Add "gm_aggregate","GM_Aggregate"
	.Add "gm_multipoint","GM_MultiPoint"
	.Add "gm_multicurve","GM_MultiCurve"
	.Add "gm_multisurface","GM_MultiSurface"
	.Add "gm_multisolid","GM_MultiSolid"
	.Add "gm_multiprimitive", "GM_MultiPrimitive"
	.Add "gm_curvesegment", "GM_CurveSegment"
	.Add "gm_arc", "GM_Arc"
	.Add "gm_arcbybulge", "GM_ArcByBulge"
	.Add "gm_arcstring", "GM_ArcString"
	.Add "gm_arcstringbybulge", "GM_ArcStringByBulge"
	.Add "gm_bezier", "GM_Bezier"
	.Add "gm_bsplinecurve", "GM_BsplineCurve"
	.Add "gm_circle", "GM_Circle"
	.Add "gm_clothoid", "GM_Clothoid"
	.Add "gm_cubicspline", "GM_CubicSpline"
	.Add "gm_geodesicstring", "GM_GeodesicString"
	.Add "gm_linestring", "GM_LineString"
	.Add "gm_offsetcurve", "GM_OffsetCurve"
	.Add "gm_surfacepatch", "GM_SurfacePatch"
	.Add "gm_griddedsurface", "GM_GriddedSurface"
	.Add "gm_parametriccurvesurface", "GM_ParametricCurveSurface"
	.Add "gm_cone", "GM_Cone"
	.Add "gm_cylinder", "GM_Cylinder"
	.Add "gm_geodesic", "GM_Geodesic"
	.Add "gm_polygon", "GM_Polygon"
	.Add "gm_sphere", "GM_Sphere"
	.Add "gm_triangle", "GM_Triangle"
	.Add "tp_object", "TP_Object"
	.Add "tp_node", "TP_Node"
	.Add "tp_edge", "TP_Edge"
	.Add "tp_face", "TP_Face"
	.Add "tp_solid", "TP_Solid"
	.Add "tp_directednode", "TP_DirectedNode"
	.Add "tp_directededge", "TP_DirectedEdge"
	.Add "tp_directedface", "TP_DirectedFace"
	.Add "tp_directedsolid", "TP_DirectedSolid"
	.Add "tp_complex", "TP_Complex"
	.Add "tm_object", "TM_Object"
	.Add "tm_complex", "TM_Complex"
	.Add "tm_geometricprimitive", "TM_GeometricPrimitive"
	.Add "tm_instant", "TM_Instant"
	.Add "tm_period", "TM_Period"
	.Add "tm_topologicalcomplex", "TM_TopologicalComplex"
	.Add "tm_topologicalprimitive", "TM_TopologicalPrimitive"
	.Add "tm_node", "TM_Node"
	.Add "tm_edge", "TM_Edge"
	.Add "tm_periodduration", "TM_PeriodDuration"
	.Add "tm_intervallength", "TM_IntervalLength"
	.Add "tm_duration", "TM_Duration"
	.Add "tm_position", "TM_Position"
	.Add "tm_indeterminatevalue", "TM_IndeterminateValue"
	.Add "tm_coordinate", "TM_Coordinate"
	.Add "tm_caldate", "TM_CalDate"
	.Add "tm_clocktime", "TM_ClockTime"
	.Add "tm_dateandtime", "TM_DateAndTime"
	.Add "tm_calendar", "TM_Calendar"
	.Add "tm_calendarera", "TM_CalendarEra"
	.Add "tm_clock", "TM_Clock"
	.Add "tm_coordinatesystem", "TM_CoordinateSystem"
	.Add "tm_ordinalreferencesystem", "TM_OrdinalReferenceSystem"
	.Add "tm_ordinalera", "TM_OrdinalEra"
	.Add "sc_crs", "SC_CRS"
	.Add "si_locationinstance", "SI_LocationInstance"
	.Add "cv_coverage", "CV_Coverage"
	.Add "cv_continuouscoverage", "CV_ContinuousCoverage"
	.Add "cv_discretecoverage", "CV_DiscreteCoverage"
	.Add "cv_discretepointcoverage", "CV_DiscretePointCoverage"
	.Add "cv_discretecurvecoverage", "CV_DiscreteCurveCoverage"
	.Add "cv_discretesurfacecoverage", "CV_DiscreteSurfaceCoverage"
	.Add "cv_discretesolidcoverage", "CV_DiscreteSolidCoverage"
	.Add "cv_discretegridpointcoverage", "CV_DiscreteGridPointCoverage"

end with
	
	dim el as EA.Element
	for each el In p.elements
		if el.Stereotype <> "codeList" and el.Stereotype <> "CodeList" and el.Stereotype <> "enumeration" then
			dim att as EA.Attribute
			for each att in el.Attributes
				if att.ClassifierID = 0 then
				    if BasicTypes.Exists(LCase(att.Type)) then
						if att.Type <> BasicTypes.Item(LCase(att.Type)) then
							att.Type = BasicTypes.Item(LCase(att.Type))
							att.Update()
							Repository.WriteOutput "Script", "[FIXED] Class [" & el.Name & "]\Attribute [" & att.Name & "] has known type [" & att.Type & "] but wrong case. Changed to correct case [" & BasicTypes.Item(LCase(att.Type)) & "].",0
						end if
					elseif Len(att.Type) = 0 then
						Repository.WriteOutput "Script", "[ERROR] Class [" & el.Name & "]\Attribute [" & att.Name & "] has no type.",0
					elseif stringTypes.IndexOf(LCase(att.Type),0) <> -1 then
						att.Type = "CharacterString"
						att.Update()
						Repository.WriteOutput "Script", "[FIXED] Class [" & el.Name & "]\Attribute [" & att.Name & "] with unknown type [" & att.Type & "]. Changed to type [CharacterString].",0
					elseif intTypes.IndexOf(LCase(att.Type),0) <> -1 then
						att.Type = "Integer"
						att.Update()
						Repository.WriteOutput "Script", "[FIXED] Class [" & el.Name & "]\Attribute [" & att.Name & "] with unknown type [" & att.Type & "]. Changed to type [Integer].",0
					elseif realTypes.IndexOf(LCase(att.Type),0) <> -1 then
						att.Type = "Real"
						att.Update()
						Repository.WriteOutput "Script", "[FIXED] Class [" & el.Name & "]\Attribute [" & att.Name & "] with unknown type [" & att.Type & "]. Changed to type [Real].",0
					elseif boolTypes.IndexOf(LCase(att.Type),0) <> -1 then
						att.Type = "Boolean"
						att.Update()
						Repository.WriteOutput "Script", "[FIXED] Class [" & el.Name & "]\Attribute [" & att.Name & "] with unknown type [" & att.Type & "]. Changed to type [Boolean].",0
					else
						dim classifierid
						classifierid = SearchTypeInPackage(att.Type, p)
						if classifierid <> 0 then
							att.ClassifierID = classifierid
							att.Update()
							Repository.WriteOutput "Script", "[FIXED] Class [" & el.Name & "]\Attribute [" & att.Name & "] with type [" & att.Type & "] is now reconnected to class [" & att.Type & "].",0
						else
							Repository.WriteOutput "Script", "[ERROR] Class [" & el.Name & "]\Attribute [" & att.Name & "] with type [" & att.Type & "] is not connected to class [" & att.Type & "]. Please reconnect manually to correct class.",0
						end if
					end if
				end if
			next
		end if
	next
	
	dim subP as EA.Package
	for each subP in p.packages
	    ValidatePackage(subP)
	next
end sub

function SearchTypeInPackage(classifierType , p)
	SearchTypeInPackage = 0
	dim el as EA.Element
	for each el In p.elements
		if el.Name = classifierType then
			SearchTypeInPackage = el.ClassifierID
			exit function
		end if
	next
end function