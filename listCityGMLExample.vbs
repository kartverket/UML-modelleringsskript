option explicit

!INC Local Scripts.EAConstants-VBScript

' script:		listCityGMLExample
' purpose:		Generates a CityGML example object
' version:		2019-06-07 
' author: 		Kent Jonsrud

sub listCityGMLExample()
	' Show and clear the script output windows
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
	Repository.CreateOutputTab "Error"
	Repository.ClearOutput "Error"
	Dim indent
	indent = "  					"
	Dim height
	height = 6.0


	SessionOutput("<?xml version=""1.0"" encoding=""utf-8""?>")
	SessionOutput("<CityModel gml:id=""Norwegian CityGML Example""")
	SessionOutput("  xmlns:app=""http://www.opengis.net/citygml/appearance/3.0""")
	SessionOutput("  xmlns:brid=""http://www.opengis.net/citygml/bridge/3.0""")
	SessionOutput("  xmlns:bldg=""http://www.opengis.net/citygml/building/3.0""")
	SessionOutput("  xmlns:frn=""http://www.opengis.net/citygml/cityfurniture/3.0""")
	SessionOutput("  xmlns:grp=""http://www.opengis.net/citygml/cityobjectgroup/3.0""")
	SessionOutput("  xmlns:con=""http://www.opengis.net/citygml/construction/3.0""")
	SessionOutput("  xmlns:pcl=""http://www.opengis.net/citygml/pointcloud/3.0""")
	SessionOutput("  xmlns:core=""http://www.opengis.net/citygml/3.0""")
	SessionOutput("  xmlns:dyn=""http://www.opengis.net/citygml/dynamizer/3.0""")
	SessionOutput("  xmlns:gen=""http://www.opengis.net/citygml/generics/3.0""")
	SessionOutput("  xmlns:luse=""http://www.opengis.net/citygml/landuse/3.0""")
	SessionOutput("  xmlns:dem=""http://www.opengis.net/citygml/relief/3.0""")
	SessionOutput("  xmlns:tran=""http://www.opengis.net/citygml/transportation/3.0""")
	SessionOutput("  xmlns:tun=""http://www.opengis.net/citygml/tunnel/3.0""")
	SessionOutput("  xmlns:veg=""http://www.opengis.net/citygml/vegetation/3.0""")
	SessionOutput("  xmlns:vers=""http://www.opengis.net/citygml/versioning/3.0""")
	SessionOutput("  xmlns:wtr=""http://www.opengis.net/citygml/waterbody/3.0""")
	SessionOutput("  xmlns:tsml=""http://www.opengis.net/tsml/1.0""")
	SessionOutput("  xmlns:sos=""http://www.opengis.net/sos/2.0""")
	SessionOutput("  xmlns:xAL=""urn:oasis:names:tc:ciq:xsdschema:xAL:2.0""")
	SessionOutput("  xmlns:xlink=""http://www.w3.org/1999/xlink""")
	SessionOutput("  xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""")
	SessionOutput("  xmlns:gml=""http://www.opengis.net/gml/3.2""")
	SessionOutput("  xmlns=""http://www.opengis.net/citygml/3.0""")
	SessionOutput("  xsi:schemaLocation=""http://www.opengis.net/citygml/appearance/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/appearance.xsd")
	SessionOutput("     http://www.opengis.net/citygml/bridge/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/bridge.xsd")
	SessionOutput("     http://www.opengis.net/citygml/building/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/building.xsd")
	SessionOutput("     http://www.opengis.net/citygml/cityfurniture/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/cityFurniture.xsd")
	SessionOutput("     http://www.opengis.net/citygml/cityobjectgroup/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/cityObjectGroup.xsd")
	SessionOutput("     http://www.opengis.net/citygml/construction/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/construction.xsd")
	SessionOutput("     http://www.opengis.net/citygml/pointcloud/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/pointCloud.xsd")
	SessionOutput("     http://www.opengis.net/citygml/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/cityGMLBase.xsd")
	SessionOutput("     http://www.opengis.net/citygml/dynamizer/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/dynamizer.xsd")
	SessionOutput("     http://www.opengis.net/citygml/generics/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/generics.xsd")
	SessionOutput("     http://www.opengis.net/citygml/landuse/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/landUse.xsd")
	SessionOutput("     http://www.opengis.net/citygml/relief/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/relief.xsd")
	SessionOutput("     http://www.opengis.net/citygml/transportation/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/transportation.xsd")
	SessionOutput("     http://www.opengis.net/citygml/tunnel/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/tunnel.xsd")
	SessionOutput("     http://www.opengis.net/citygml/vegetation/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/vegetation.xsd")
	SessionOutput("     http://www.opengis.net/citygml/versioning/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/versioning.xsd")
	SessionOutput("     http://www.opengis.net/citygml/waterbody/3.0")
	SessionOutput("     http://www.3dcitydb.org/citygml3/2018-12-06/xsds/waterBody.xsd")
	SessionOutput("     http://www.opengis.net/tsml/1.0")
	SessionOutput("     http://schemas.opengis.net/tsml/1.0/timeseriesML.xsd")
	SessionOutput("     http://www.opengis.net/sos/2.0")
	SessionOutput("     http://schemas.opengis.net/sos/2.0/sosGetObservation.xsd")
	SessionOutput("     urn:oasis:names:tc:ciq:xsdschema:xAL:2.0")
	SessionOutput("     http://schemas.opengis.net/citygml/xAL/xAL.xsd"">")
	SessionOutput("  		<gml:name>Norwegian CityGML Example</gml:name>")
	SessionOutput("  		<gml:boundedBy>")
	SessionOutput("  			<gml:Envelope srsName=""http://www.opengis.net/def/crs/epsg/0/5972"" srsDimension=""3"">")
	SessionOutput("  				<gml:lowerCorner>4490655.500 5322005.280 548.470</gml:lowerCorner>")
	SessionOutput("  				<gml:upperCorner>4490671.290 5322017.800 557.020</gml:upperCorner>")
	SessionOutput("  			</gml:Envelope>")
	SessionOutput("  		</gml:boundedBy>")

	SessionOutput("  		<cityObjectMember>")

	SessionOutput("  			<bldg:Building>")


	SessionOutput("  				<gml:name>PropertyParcel_KentJonsrud2019</gml:name>")
	SessionOutput("  					<creationDate>1987-02-18T00:00:00</creationDate>")
	SessionOutput("  					<externalReference>")
	SessionOutput("  						<ExternalReference>")
	SessionOutput("  							<targetResource>http://data.geonorge.no/matrikkel#PropertyParcel_KentJonsrud2019</targetResource>")
	SessionOutput("  							<informationSystem>http://matrikkelen.no</informationSystem>")
	SessionOutput("  						</ExternalReference>")
	SessionOutput("  					</externalReference>")
	SessionOutput("  					<genericAttribute>")
	SessionOutput("  						<gen:StringAttribute>")
	SessionOutput("  							<gen:name>Hack1</gen:name>")
	SessionOutput("  							<gen:value>2014-07-28</gen:value>")
	SessionOutput("  						</gen:StringAttribute>")
	SessionOutput("  					</genericAttribute>")
	SessionOutput("  					<genericAttribute>")
	SessionOutput("  						<gen:StringAttribute>")
	SessionOutput("  							<gen:name>Hack2</gen:name>")
	SessionOutput("  							<gen:value>09175128</gen:value>")
	SessionOutput("  						</gen:StringAttribute>")
	SessionOutput("  					</genericAttribute>")
	

	call generateCityGMLSolidExample("PropertyParcel", indent, height)


	SessionOutput("  					<con:heightAboveGround>")
	SessionOutput("  						<con:HeightAboveGround>")
	SessionOutput("  							<con:heightReference>highestRoofEdge</con:heightReference>")
	SessionOutput("  							<con:lowReference>lowestGroundPoint</con:lowReference>")
	SessionOutput("  							<con:status>measured</con:status>")
	SessionOutput("  							<con:value uom=""urn:adv:uom:m"">4.55</con:value>")
	SessionOutput("  						</con:HeightAboveGround>")
	SessionOutput("  					</con:heightAboveGround>")
	SessionOutput("  					<bldg:function>31001_9998</bldg:function>")
	SessionOutput("  					<bldg:roofType>3100</bldg:roofType>")
	SessionOutput("  					<bldg:address>")
	SessionOutput("  						<Address gml:id=""Address.06Bygning_KentJonsrud2019"">")
	SessionOutput("  							<xalAddress>")
	SessionOutput("  								<xAL:AddressDetails>")
	SessionOutput("  									<xAL:Country>")
	SessionOutput("  										<xAL:CountryName>Norway</xAL:CountryName>")
	SessionOutput("  										<xAL:Locality Type=""Town"">")
	SessionOutput("  											<xAL:LocalityName>Svensrud</xAL:LocalityName>")
	SessionOutput("  											<xAL:Thoroughfare Type=""Street"">")
	SessionOutput("  												<xAL:ThoroughfareName>Meieriveien 24</xAL:ThoroughfareName>")
	SessionOutput("  											</xAL:Thoroughfare>")
	SessionOutput("  										</xAL:Locality>")
	SessionOutput("  									</xAL:Country>")
	SessionOutput("  								</xAL:AddressDetails>")
	SessionOutput("  							</xalAddress>")
	SessionOutput("  						</Address>")
	SessionOutput("  					</bldg:address>")
	SessionOutput("  			</bldg:Building>")
	SessionOutput("  		</cityObjectMember>")

	SessionOutput("</CityModel>")		
end sub


sub generateCityGMLSolidExample(elementName, indent, height)

'	start with a small surface with different elevations in each coordinate position, and with no interiors

'	test whether the whole surface is in a single plane, and if so consider skipping the center point part(?)

'	split the surface in subsurfaces where it is possible to generate a central point thet has direct vision to all its perimeter points(?)

'	find the central point and the mean height

'	construct the set of floor surface slices from the central point to every two consecutive points on the perimeter 

'	erect a set of sheet piles from two and two perimeter points up the given height above the floor

'	copy the reverse of the floor as a roof and add the given height to it

'	.

	Dim pnteller, cuteller, suteller, soteller, obteller
	pnteller = 0
	cuteller = 0
	suteller = 0
	soteller = 0
	obteller = 0

'	hardcode a totally random surface to start with
'   srsName="urn:ogc:def:crs:EPSG::5972" srsDimension="3">
'	568444.03 6661981.48 89.20
'	568506.41 6662009.49 91.20
'	568525.84 6661998.97 90.80
'	568529.64 6662001.85 91.00
'	568535.02 6662054.94 91.50
'	568476.33 6662067.85 90.50
'	568466.50 6662054.49 90.50
'	568444.03 6661981.48 89.20
	dim s1(7,2), c1(2), z1(2), h1, posNum, i
	h1 = height
	s1(0,0) = 568444.03
	s1(0,1) = 6661981.48 
	s1(0,2) = 89.20
	s1(1,0) = 568506.41 
	s1(1,1) = 6662009.49 
	s1(1,2) = 91.20
	s1(2,0) = 568525.84 
	s1(2,1) = 6661998.97
	s1(2,2) = 90.80
	s1(3,0) = 568529.64 
	s1(3,1) = 6662001.85 
	s1(3,2) = 91.00
	s1(4,0) = 568535.02 
	s1(4,1) = 6662054.94 
	s1(4,2) = 91.50
	s1(5,0) = 568476.33 
	s1(5,1) = 6662067.85 
	s1(5,2) = 90.50
	s1(6,0) = 568466.50 
	s1(6,1) = 6662054.49 
	s1(6,2) = 90.50
	s1(7,0) = 568444.03 
	s1(7,1) = 6661981.48 
	s1(7,2) = 89.20
	posNum = 8
	z1(0) =0.0
	z1(1) =0.0
	z1(2) =0.0
	
'	calculate the central point and mean height

	for i = 0 to posNum - 2
		z1(0) = z1(0) + s1(i,0)
		z1(1) = z1(1) + s1(i,1)
		z1(2) = z1(2) + s1(i,2)
	next
	
	c1(0) = Round( z1(0) / (posNum - 1),2)
	c1(1) = Round( z1(1) / (posNum - 1),2)
	c1(2) = Round( z1(2) / (posNum - 1),2)
	
'	start the xml structure of the gml:Solid
	SessionOutput(indent & "<lod2Solid>")
	SessionOutput(indent & "	<gml:Solid gml:id=""" & elementName & """>")
	SessionOutput(indent & "		<gml:exterior>")
	SessionOutput(indent & "			<gml:Shell>")
	
	for i = 0 to posNum - 2
		soteller = soteller + 1	
		SessionOutput(indent & "				<gml:surfaceMember xlink:href=""#Floor." & soteller & "P""/>")
	next
	
	for i = 0 to posNum - 2
		soteller = soteller + 1	
		SessionOutput(indent & "				<gml:surfaceMember xlink:href=""#Wall." & soteller & "P""/>")
	next

	for i = 0 to posNum - 2
		soteller = soteller + 1	
		SessionOutput(indent & "				<gml:surfaceMember xlink:href=""#Roof." & soteller & "P""/>")
	next

	SessionOutput(indent & "			</gml:Shell>")
	SessionOutput(indent & "		</gml:exterior>")
	SessionOutput(indent & "	</gml:Solid>")
	SessionOutput(indent & "</lod2Solid>")

'	generate the floor tiles
	for i = 0 to posNum - 2
		suteller = suteller + 1
		SessionOutput(indent & "<boundary>")
		SessionOutput(indent & "	<con:GroundSurface gml:id=""Floor." & suteller & "C"">")
		SessionOutput(indent & "		<lod2MultiSurface>")
		SessionOutput(indent & "			<gml:MultiSurface gml:id=""Floor." & suteller & "M"">")
		SessionOutput(indent & "				<gml:surfaceMember>")
		SessionOutput(indent & "					<gml:Polygon gml:id=""Floor." & suteller & "P"">")
		SessionOutput(indent & "						<gml:exterior>")
		SessionOutput(indent & "							<gml:LinearRing>")
				
		SessionOutput(indent & "								<gml:posList>" & c1(0) & " " & c1(1) & " " & c1(2) & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2) & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & " " & c1(0) & " " & c1(1) & " " & c1(2) & "</gml:posList>")
		
		SessionOutput(indent & "							</gml:LinearRing>")
		SessionOutput(indent & "						</gml:exterior>")
		SessionOutput(indent & "					</gml:Polygon>")
		SessionOutput(indent & "				</gml:surfaceMember>")
		SessionOutput(indent & "			</gml:MultiSurface>")
		SessionOutput(indent & "		</lod2MultiSurface>")
		SessionOutput(indent & "	</con:GroundSurface>")
		SessionOutput(indent & "</boundary>")
	next
	
'	erect the sheet piles
	for i = 0 to posNum - 2
		suteller = suteller + 1
		SessionOutput(indent & "<boundary>")
		SessionOutput(indent & "	<con:WallSurface gml:id=""Wall." & suteller & "C"">")
		SessionOutput(indent & "		<lod2MultiSurface>")
		SessionOutput(indent & "			<gml:MultiSurface gml:id=""Wall." & suteller & "M"">")
		SessionOutput(indent & "				<gml:surfaceMember>")
		SessionOutput(indent & "					<gml:Polygon gml:id=""Wall." & suteller & "P"">")
		SessionOutput(indent & "						<gml:exterior>")
		SessionOutput(indent & "							<gml:LinearRing>")
				
		SessionOutput(indent & "								<gml:posList>" & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2) & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2)+h1 & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2)+h1 & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2) & "</gml:posList>")
		
		SessionOutput(indent & "							</gml:LinearRing>")
		SessionOutput(indent & "						</gml:exterior>")
		SessionOutput(indent & "					</gml:Polygon>")
		SessionOutput(indent & "				</gml:surfaceMember>")
		SessionOutput(indent & "			</gml:MultiSurface>")
		SessionOutput(indent & "		</lod2MultiSurface>")
		SessionOutput(indent & "	</con:WallSurface>")
		SessionOutput(indent & "</boundary>")
	next

'	generate the roof tiles
	for i = 0 to posNum - 2
		suteller = suteller + 1
		SessionOutput(indent & "<boundary>")
		SessionOutput(indent & "	<con:RoofSurface gml:id=""Roof." & suteller & "C"">")
		SessionOutput(indent & "		<lod2MultiSurface>")
		SessionOutput(indent & "			<gml:MultiSurface gml:id=""Roof." & suteller & "M"">")
		SessionOutput(indent & "				<gml:surfaceMember>")
		SessionOutput(indent & "					<gml:Polygon gml:id=""Roof." & suteller & "P"">")
		SessionOutput(indent & "						<gml:exterior>")
		SessionOutput(indent & "							<gml:LinearRing>")
				
		SessionOutput(indent & "								<gml:posList>" & c1(0) & " " & c1(1) & " " & c1(2)+h1 & " " & s1(i,0) & " " & s1(i,1) & " " & s1(i,2)+h1 & " " & s1(i+1,0) & " " & s1(i+1,1) & " " & s1(i+1,2)+h1 & " " & c1(0) & " " & c1(1) & " " & c1(2)+h1 & "</gml:posList>")
		
		SessionOutput(indent & "							</gml:LinearRing>")
		SessionOutput(indent & "						</gml:exterior>")
		SessionOutput(indent & "					</gml:Polygon>")
		SessionOutput(indent & "				</gml:surfaceMember>")
		SessionOutput(indent & "			</gml:MultiSurface>")
		SessionOutput(indent & "		</lod2MultiSurface>")
		SessionOutput(indent & "	</con:RoofSurface>")
		SessionOutput(indent & "</boundary>")
	next

end sub

sub SessionOutput(text)
	Session.Output(text)
end sub

listCityGMLExample
