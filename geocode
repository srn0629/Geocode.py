#Savanna Nagorski
#University of Washington Tacoma
#Masters of Science in Geospatial Technologies
#Campus Planning and Real Estate / Institutional Research
#March 2016


import arcpy
from arcpy import env

env.workspace = 'FOLDER FILEPATH'
arcpy.env.overwriteOutput = True

#Convert and make Excel file accessible in ArcMap for geoprocessing
TParks='TParks'
arcpy.ExcelToTable_conversion("TParks.xlsx", TParks, "Tacoma_parks")
arcpy.TableToDBASE_conversion(TParks, 'Project_Data.gdb')

#Build parameters and begin geocoding addresses
for t in TParks:
	table='/Project_Data.gdb/TParks'
	locator="/Project_Data.gdb/ADDRESS_LOCATOR"
	fields="Street Address VISIBLE NONE; City City VISIBLE NONE; ZIP ZIP VISIBLE NONE"
	result='TParks_GR.shp'
	arcpy.GeocodeAddresses_geocoding(table, locator, fields, result, "STATIC") 
print "Addresses Geocoded!"

#Update attribute data for needs of the project
for u in result:
	result2='TParks_update.shp'
	dropFields = ["Status", "Score", "Match_type", "Match_Addr", "Side", "Ref_ID", "User_fld", "Addr_type", "ARC_Street", "ARC_City", "ARC_ZIP"]
	arcpy.Select_analysis(result, result2, "\"Score\" >  0")	#Select rows that came back with matched addresses
	arcpy.DeleteField_management(result2, dropFields)			#Delete Unwanted Fields
print "Data table updated!"

#Project (Spatial) Data
for p in result3:
	TacomaParks='TParks_Project.shp'
	sr = arcpy.SpatialReference('WGS 1984')
	arcpy.Project_management(result2, TacomaParks, sr)
print "Parks Projected!"

#Add XY Coordinates and convert table back to Excel
for x in TacomaParks:
	arcpy.AddXY_management(TacomaParks)
	arcpy.TableToExcel_conversion(TacomaParks, 'TacomaParks.xls')
print "XY added to data and converted back to Excel!"
