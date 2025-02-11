import pandas as pd
import pyproj
from pyproj import CRS, Transformer
from arcgis.gis import GIS
import os
import datetime
from pathlib import Path
from config import user, password

# User login info
gis = GIS('https://usys-ethz.maps.arcgis.com', user, password)

# Account and Hosted Feature Layer Name, with Feature Layer structure
username = 'B.Baeume'
groups = ['12', '34']
year = str(datetime.date.today().year)

# Ask the user for the folder path
folder_path = input("Please enter the folder path where you want to save the files: ")

# Download all data
def export_data(year, item, folder_path):
    downloadFormat = 'File Geodatabase'
    try:
        result = item.export(year + '_{}'.format(item.title), downloadFormat)
        file_path = result.download(folder_path)
        result.delete()
        print("Successfully downloaded " + item.title)
        return file_path
    except Exception as e:
        print(e)
        return None

def export_dataworkshop(allFeatureLayer):
    # Define Input and Output Coordinate System
    inProj = CRS('epsg:3857')
    outProj = CRS('epsg:4326')
    transformer = Transformer.from_crs(inProj, outProj)

    # Counter for Group Name and empty Dataframe
    counter = 0
    dfAll = pd.DataFrame()

    # Loop through all Feature layers
    for featureLayer in allFeatureLayer:
        # Search all Features in a Feature Layer
        print(featureLayer.properties.name)
        featureSet = featureLayer.query()
        if not featureSet:
            counter += 1
            print('Feature Layer {} is empty'.format(featureLayer.properties.name))
            continue
        else:
            print('Feature Layer {} is being processed'.format(counter))
            # Create Panda Dataframe from all Features
            df = featureSet.sdf
            # Create empty X and Y List
            newX = []
            newY = []
            counter += 1
            # Convert each Coordinate to the new projection
            for f in featureSet:
                x1, y1 = f.geometry["x"], f.geometry["y"]
                x2, y2 = transformer.transform(x1, y1)
                newX.append(x2)
                newY.append(y2)

            # Insert new Column X and Y in Data Frame
            df.insert(len(df.columns), 'X', newY)
            df.insert(len(df.columns), 'Y', newX)

            # Insert Groupname
            df.insert(len(df.columns), 'Gruppe', '{}{}'.format(group, f"{counter:02d}"))
            print(df)

        # Append Dataframe 
        dfAll = pd.concat([dfAll, df], ignore_index=True)

    return dfAll

for group in groups:
    print(group)
    # Run through the groups defined to download data from all layers
    title = 'Baeume_' + group
    print('Starting:')
    print(title)

    # Search hosted Feature Layer
    hostedFeatureLayer = gis.content.search('owner: {} AND title: {} AND type: Feature Service'.format(username, title))
    print(hostedFeatureLayer)

    # Store first Item of hostedFeature Search
    item = hostedFeatureLayer[0]
    print(item)
    file_path = export_data(year, item, folder_path)
    print(f"Downloaded file path: {file_path}")
    
    # Store all Feature Layers of the item
    allFeatureLayer = item.layers
    print('FeatureLayer')
    print(allFeatureLayer)
    dfAll = export_dataworkshop(allFeatureLayer)
    
    # Output the combined DataFrame to an Excel file in the specified folder path
    output_excel_path = os.path.join(folder_path, f"{title}.xlsx")
    try:
        dfAll.to_excel(output_excel_path, index=False)
        print(f"Data successfully written to {output_excel_path}")
    except Exception as e:
        print(e)
