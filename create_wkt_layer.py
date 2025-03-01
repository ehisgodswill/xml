from qgis.core import *
from qgis.utils import iface
from PyQt5.QtGui import QColor
import random


# Function to generate a random color
def random_color():
    # Generate random values for R, G, and B components (between 0 and 255)
    r = random.randint(0, 255)
    g = random.randint(0, 255)
    b = random.randint(0, 255)
    return QColor(r, g, b)


# Create a new empty vector layer to store the geometries
def create_wkt_layer(input_layer, wkt_field):
    # Create a memory layer to store the output geometries
    output_layer = QgsVectorLayer("Polygon?crs=EPSG:4326", "WKT_to_Geometry_Layer", "memory")
    
    # Define the layer's fields by copying the input layer's fields (without the WKT field)
    fields = input_layer.fields()
    output_layer.dataProvider().addAttributes(fields)
    output_layer.updateFields()

    # Start an edit session to add geometries
    output_layer.startEditing()

    # Go through each feature in the input layer and convert the WKT to a geometry
    for feature in input_layer.getFeatures():
        # Get the WKT from the specified field and convert it to string
        wkt = str(feature[wkt_field])
        print(wkt)

        # Convert WKT to geometry
        geom = QgsGeometry.fromWkt(wkt)

        # Create a new feature with the geometry and attributes
        new_feature = QgsFeature()
        new_feature.setGeometry(geom)
        new_feature.setAttributes(feature.attributes())
        
        # Add the new feature to the output layer
        output_layer.dataProvider().addFeature(new_feature)
        
        # Enable labeling on the new layer
        label_settings = QgsPalLayerSettings()
        label_settings.fieldName = 'FarmerID'  # The field you want to display as a label (e.g., an ID or name)

        # Set label placement (optional, here we use "centroid")
        #label_settings.placement = QgsPalLayerSettings.OverPoint

        # Create a label engine and set the settings
        label = QgsVectorLayerSimpleLabeling(label_settings)
        output_layer.setLabelsEnabled(True)  # Enable the labels on the layer
        output_layer.setLabeling(label)  # Apply the labeling settings

    
    # Get the layer's symbol
    symbol = output_layer.renderer().symbol().symbolLayer(0)
    # Set the fill color to transparent (alpha = 0)
    symbol.setColor(QColor(0, 0, 0, 0))  # (Red, Green, Blue, Alpha)
    
    # Set the outline (stroke) color to black (or any color you prefer)
    symbol.setStrokeColor(random_color())  # Black outline
    
    # Update the layer's symbology
    output_layer.triggerRepaint()
    
    # Commit changes and finalize the layer
    output_layer.commitChanges()
    
    # Add the new layer to the map
    QgsProject.instance().addMapLayer(output_layer)

    return output_layer


# Usage Example
# Replace 'input_layer' with your actual layer and 'Polygon_WKT' with the correct WKT field name
input_layer = iface.activeLayer()  # The currently selected layer in QGIS
wkt_field = 'Cordinates'  # The column containing WKT data

# Call the function to create the layer with geometries
create_wkt_layer(input_layer, wkt_field)
