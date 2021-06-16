import io
import sys
import pandas as pd
import numpy as np

import folium
from routingpy import Graphhopper
from geopy.geocoders import Nominatim

from PyQt5 import QtWidgets, QtWebEngineWidgets

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    
    # create map object
    m = folium.Map(
        location=[47.392466, 8.612041], tiles="Stamen Toner", zoom_start=4
    )

    plant_geo = [8.610720, 47.399280]
    vendor_data = pd.read_excel('Pyqt5_6 tests/test2/vendor_zipcodes.xlsx', sheet_name='Sheet1')

    # find geolocation of vendors based on postcodes provided
    geolocator = Nominatim(user_agent="tco")
    vendor_geocodes = []
    for postcode in vendor_data['Vendor - Postal code'].tolist():
        try:
            location = geolocator.geocode(postcode)
            vendor_geocodes.append([location.latitude,location.longitude])
        except:
            vendor_geocodes.append(np.NaN)
    vendor_data['Geocodes'] = vendor_geocodes
    vendor_data = vendor_data.dropna(subset=['Geocodes'])
    vendor_data_dict = vendor_data.set_index('Vendor code').to_dict(orient='index')
    routes = {}
    route_index = 0

    for vendor in vendor_data_dict:
        
        # calculate the optimal driving route between the plant and each vendor using Graphhopper api
        coords = [vendor_data_dict[vendor]['Geocodes'][::-1], plant_geo]
        client = Graphhopper(api_key='your api key goes here')
        try:
            route_index += 1
            route = client.directions(locations=coords, profile="truck")
            routes.update({route_index:{'route':route,'distance':route.distance}})

            points = [point[::-1] for point in route.geometry]
            vendor_loc = vendor_data_dict[vendor]['Vendor - City'] + ' ' + vendor_data_dict[vendor]['Vendor - Country Key']
            
            # html formatting for folium popups
            html=f'<span style="font-family: helvetica; font-size:9pt; color: #666666;"><strong>Vendor: <span style="color: #333333;">{vendor}</span><br /><span style="color: #666666;">Location: <span style="color: #333333;">{vendor_loc}</span><br /><span style="color: #666666;">Total distance: <span style="color: #333333;">{round(route.distance/1000)} km</strong></span>'
            iframe = folium.IFrame(html=html, width=200, height=60)
            popup = folium.Popup(iframe, max_width=2650)

            # draw markers and routes
            folium.PolyLine(points, color="red", weight=2.5, opacity=1).add_to(m)
            folium.Marker(location=[coords[0][1],coords[0][0]],popup=popup).add_to(m)
        except:
            pass
   
    folium.Marker(location=[coords[1][1],coords[1][0]],popup='Plant').add_to(m)
    
    # find shortest route and display in blue
    best_route = min(routes.keys(), key=(lambda k: routes[k]['distance']))
    points = [point[::-1] for point in routes[best_route]['route'].geometry]
    folium.PolyLine(points, color="blue", weight=2.5, opacity=1).add_to(m)

    data = io.BytesIO()
    m.save(data, close_file=False)

    w = QtWebEngineWidgets.QWebEngineView()
    w.setHtml(data.getvalue().decode())
    w.resize(800, 640)
    w.show()

    sys.exit(app.exec_())