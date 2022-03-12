import folium

map = folium.Map(location=[48.1995393, 17.2985114],
                 tiles="Stamen Terrain",
                 zoom_start=15)
map.save("mapa.html")