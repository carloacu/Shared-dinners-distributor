import xml.etree.ElementTree as ET
import csv

tree = ET.parse("input.kml")
root = tree.getroot()

ns = {'kml': 'http://www.opengis.net/kml/2.2'}

points = []

for placemark in root.findall(".//kml:Placemark", ns):
    name = placemark.find("kml:name", ns).text
    coords = placemark.find(".//kml:coordinates", ns).text.strip()
    lon, lat, _ = coords.split(",")
    points.append([name, lat, lon])

with open("points.csv", "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["name", "lat", "lon"])
    writer.writerows(points)

print("Export terminé → points.csv")
