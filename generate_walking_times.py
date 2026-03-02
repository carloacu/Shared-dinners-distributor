#!/usr/bin/env python3
import argparse
import csv
import json
import math
import os
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from typing import List, Tuple, Dict


KML_NS = {"kml": "http://www.opengis.net/kml/2.2"}
EARTH_RADIUS_M = 6_371_000
DIRECTIONS_URL = "https://maps.googleapis.com/maps/api/directions/json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Read a KML file with at least 2 folders (layers) and generate pairwise "
            "walking times from layer 1 to layer 2."
        )
    )
    parser.add_argument(
        "--input",
        default="input.kml",
        help="Path to input KML file (default: input.kml).",
    )
    parser.add_argument(
        "--output",
        default="walking_times.csv",
        help="Path to output CSV file (default: walking_times.csv).",
    )
    parser.add_argument(
        "--source",
        choices=["directions", "estimate"],
        default="directions",
        help=(
            "Time source: 'directions' uses Google Directions API route times; "
            "'estimate' uses straight-line approximation (default: directions)."
        ),
    )
    parser.add_argument(
        "--api-key",
        default=os.environ.get("GOOGLE_MAPS_API_KEY", ""),
        help=(
            "Google Maps API key. If omitted, uses GOOGLE_MAPS_API_KEY env var. "
            "Required when --source=directions."
        ),
    )
    parser.add_argument(
        "--travel-mode",
        default="walking",
        help="Google Directions travel mode (default: walking).",
    )
    parser.add_argument(
        "--request-delay-sec",
        type=float,
        default=0.0,
        help=(
            "Delay in seconds between API calls to reduce rate-limit pressure "
            "(default: 0)."
        ),
    )
    parser.add_argument(
        "--walk-speed-kmh",
        type=float,
        default=4.5,
        help="Walking speed in km/h used only with --source=estimate (default: 4.5).",
    )
    parser.add_argument(
        "--path-factor",
        type=float,
        default=1.3,
        help=(
            "Multiplier applied to straight-line distance with --source=estimate "
            "(default: 1.3)."
        ),
    )
    return parser.parse_args()


def haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    lat1_rad = math.radians(lat1)
    lon1_rad = math.radians(lon1)
    lat2_rad = math.radians(lat2)
    lon2_rad = math.radians(lon2)

    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad

    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2) ** 2
    )
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return EARTH_RADIUS_M * c


def parse_point(placemark: ET.Element) -> Tuple[str, float, float]:
    name_node = placemark.find("kml:name", KML_NS)
    if name_node is None or name_node.text is None:
        raise ValueError("Placemark without a name found.")

    coords_node = placemark.find(".//kml:Point/kml:coordinates", KML_NS)
    if coords_node is None or coords_node.text is None:
        raise ValueError(f"Placemark '{name_node.text}' has no point coordinates.")

    coord_text = coords_node.text.strip()
    lon_str, lat_str, *_ = [part.strip() for part in coord_text.split(",")]
    return name_node.text.strip(), float(lat_str), float(lon_str)


def load_first_two_layers(input_path: str) -> Tuple[Dict, Dict]:
    tree = ET.parse(input_path)
    root = tree.getroot()

    folders = root.findall(".//kml:Document/kml:Folder", KML_NS)
    if len(folders) < 2:
        raise ValueError(
            f"KML must contain at least 2 folders/layers. Found: {len(folders)}."
        )

    def folder_to_layer(folder: ET.Element, layer_index: int) -> Dict:
        layer_name_node = folder.find("kml:name", KML_NS)
        layer_name = (
            layer_name_node.text.strip()
            if layer_name_node is not None and layer_name_node.text
            else f"Layer {layer_index}"
        )

        points: List[Dict] = []
        for placemark in folder.findall("kml:Placemark", KML_NS):
            point_name, lat, lon = parse_point(placemark)
            points.append({"name": point_name, "lat": lat, "lon": lon})

        if not points:
            raise ValueError(f"Layer '{layer_name}' contains no point placemarks.")

        return {"name": layer_name, "points": points}

    return folder_to_layer(folders[0], 1), folder_to_layer(folders[1], 2)


def compute_rows_estimate(
    layer1: Dict, layer2: Dict, walk_speed_kmh: float, path_factor: float
) -> List[List]:
    if walk_speed_kmh <= 0:
        raise ValueError("--walk-speed-kmh must be > 0.")
    if path_factor <= 0:
        raise ValueError("--path-factor must be > 0.")

    speed_m_per_min = (walk_speed_kmh * 1000) / 60
    rows: List[List] = []

    for origin in layer1["points"]:
        for destination in layer2["points"]:
            straight_line_m = haversine_m(
                origin["lat"], origin["lon"], destination["lat"], destination["lon"]
            )
            estimated_path_m = straight_line_m * path_factor
            walking_min = estimated_path_m / speed_m_per_min

            rows.append(
                [
                    layer1["name"],
                    origin["name"],
                    layer2["name"],
                    destination["name"],
                    round(estimated_path_m),
                    round(walking_min, 1),
                    "estimate",
                ]
            )
    return rows


def fetch_directions_metrics(
    origin: Dict, destination: Dict, api_key: str, travel_mode: str
) -> Tuple[int, float]:
    params = {
        "origin": f"{origin['lat']},{origin['lon']}",
        "destination": f"{destination['lat']},{destination['lon']}",
        "mode": travel_mode,
        "key": api_key,
    }
    url = f"{DIRECTIONS_URL}?{urllib.parse.urlencode(params)}"
    request = urllib.request.Request(url, headers={"User-Agent": "kml-directions-script"})
    with urllib.request.urlopen(request, timeout=30) as response:
        payload = json.load(response)

    status = payload.get("status", "")
    if status != "OK":
        error_message = payload.get("error_message", "No error message provided.")
        raise RuntimeError(
            f"Directions API failed for '{origin['name']}' -> '{destination['name']}': "
            f"status={status}, error='{error_message}'"
        )

    routes = payload.get("routes", [])
    if not routes:
        raise RuntimeError(
            f"No routes returned for '{origin['name']}' -> '{destination['name']}'."
        )
    legs = routes[0].get("legs", [])
    if not legs:
        raise RuntimeError(
            f"No route legs returned for '{origin['name']}' -> '{destination['name']}'."
        )

    leg = legs[0]
    distance_m = leg["distance"]["value"]
    duration_min = leg["duration"]["value"] / 60
    return int(distance_m), round(duration_min, 1)


def compute_rows_directions(
    layer1: Dict,
    layer2: Dict,
    api_key: str,
    travel_mode: str,
    request_delay_sec: float,
) -> List[List]:
    if not api_key:
        raise ValueError(
            "Missing API key for Directions API. Set GOOGLE_MAPS_API_KEY or use --api-key."
        )
    if request_delay_sec < 0:
        raise ValueError("--request-delay-sec must be >= 0.")

    rows: List[List] = []
    for origin in layer1["points"]:
        for destination in layer2["points"]:
            route_distance_m, walking_min = fetch_directions_metrics(
                origin, destination, api_key, travel_mode
            )
            rows.append(
                [
                    layer1["name"],
                    origin["name"],
                    layer2["name"],
                    destination["name"],
                    route_distance_m,
                    walking_min,
                    "google_directions_api",
                ]
            )
            if request_delay_sec > 0:
                time.sleep(request_delay_sec)
    return rows


def write_csv(output_path: str, rows: List[List]) -> None:
    with open(output_path, "w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(
            [
                "from_layer",
                "from_point",
                "to_layer",
                "to_point",
                "route_distance_m",
                "walking_time_min",
                "time_source",
            ]
        )
        writer.writerows(rows)


def main() -> None:
    args = parse_args()
    layer1, layer2 = load_first_two_layers(args.input)
    if args.source == "directions":
        rows = compute_rows_directions(
            layer1,
            layer2,
            args.api_key,
            args.travel_mode,
            args.request_delay_sec,
        )
    else:
        rows = compute_rows_estimate(
            layer1,
            layer2,
            args.walk_speed_kmh,
            args.path_factor,
        )
    write_csv(args.output, rows)
    print(
        f"Wrote {len(rows)} rows to {args.output} "
        f"(from '{layer1['name']}' to '{layer2['name']}', source={args.source})."
    )


if __name__ == "__main__":
    main()
