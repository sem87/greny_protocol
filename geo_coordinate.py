import requests
import time
import math
from logi.logi import logger


def format_to_dms(lat, lon):
    """Переводит десятичные координаты в формат DMS и возвращает кортеж (lat_dms, lon_dms)"""

    def convert(dec, is_lat):
        dec = abs(dec)
        d = int(dec)
        m = int((dec - d) * 60)
        s = (dec - d - m / 60) * 3600
        direction = ('N' if lat >= 0 else 'S') if is_lat else ('E' if lon >= 0 else 'W')
        return f"{d}°{m}'{s:.2f}\"{direction}"

    lat_dms = convert(lat, True)
    lon_dms = convert(lon, False)

    return (lat_dms, lon_dms)

def haversine(lat1, lon1, lat2, lon2):
    """Расстояние между двумя точками в метрах"""
    R = 6371000
    phi1, phi2 = map(lambda x: x * math.pi / 180, (lat1, lat2))
    dphi = (lat2 - lat1) * math.pi / 180
    dlambda = (lon2 - lon1) * math.pi / 180
    a = (math.sin(dphi / 2) ** 2 +
         math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2)
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def geocode_nominatim_direct(address):
    """Прямой запрос к Nominatim"""
    # print(f"  [DEBUG] Геокодинг: '{address}'...")
    url = "https://nominatim.openstreetmap.org/search"
    params = {
        'q': address,
        'format': 'json',
        'limit': 1,
        'accept-language': 'ru'
    }
    headers = {'User-Agent': 'RoadParserApp/1.0 (educational)'}
    try:
        response = requests.get(url, params=params, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
        if not data:
            return None, None, None
        result = data[0]
        return float(result['lat']), float(result['lon']), result.get('display_name', '')
    except Exception as e:
        print(f"  [DEBUG] Ошибка геокодинга: {e}")
        return None, None, None


def try_overpass_server(server_url, query, timeout=45, retries=2):
    """Пытается сделать запрос к одному серверу с повторными попытками"""
    for attempt in range(retries):
        try:
            if attempt > 0:
                print(f"    [RETRY] Попытка {attempt + 1}/{retries} для {server_url}")
                time.sleep(2)  # Пауза перед повтором
            response = requests.get(
                server_url,
                params={'data': query},
                timeout=timeout,
                headers={'User-Agent': 'RoadParserApp/1.0'}
            )
            if response.status_code == 200:
                data = response.json()
                if 'elements' in data and len(data['elements']) > 0:
                    return data, len(data['elements'])
                elif 'elements' in data:
                    print(f"    [INFO] Сервер ответил, но элементов: 0")
                    return None, 0
            else:
                print(f"    [WARN] Статус {response.status_code}")

        except requests.exceptions.Timeout:
            print(f"    [TIMEOUT] Сервер не ответил за {timeout}с")
        except requests.exceptions.ConnectionError as e:
            print(f"    [CONN] Ошибка соединения: {str(e)[:80]}...")
        except Exception as e:
            print(f"    [ERR] {type(e).__name__}: {str(e)[:80]}...")

    return None, 0


def query_overpass_all_servers(query, lat, lon):
    """Перебирает ВСЕ возможные Overpass-серверы"""

    # 🔴 МАКСИМАЛЬНЫЙ список серверов (HTTP и HTTPS)
    servers = [
        # HTTPS серверы
        "https://overpass-api.de/api/interpreter",
        "https://overpass.kumi.systems/api/interpreter",
        "https://overpass.openstreetmap.fr/api/interpreter",
        "https://overpass.osm.ch/api/interpreter",
        "https://overpass.private.coffee/api/interpreter",
        "https://overpass.brevy.com/api/interpreter",
        "https://overpass.openstreetmap.ru/api/interpreter",
        "https://overpass.openstreetmap.hu/api/interpreter",
        "https://overpass.api.openstreetmap.fr/api/interpreter",

        # HTTP серверы (иногда работают лучше в РФ)
        "http://overpass-api.de/api/interpreter",
        "http://overpass.kumi.systems/api/interpreter",
        "http://overpass.openstreetmap.ru/api/interpreter",
        "http://overpass.osm.rambler.ru/cgi/interpreter",
        "http://maps.mail.ru/osm/tools/overpass/api/interpreter",
    ]

    # print(f"\n  [INFO] Перебираем {len(servers)} Overpass-серверов...")
    for i, server in enumerate(servers, 1):
        print(f"\n  [{i}/{len(servers)}] {server}")
        data, count = try_overpass_server(server, query, timeout=45, retries=2)
        if data is not None and count > 0:
            print(f"  ✅ УСПЕХ! Сервер: {server}")
            # print(f"  ✅ Получено элементов: {count}")
            return data

        if count == 0 and data is not None:
            print(f"  ⚠️  Сервер ответил, но дорог не найдено")
            # Продолжаем перебор, вдруг другой сервер найдет

    return None


def get_road_points(lat, lon, n_meters=50, points_count=6):
    """Находит ближайшую дорогу и генерирует точки через каждые n_meters"""
    query = f"""
    [out:json][timeout:25];
    (
      way["highway"~"primary|secondary|tertiary|residential|unclassified"](around:500,{lat},{lon});
    );
    out body;
    >;
    out skel qt;
    """
    data = query_overpass_all_servers(query, lat, lon)
    if data is None or not data.get('elements'):
        raise ValueError(
            "❌ НИ ОДИН Overpass-сервер не ответил или не нашел дорог.\n"
            "   Возможные причины:\n"
            "   1. Проблемы с интернетом/прокси\n"
            "   2. Все серверы временно недоступны\n"
            "   3. В радиусе 500м нет дорог\n"
            "   Попробуйте позже или увеличьте радиус поиска."
        )
    # Собираем узлы и геометрию дорог
    nodes = {elem['id']: (elem['lat'], elem['lon'])
             for elem in data['elements'] if elem['type'] == 'node'}
    ways = []
    for elem in data['elements']:
        if elem['type'] == 'way':
            way_nodes = [nodes[nid] for nid in elem['nodes'] if nid in nodes]
            if len(way_nodes) > 1:
                ways.append(way_nodes)
    # print(f"  [DEBUG] Восстановлено дорог: {len(ways)}")
    if not ways:
        raise ValueError("Не удалось восстановить геометрию дорог.")

    # Находим ближайший узел дороги
    best_way = None
    best_idx = -1
    min_dist = float('inf')

    for way in ways:
        for i, node in enumerate(way):
            dist = haversine(lat, lon, node[0], node[1])
            if dist < min_dist:
                min_dist = dist
                best_way = way
                best_idx = i

    # print(f"  [DEBUG] Ближайший узел: {min_dist:.1f}м от центра")
    # Прогулка вдоль дороги
    def walk_along_way(way, start_idx, direction):
        points = [way[start_idx]]
        current_point = way[start_idx]
        remaining_dist = float(n_meters)
        max_iterations = 10000
        iterations = 0

        indices = range(start_idx + 1, len(way)) if direction > 0 else range(start_idx - 1, -1, -1)

        for i in indices:
            iterations += 1
            if iterations > max_iterations:
                break

            next_node = way[i]
            seg_dist = haversine(current_point[0], current_point[1], next_node[0], next_node[1])

            if seg_dist < 0.001:
                current_point = next_node
                continue

            original_seg_dist = seg_dist

            while remaining_dist <= seg_dist and len(points) < points_count:
                ratio = remaining_dist / seg_dist
                new_lat = current_point[0] + ratio * (next_node[0] - current_point[0])
                new_lon = current_point[1] + ratio * (next_node[1] - current_point[1])

                current_point = (new_lat, new_lon)
                points.append(current_point)

                seg_dist -= remaining_dist
                remaining_dist = float(n_meters)

            passed_in_segment = original_seg_dist - seg_dist
            remaining_dist -= passed_in_segment
            current_point = next_node

        return points

    result_points = walk_along_way(best_way, best_idx, 1)

    if len(result_points) < points_count:
        result_points = walk_along_way(best_way, best_idx, -1)

    return result_points[:points_count]

def main_geo_coordinate(location_measure_metrics):
    # location_measure_metrics = "Шаймуратово"
    itog_location_measure_metrics = location_measure_metrics.strip().capitalize() + " Республика Башкортостан"
    # Шаг 1: Геокодинг
    lat, lon, display_name = geocode_nominatim_direct(itog_location_measure_metrics)
    # Шаг 2: Поиск дороги
    try:
        road_points = get_road_points(lat, lon, n_meters=50, points_count=6)
        # print(f"&&&&&&&&&&&&&&&&&&&&&&7{road_points}")
        return road_points
    except Exception as e:
        print(f"\n❌ ОШИБКА: {type(e).__name__}: {e}")
# ==================== ОСНОВНОЙ СКРИПТ ====================

if __name__ == "__main__":
    main_geo_coordinate(location_measure_metrics = "Шаймуратово")