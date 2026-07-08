import math
import re
import random
def parse_point(point_str):
    """
    Парсит строку с координатами формата DMS (градусы, минуты, секунды).
    Возвращает широту и долготу в десятичном виде.
    """
    # Регулярное выражение ищет паттерн: Число° Число' Число" Направление
    # Поддержка "None" на случай невалидных данных
    pattern = r"(\d+)\s*°\s*([\d\.]+|None)\s*['′]\s*([\d\.]+)\s*[\"″]?\s*([NSEW])"
    matches = re.findall(pattern, point_str, re.IGNORECASE)
    if len(matches) < 2:
        raise ValueError(f"Не удалось найти две координаты в строке: {point_str}")
    coords = []
    for match in matches:
        deg = float(match[0])
        # Обработка "None" (считаем как 0.0)
        min_val = 0.0 if match[1].lower() == 'none' else float(match[1])
        sec = float(match[2])
        direction = match[3].upper()
        # Перевод в десятичные градусы
        decimal_deg = deg + min_val / 60 + sec / 3600
        # Юг и Запад имеют отрицательный знак
        if direction in ['S', 'W']:
            decimal_deg = -decimal_deg
        coords.append(decimal_deg)
    return coords[0], coords[1]  # lat, lon
def calculate_azimuth(point1_str, point2_str):
    """
    Вычисляет расчетный азимут (дирекционный угол / пеленг)
    от первой точки ко второй в диапазоне от 0 до 360 градусов.
    """
    lat1, lon1 = "55°50'56 N", "56°50'34 E"
    lat2, lon2 = "54°22'0.38N","56°2'38.41E"
    lat1, lon1 = parse_point(point1_str)
    lat2, lon2 = parse_point(point2_str)
    # Переводим градусы в радианы
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    lambda1 = math.radians(lon1)
    lambda2 = math.radians(lon2)
    d_lambda = lambda2 - lambda1
    # Формула азимута (Initial bearing)
    x = math.cos(phi2) * math.sin(d_lambda)
    y = math.cos(phi1) * math.sin(phi2) - math.sin(phi1) * math.cos(phi2) * math.cos(d_lambda)
    theta = math.atan2(x, y)
    azimuth = math.degrees(theta)
    # Нормализуем значение в диапазон [0, 360)
    return (azimuth + 360) % 360

def main_raschetnii_azimut(p1,p2):
    # Для наглядности покажем, как распарсились координаты
    lat1, lon1 = parse_point(p1)
    lat2, lon2 = parse_point(p2)
    # print(f"Точка 1 (д): {lat1:.6f}, {lon1:.6f}")
    # print(f"Точка 2 (десятичные): {lat2:.6f}, {lon2:.6f}\n")
    # Считаем азимут
    az = round(calculate_azimuth(p1, p2),2)
    # print(f"Расчетный азимут от Точки 1 к Точке 2: {az:.1f}°")
    return az


def simulate_measured_azimuth(true_azimuth, sigma_deg=0.05):
    """
    Имитирует измерение азимута прибором с заданной точностью.

    :param true_azimuth: истинный (расчетный) азимут в градусах
    :param sigma_deg: среднеквадратическая погрешность прибора (в градусах)
    :return: измеренный азимут в диапазоне [0, 360)
    """
    # Гауссовская погрешность
    error = random.gauss(0, sigma_deg)
    measured = true_azimuth + error
    return measured % 360



if __name__ == "__main__":
    p1 = "55° 50' 56\" N, 56° 1'34\" E"
    p2 = "54°22'0.38\"N,56°2'38.41\"E"
    main_raschetnii_azimut(p1=p1,p2=p2)