import logging
from logging.handlers import RotatingFileHandler


# Настройка логирования
# Настраиваем логирование с двумя обработчиками: файл и консоль
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    datefmt='%d.%m.%Y %H:%M:%S',
    encoding='utf-8',
    handlers=[
        # Заменяем FileHandler на RotatingFileHandler
        RotatingFileHandler(
            'logi/logi.log',
            maxBytes=5*1024*1024,  # 5 МБ (5 * 1024 байт * 1024 байт)
            backupCount=3,         # Хранить 3 старых файла
            encoding='utf-8'
        ),
        logging.StreamHandler()    # Вывод в консоль
    ]
)
logger = logging.getLogger(__name__)