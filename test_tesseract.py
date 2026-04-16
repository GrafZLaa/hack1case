import pytesseract
from PIL import Image
import cv2
import numpy as np

# Указываем путь явно
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

print("🔍 Проверка Tesseract...")
print(f"Версия: {pytesseract.get_tesseract_version()}")

# Проверка доступных языков
try:
    langs = pytesseract.get_languages(config='')
    print(f"Доступные языки: {langs}")
    if 'rus' not in langs:
        print("⚠️  ВНИМАНИЕ: Русский язык (rus) НЕ найден!")
        print("📥 Скачайте rus.traineddata: https://github.com/tesseract-ocr/tessdata/raw/main/rus.traineddata")
        print("📁 Скопируйте в: C:\\Program Files\\Tesseract-OCR\\tessdata\\")
except Exception as e:
    print(f"Ошибка при получении языков: {e}")

# Тест на простом изображении
print("\n🧪 Тест распознавания...")
test_text = pytesseract.image_to_string(
    Image.new('RGB', (200, 50), color='white'),
    lang='rus+eng',
    config='--psm 6'
)
print(f"Тестовый результат: '{test_text}'")