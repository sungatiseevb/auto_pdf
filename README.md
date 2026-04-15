# auto_pdf W-9 Extractor

Утилита для извлечения данных из IRS Form W-9 в PDF.

## Что делает

Скрипт `w9_extractor_fixed.py` пытается извлечь ключевые поля W-9 из PDF-файла следующими способами:

1. Сначала пытается получить данные из PDF-виджетов (если форма интерактивная).
2. Затем пробует извлечь текст из PDF через PyMuPDF.
3. Если документ отсканирован или текстовые данные недостаточны, выполняется OCR через Tesseract.

## Поля в результате

Возвращается словарь Python с такими ключами:

- `filename` — имя PDF-файла
- `name_entity` — имя физического/юридического лица на строке 1
- `business_name` — название бизнеса / disregarded entity на строке 2
- `tob` — тип бизнеса / классификация налогоплательщика
- `address` — адрес строки 5
- `city_state_zip` — строка города / штата / ZIP
- `ssn` — SSN, если найден
- `ein` — EIN, если найден
- `date` — дата из зоны подписи / Part II
- `status` — `SUCCESS`, `PARTIAL`, `NOT_FOUND` или `ERROR`
- `method` — `widgets`, `mixed`, `text` или `ocr`
- `raw_snippet` — фрагмент текста, который использовался для извлечения

## Зависимости

Для работы требуются:

- Python 3
- `PyMuPDF` (`fitz`)
- `pytesseract`
- `pdf2image`
- `opencv-python`
- `numpy`
- Установленный Tesseract OCR в системе

Пример установки:

```bash
pip install pymupdf pytesseract pdf2image opencv-python numpy
```

Если используете Linux, установите также системный пакет Tesseract:

```bash
sudo apt install tesseract-ocr
```

## Использование

### Быстрая проверка

```bash
cd /workspaces/auto_pdf
python w9_extractor_fixed.py input_pdfs/Testpdffinal.pdf
```

### Как подключать в коде

```python
from w9_extractor_fixed import extract_w9

record = extract_w9("input_pdfs/Testpdffinal.pdf")
print(record)
```

### Тестовый запуск

В репозитории присутствует вспомогательный файл `test_w9_extractor.py` для проверки обработки каталога `input_pdfs`.

```bash
python test_w9_extractor.py
```

## Стратегия извлечения

- `widgets`: извлечение из интерактивных полей формы
- `text`: извлечение через текстовый слой PDF
- `ocr`: распознавание текста с изображения страниц, если PDF «сканированный» или текст не найден
- `mixed`: комбинация виджетов и текстовой обработки

## Примечания

- Если PDF содержит только сканированные данные, результат зависит от качества OCR.
- SSN/EIN извлекаются как `XXX-XX-XXXX` и `XX-XXXXXXX`.
- Документ поддерживает разные варианты меток W-9, но могут встречаться сложные макеты, требующие дополнительной доработки.
