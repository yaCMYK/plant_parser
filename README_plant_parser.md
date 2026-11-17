# Production parser базы растений

## Что делает

Парсер обходит страницы одного домена, выделяет вероятные карточки растений, извлекает пары «признак → значение», сопоставляет их с целевыми колонками по строгим смысловым правилам и экспортирует XLSX.

Главный принцип: **raw никогда не удаляется**. Нормализованное значение используется в основном листе, а исходное значение, min/max, исходная единица, canonical unit, URL, evidence и confidence сохраняются на листе `Raw evidence`.

## XLSX

Файл содержит:

- `База растений` — одна строка = одно уникальное растение; в одной ячейке может быть несколько значений, разделённых переносом строки.
- `Raw evidence` — длинная таблица всех принятых и отклонённых observations.
- `Правила полей` — строгие правила для каждой целевой колонки.
- `Crawl log` — журнал обхода.
- `Отчёт парсинга` — метрики заполнения.

Все листы оформляются в чёрно-белом стиле.

## Единицы

- длина: canonical = `m`;
- объём контейнера: canonical = `L`;
- pH: безразмерный;
- температура: canonical = `°C`;
- валюта: сохраняется исходная, **без автоматической конвертации**;
- диапазоны сохраняются как `min` + `max`;
- если на странице несколько размеров, сохраняются все observations.

Пример: `150–200 cm` → normalized `1.5–2 m`, min `1.5`, max `2.0`, source unit `cm`; raw `150–200 cm` остаётся в `Raw evidence`.

## Строгие различия

- `wet/moist soil` **не** означает устойчивость к переувлажнению;
- `waterlogging tolerant / переносит переувлажнение` означает;
- `hardy` не превращается автоматически в USDA;
- `USDA zone 4–6` хранится как USDA;
- `-30 °C` хранится отдельно как температура;
- размер торгового контейнера не означает пригодность к выращиванию в кашпо;
- взрослая высота не смешивается с высотой саженца;
- `price` попадает в стоимость материала только при отсутствии контекста посадки/ухода/гарантии.

## Установка

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

Для JS-сайта:

```bash
pip install playwright
playwright install chromium
```

## Запуск

```bash
python plant_parser_production.py \
  --url "https://example.com/catalog/plants" \
  --output "plants.xlsx" \
  --max-pages 500
```

Если canonical длины нужно показать в основной таблице в сантиметрах, не меняя хранилище raw/min/max:

```bash
python plant_parser_production.py \
  --url "https://example.com/catalog/plants" \
  --output "plants.xlsx" \
  --display-length-unit cm
```

Если URL карточек имеют узнаваемый путь, лучше сузить обход:

```bash
python plant_parser_production.py \
  --url "https://example.com/catalog/plants" \
  --output "plants.xlsx" \
  --include-regex "/(plants|catalog)/" \
  --exclude-regex "/(cart|login|news|blog)/"
```

## Сайт-специфичные синонимы

Не меняйте ядро. Добавляйте в `rules_overrides.json` только реально обнаруженные на сайте подписи:

```json
{
  "fields": {
    "mature_height": {"add_aliases": ["Высота через 10 лет"]}
  }
}
```

Запуск:

```bash
python plant_parser_production.py --url "..." --rules-overrides rules_overrides.json
```

## Важное ограничение

Это универсальное ядро. После первого запуска на конкретном сайте нужно посмотреть `Raw evidence`, `Crawl log` и 10–20 исходных карточек. Если сайт использует собственные термины или необычную HTML-разметку, добавляются site-specific aliases / selectors, но правила смысла и нормализации остаются общими.


## Проверка ядра

В комплект включён `test_parser_core.py`. Запуск: `python -m unittest -v test_parser_core.py`. Тесты проверяют диапазоны, min/max, one-sided значения, контейнеры, pH, USDA, засуху/обводнение, growth rate и преобразование единиц.
