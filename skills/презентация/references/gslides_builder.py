#!/usr/bin/env python3
"""
Mindbox Presentation Builder — Google Slides API

Создаёт презентацию в Google Slides из шаблона.

Использование:
    python3 gslides_builder.py build <config.json> [--title "Название"]
    python3 gslides_builder.py check              # проверить доступ к API
    python3 gslides_builder.py inspect <slide_num> [slide_num2 ...]  # осмотреть слайды шаблона
    python3 gslides_builder.py list-slides         # показать все слайды

Формат config.json:
{
  "title": "Название презентации",
  "slides": [
    {
      "template_slide": 4,
      "replacements": {
        "Заголовок слайда": "Мой заголовок"
      }
    }
  ]
}

Результат: ссылка на готовую презентацию в Google Drive.
"""

import sys
import os
import json
from pathlib import Path

# Google API
sys.path.insert(0, str(Path.home() / '.claude' / 'google-integration' / 'scripts'))

TEMPLATE_ID = "1r3kgFFgQ_bF-vv53w5NVuskQl9hJ0AEbzIDbkg_t7ag"


def _normalize_text(text):
    """Нормализовать спецсимволы для надёжного матчинга."""
    text = text.replace('\xa0', ' ')   # неразрывный пробел → обычный
    text = text.replace('\x0b', '')    # вертикальный таб → убрать
    text = text.replace('\r', '')      # carriage return → убрать
    return text.strip()


def _get_services():
    """Получить авторизованные сервисы Google."""
    from google_auth import get_credentials
    from googleapiclient.discovery import build

    creds = get_credentials()
    slides_service = build('slides', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return slides_service, drive_service


def check_access():
    """Проверить доступ к Google Slides API."""
    try:
        slides, drive = _get_services()
        # Проверяем доступ к шаблону
        pres = slides.presentations().get(presentationId=TEMPLATE_ID).execute()
        total = len(pres.get('slides', []))
        print(f"Google Slides API: OK")
        print(f"Шаблон: {pres.get('title', 'N/A')}")
        print(f"Слайдов в шаблоне: {total}")
        return True
    except Exception as e:
        print(f"Ошибка доступа: {e}")
        return False


def _get_slide_texts(slide):
    """Извлечь тексты из слайда Google Slides."""
    texts = []
    for element in slide.get('pageElements', []):
        shape = element.get('shape', {})
        text_content = shape.get('text', {})
        for text_elem in text_content.get('textElements', []):
            text_run = text_elem.get('textRun', {})
            content = _normalize_text(text_run.get('content', ''))
            if content:
                texts.append(content)
    return texts


def _get_full_shape_text(element):
    """Получить полный текст шейпа."""
    shape = element.get('shape', {})
    text_content = shape.get('text', {})
    parts = []
    for text_elem in text_content.get('textElements', []):
        text_run = text_elem.get('textRun', {})
        content = text_run.get('content', '')
        if content:
            parts.append(content)
    return _normalize_text(''.join(parts))


def list_slides():
    """Показать все слайды шаблона."""
    slides_svc, _ = _get_services()
    pres = slides_svc.presentations().get(presentationId=TEMPLATE_ID).execute()

    print(f"Шаблон: {pres.get('title', 'N/A')}")
    print(f"Всего слайдов: {len(pres.get('slides', []))}\n")

    for idx, slide in enumerate(pres.get('slides', [])):
        texts = _get_slide_texts(slide)
        elements = len(slide.get('pageElements', []))
        print(f"Слайд {idx + 1}: {elements} elements")
        for t in texts[:5]:
            print(f"  → {t[:100]}")
        print()


def _inspect_slide_data(slide, slide_num):
    """Вывести информацию о слайде (без загрузки шаблона)."""
    print(f"=== Слайд {slide_num} ===")
    print(f"Slide ID: {slide.get('objectId')}")
    print(f"Elements: {len(slide.get('pageElements', []))}\n")

    for i, element in enumerate(slide.get('pageElements', [])):
        obj_id = element.get('objectId', '')

        el_type = 'UNKNOWN'
        if 'shape' in element:
            el_type = 'SHAPE'
        elif 'image' in element:
            el_type = 'IMAGE'
        elif 'table' in element:
            el_type = 'TABLE'
        elif 'group' in element:
            el_type = 'GROUP'

        print(f"Element {i}: {el_type} (id={obj_id})")

        if 'shape' in element:
            full_text = _get_full_shape_text(element)
            if full_text:
                print(f"  Text: {full_text[:200]}")

                # Font info
                for text_elem in element['shape'].get('text', {}).get('textElements', []):
                    text_run = text_elem.get('textRun', {})
                    style = text_run.get('style', {})
                    content = _normalize_text(text_run.get('content', ''))
                    if content and style:
                        font_size = style.get('fontSize', {}).get('magnitude', '')
                        font_unit = style.get('fontSize', {}).get('unit', '')
                        bold = style.get('bold', False)
                        info = []
                        if font_size:
                            info.append(f"size={font_size}{font_unit}")
                        if bold:
                            info.append("bold")
                        if info:
                            print(f"    Run: '{content[:80]}' [{', '.join(info)}]")

        if 'image' in element:
            print(f"  [IMAGE]")

        print()


def inspect_slides(slide_nums):
    """Осмотреть несколько слайдов за один запрос к API."""
    slides_svc, _ = _get_services()
    pres = slides_svc.presentations().get(presentationId=TEMPLATE_ID).execute()
    all_slides = pres.get('slides', [])

    for slide_num in slide_nums:
        if slide_num < 1 or slide_num > len(all_slides):
            print(f"Слайд {slide_num} не найден (всего {len(all_slides)})\n")
            continue
        _inspect_slide_data(all_slides[slide_num - 1], slide_num)


def build_presentation(config, title_override=None):
    """
    Собрать презентацию в Google Slides.

    1. Копировать шаблон
    2. Удалить ненужные слайды
    3. Переупорядочить оставшиеся
    4. Заменить текст
    5. Вернуть ссылку
    """
    slides_svc, drive_svc = _get_services()

    slides_config = config.get("slides", [])
    if not slides_config:
        raise ValueError("Конфиг не содержит слайдов")

    title = title_override or config.get("title", "Новая презентация Mindbox")

    # Шаг 1: Копируем шаблон
    print(f"Копирую шаблон...")
    copy_result = drive_svc.files().copy(
        fileId=TEMPLATE_ID,
        body={'name': title}
    ).execute()
    new_id = copy_result['id']
    print(f"Создана копия: {new_id}")

    # Получаем структуру копии
    pres = slides_svc.presentations().get(presentationId=new_id).execute()
    all_slides = pres.get('slides', [])
    total = len(all_slides)

    # Маппинг: slide_index (0-based) -> slide_object_id
    slide_ids = [s['objectId'] for s in all_slides]

    # Определяем какие слайды нужны (0-based indices)
    needed_indices = [sc["template_slide"] - 1 for sc in slides_config]

    # Шаг 2: Дублируем слайды если нужны повторы
    # Сначала создаём дубликаты для повторных использований
    from collections import Counter
    index_counts = Counter(needed_indices)
    # Для каждого индекса, который нужен >1 раз, создаём дубликаты
    requests = []
    for idx, count in sorted(index_counts.items()):
        if count > 1:
            for extra in range(count - 1):
                requests.append({
                    'duplicateObject': {
                        'objectId': slide_ids[idx]
                    }
                })

    new_slide_ids_from_dupes = []
    if requests:
        print(f"Дублирую {len(requests)} слайдов...")
        result = slides_svc.presentations().batchUpdate(
            presentationId=new_id,
            body={'requests': requests}
        ).execute()

        for reply in result.get('replies', []):
            dup_reply = reply.get('duplicateObject', {})
            new_obj_id = dup_reply.get('objectId', '')
            if new_obj_id:
                new_slide_ids_from_dupes.append(new_obj_id)

    # Строим финальный маппинг: позиция в выходе -> objectId слайда
    seen_count = {}
    final_slide_ids = []

    for idx in needed_indices:
        count = seen_count.get(idx, 0)
        if count == 0:
            final_slide_ids.append(slide_ids[idx])
        else:
            # Берём дублированный слайд
            # Нужно найти какой дубликат соответствует этому extra
            dupe_offset = sum(
                min(seen_count.get(i, 0), index_counts[i] - 1)
                for i in sorted(index_counts.keys())
                if i < idx and index_counts[i] > 1
            ) + count - 1
            if dupe_offset < len(new_slide_ids_from_dupes):
                final_slide_ids.append(new_slide_ids_from_dupes[dupe_offset])
        seen_count[idx] = count + 1

    # Шаг 3: Удаляем ненужные слайды
    # Обновляем список слайдов (после дублирования могут быть новые)
    pres = slides_svc.presentations().get(presentationId=new_id).execute()
    current_slide_ids = [s['objectId'] for s in pres.get('slides', [])]

    to_delete = [sid for sid in current_slide_ids if sid not in final_slide_ids]

    if to_delete:
        print(f"Удаляю {len(to_delete)} лишних слайдов...")
        delete_requests = [
            {'deleteObject': {'objectId': sid}}
            for sid in to_delete
        ]
        slides_svc.presentations().batchUpdate(
            presentationId=new_id,
            body={'requests': delete_requests}
        ).execute()

    # Шаг 4: Переупорядочиваем
    # После удаления, нужно расставить слайды в правильном порядке
    pres = slides_svc.presentations().get(presentationId=new_id).execute()
    current_ids_after_delete = [s['objectId'] for s in pres.get('slides', [])]

    reorder_requests = []
    for target_pos, desired_id in enumerate(final_slide_ids):
        if desired_id in current_ids_after_delete:
            reorder_requests.append({
                'updateSlidesPosition': {
                    'slideObjectIds': [desired_id],
                    'insertionIndex': target_pos
                }
            })

    if reorder_requests:
        print(f"Переупорядочиваю слайды...")
        slides_svc.presentations().batchUpdate(
            presentationId=new_id,
            body={'requests': reorder_requests}
        ).execute()

    # Шаг 5: Заменяем текст и чистим шаблонные артефакты
    pres = slides_svc.presentations().get(presentationId=new_id).execute()
    final_slides = pres.get('slides', [])

    replace_requests = []
    style_requests = []
    placeholder_creates = []  # данные для текстовых плейсхолдеров вместо изображений
    replacements_count = 0
    cleared_count = 0
    deleted_images = 0

    if len(final_slides) != len(slides_config):
        print(f"⚠️ НЕСОВПАДЕНИЕ: конфиг = {len(slides_config)} слайдов, в презентации = {len(final_slides)}")
        print(f"   final_slide_ids = {final_slide_ids}")
        print(f"   actual_ids = {[s['objectId'] for s in final_slides]}")

    for i, sc in enumerate(slides_config):
        replacements = sc.get("replacements", {})
        if i >= len(final_slides):
            print(f"⚠️ Слайд {i+1} (template:{sc.get('template_slide')}) — ПРОПУЩЕН, нет в презентации!")
            continue

        slide = final_slides[i]
        # 5a. Обрабатываем текстовые шейпы
        for element in slide.get('pageElements', []):
            shape = element.get('shape', {})
            text_content = shape.get('text', {})
            if not text_content:
                continue

            full_text = _get_full_shape_text(element)
            if not full_text:
                continue

            obj_id = element['objectId']

            # Считаем длину текста
            total_length = 0
            for text_elem in text_content.get('textElements', []):
                text_run = text_elem.get('textRun', {})
                content = text_run.get('content', '')
                total_length += len(content)
            delete_length = total_length - 1 if total_length > 0 else 0

            if delete_length <= 0:
                continue

            # Сохраняем стиль первого текстового run
            saved_style = None
            for text_elem in text_content.get('textElements', []):
                text_run = text_elem.get('textRun', {})
                if text_run.get('content', '').strip():
                    saved_style = text_run.get('style', {})
                    break

            # Ищем замену для этого шейпа (сначала точное совпадение, потом подстрока)
            matched_key = None
            if replacements:
                for old_text in replacements:
                    old_stripped = _normalize_text(old_text)
                    if full_text == old_stripped:
                        matched_key = old_text
                        break
                if matched_key is None:
                    for old_text in replacements:
                        old_stripped = _normalize_text(old_text)
                        if old_stripped in full_text:
                            matched_key = old_text
                            break

            if matched_key is not None:
                # Есть замена — удаляем старый текст, вставляем новый
                new_text = replacements[matched_key]
                # Пустая строка вызывает ошибку API (startIndex 0 < endIndex 0) — заменяем на пробел
                if new_text == "":
                    new_text = " "
                replace_requests.append({
                    'deleteText': {
                        'objectId': obj_id,
                        'textRange': {
                            'type': 'FIXED_RANGE',
                            'startIndex': 0,
                            'endIndex': delete_length
                        }
                    }
                })
                replace_requests.append({
                    'insertText': {
                        'objectId': obj_id,
                        'insertionIndex': 0,
                        'text': new_text
                    }
                })
                # Восстанавливаем стиль шрифта
                if saved_style:
                    text_style = {}
                    fields = []
                    if 'fontFamily' in saved_style:
                        text_style['fontFamily'] = saved_style['fontFamily']
                        fields.append('fontFamily')
                    if 'fontSize' in saved_style:
                        text_style['fontSize'] = saved_style['fontSize']
                        fields.append('fontSize')
                    if 'foregroundColor' in saved_style:
                        text_style['foregroundColor'] = saved_style['foregroundColor']
                        fields.append('foregroundColor')
                    if 'bold' in saved_style:
                        text_style['bold'] = saved_style['bold']
                        fields.append('bold')
                    if text_style:
                        style_requests.append({
                            'updateTextStyle': {
                                'objectId': obj_id,
                                'textRange': {
                                    'type': 'FIXED_RANGE',
                                    'startIndex': 0,
                                    'endIndex': len(new_text)
                                },
                                'style': text_style,
                                'fields': ','.join(fields)
                            }
                        })
                replacements_count += 1
            else:
                # Нет замены — очищаем шаблонный текст-заглушку
                replace_requests.append({
                    'deleteText': {
                        'objectId': obj_id,
                        'textRange': {
                            'type': 'FIXED_RANGE',
                            'startIndex': 0,
                            'endIndex': delete_length
                        }
                    }
                })
                cleared_count += 1

        # 5b. ВСЕГДА удаляем шаблонные изображения (они нерелевантны)
        # Если есть image_placeholder — создаём текстовый плейсхолдер на месте
        image_placeholder = sc.get("image_placeholder")
        first_image_pos = None
        for element in slide.get('pageElements', []):
            if 'image' not in element:
                continue
            img_id = element['objectId']
            replace_requests.append({
                'deleteObject': {'objectId': img_id}
            })
            deleted_images += 1
            # Запоминаем позицию первого изображения для плейсхолдера
            if image_placeholder and first_image_pos is None:
                first_image_pos = {
                    'page_id': slide.get('objectId'),
                    'size': element.get('size', {}),
                    'transform': element.get('transform', {}),
                    'text': f"[{image_placeholder}]"
                }

        if first_image_pos:
            placeholder_creates.append(first_image_pos)

    if replace_requests:
        print(f"Заменяю текст ({replacements_count} шейпов, очищаю {cleared_count} шаблонных, удаляю {deleted_images} изображений)...")
        slides_svc.presentations().batchUpdate(
            presentationId=new_id,
            body={'requests': replace_requests}
        ).execute()

    # Применяем стили отдельным батчем (после вставки текста)
    if style_requests:
        print(f"Восстанавливаю шрифты ({len(style_requests)} шейпов)...")
        slides_svc.presentations().batchUpdate(
            presentationId=new_id,
            body={'requests': style_requests}
        ).execute()

    # Шаг 6: Создаём текстовые плейсхолдеры на месте удалённых изображений
    if placeholder_creates:
        ph_requests = []
        for idx, ph in enumerate(placeholder_creates):
            ph_id = f"img_placeholder_{idx}"
            element_props = {'pageObjectId': ph['page_id']}
            if ph['size']:
                element_props['size'] = ph['size']
            if ph['transform']:
                element_props['transform'] = ph['transform']

            ph_requests.append({
                'createShape': {
                    'objectId': ph_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': element_props
                }
            })
            ph_requests.append({
                'insertText': {
                    'objectId': ph_id,
                    'insertionIndex': 0,
                    'text': ph['text']
                }
            })

        print(f"Создаю {len(placeholder_creates)} плейсхолдеров для изображений...")
        slides_svc.presentations().batchUpdate(
            presentationId=new_id,
            body={'requests': ph_requests}
        ).execute()

    # Результат
    url = f"https://docs.google.com/presentation/d/{new_id}/edit"
    print(f"\nГотово!")
    print(f"Презентация: {url}")
    print(f"Слайдов: {len(slides_config)}")
    return url, new_id


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    command = sys.argv[1]

    if command == "check":
        check_access()

    elif command == "build":
        if len(sys.argv) < 3:
            print("Использование: gslides_builder.py build <config.json> [--title 'Название']")
            sys.exit(1)

        with open(sys.argv[2]) as f:
            config = json.load(f)

        title = None
        if '--title' in sys.argv:
            idx = sys.argv.index('--title')
            if idx + 1 < len(sys.argv):
                title = sys.argv[idx + 1]

        build_presentation(config, title)

    elif command == "list-slides":
        list_slides()

    elif command == "inspect":
        if len(sys.argv) < 3:
            print("Использование: gslides_builder.py inspect <slide_num> [slide_num2 ...]")
            sys.exit(1)
        inspect_slides([int(num) for num in sys.argv[2:]])

    else:
        print(f"Неизвестная команда: {command}")
        sys.exit(1)


if __name__ == "__main__":
    main()
