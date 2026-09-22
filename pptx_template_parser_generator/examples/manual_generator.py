from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

from PIL import ImageFont
import os
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


# ============================================================
# ШАБЛОН VK TECH 1
# ============================================================

TEMPLATE = {
    "layout_id": "vktech_template1_title",
    "type": "title",

    # --------------------------------------------------------
    # РАЗМЕР СЛАЙДА
    # --------------------------------------------------------

    "slide_size": {
        "width": 25.4,
        "height": 14.288,
        "unit": "cm",
    },

    # --------------------------------------------------------
    # ЦВЕТА
    # --------------------------------------------------------

    "colors": {
        "text_primary": "#FAFCFF",
        "background_primary": "#0A0F1F",
    },

    # --------------------------------------------------------
    # ASSETS
    # --------------------------------------------------------

    "assets": {
        "background_image": str(PROJECT_ROOT / "assets/vktech/template1/title_background.png"),
    },

    # --------------------------------------------------------
    # ОБЩИЕ НАСТРОЙКИ ТЕКСТОВЫХ БЛОКОВ
    # --------------------------------------------------------

    "text_box_defaults": {
        "margins": {
            "left": 0.25,
            "right": 0.25,
            "top": 0.13,
            "bottom": 0.13,
        },
    },

    # --------------------------------------------------------
    # ШРИФТЫ
    # --------------------------------------------------------

    "fonts": {
        "title": {
            "family": "Play",
            "size": 48,
            "weight": "regular",
        },

        "subtitle": {
            "family": "Play",
            "size": 16,
            "weight": "regular",
        },
    },

    # --------------------------------------------------------
    # ВАРИАНТЫ
    # --------------------------------------------------------

    "variants": [
        {
            "id": "title_subtitle_center",

            "slots": {
                "title": {
                    "x": 6.87,
                    "y": 3.81,
                    "w": 11.91,
                    "h": 2.22,
                    "vertical_align": "bottom",
                },

                "subtitle": {
                    "x": 6.87,
                    "y": 6.11,
                    "w": 11.91,
                    "h": 0.79,
                    "vertical_align": "top",
                },

                "logo": {
                    "src": str(PROJECT_ROOT / "assets/vktech/logo.png"),
                    "x": 10.96,
                    "y": 1.86,
                    "w": 3.4,
                    "h": 1.07,
                },
            },

            "constraints": {
                "title_max_lines": 1,
                "subtitle_max_lines": 1,
            },
        },

        {
            "id": "title_subtitle_left",

            "slots": {
                "title": {
                    "x": 1.62,
                    "y": 4.25,
                    "w": 11.91,
                    "h": 2.22,
                    "vertical_align": "middle",
                },

                "subtitle": {
                    "x": 1.62,
                    "y": 6.56,
                    "w": 11.91,
                    "h": 0.79,
                    "vertical_align": "middle",
                },

                "text": {
                    "x": 3.97,
                    "y": 10.67,
                    "w": 11.91,
                    "h": 0.79,
                    "vertical_align": "middle",
                },

                "text_2": {
                    "x": 3.97,
                    "y": 11.69,
                    "w": 11.91,
                    "h": 0.79,
                    "vertical_align": "middle",
                },

                "logo": {
                    "src": str(PROJECT_ROOT / "assets/vktech/logo.png"),
                    "x": 1.86,
                    "y": 1.89,
                    "w": 2.53,
                    "h": 0.8,
                },
            },

            "constraints": {
                "title_max_lines": 1,
                "subtitle_max_lines": 1,
            },
        },

        {
            "id": "title_subtitle_left_2",

            "slots": {
                "title": {
                    "x": 1.59,
                    "y": 4.62,
                    "w": 16.00,
                    "h": 3.87,
                    "vertical_align": "middle",
                },

                "subtitle": {
                    "x": 3.97,
                    "y": 10.67,
                    "w": 11.91,
                    "h": 0.79,
                    "vertical_align": "middle",
                },

                "text": {
                    "x": 3.97,
                    "y": 11.69,
                    "w": 11.91,
                    "h": 0.79,
                    "vertical_align": "middle",
                },

                "logo": {
                    "src": str(PROJECT_ROOT / "assets/vktech/logo.png"),
                    "x": 1.86,
                    "y": 1.89,
                    "w": 2.53,
                    "h": 0.8,
                },
            },

            "constraints": {
                "title_max_lines": 2,
                "subtitle_max_lines": 1,
            },
        },
    ],

    # --------------------------------------------------------
    # ОГРАНИЧЕНИЯ
    # --------------------------------------------------------

    "allowed_content": [
        "title",
        "subtitle",
    ],

    "disallowed_content": [
        "list",
        "table",
        "chart",
        "image",
    ],
}


# ============================================================
# РАЗМЕР СЛАЙДА
# ============================================================

SLIDE_WIDTH_CM = TEMPLATE["slide_size"]["width"]
SLIDE_HEIGHT_CM = TEMPLATE["slide_size"]["height"]


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def hex_to_rgb(hex_color):
    """
    '#FAFCFF' -> RGBColor(250, 252, 255)
    """

    hex_color = hex_color.lstrip("#")

    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
    )


def find_font():
    """
    Ищет шрифт Play в Windows.
    """

    fonts_folder = r"C:\Windows\Fonts"

    possible_names = [
        "Play-Regular.ttf",
        "play-regular.ttf",
        "Play.ttf",
        "Arial.ttf",
    ]

    for name in possible_names:

        path = os.path.join(
            fonts_folder,
            name,
        )

        if os.path.exists(path):
            return path

    raise FileNotFoundError(
        "Не найден шрифт Play.\n"
        "Проверь, что шрифт Play установлен в Windows."
    )


def cm_to_px(cm):
    """
    Перевод сантиметров в пиксели.
    """

    return cm * 96 / 2.54


def text_width(text, font):
    """
    Возвращает ширину текста в пикселях.
    """

    bbox = font.getbbox(text)

    return bbox[2] - bbox[0]


def text_fits(text, slot, font_config):
    """
    Проверяет, помещается ли текст
    в ширину своего slot.

    Размер шрифта НЕ изменяется.
    """

    font_path = find_font()

    font = ImageFont.truetype(
        font_path,
        font_config["size"],
    )

    margins = TEMPLATE["text_box_defaults"]["margins"]

    max_width = cm_to_px(
        slot["w"]
        - margins["left"]
        - margins["right"]
    )

    lines = text.split("\n")

    for line in lines:

        width = text_width(
            line,
            font,
        )

        if width > max_width:
            return False

    return True


def validate_slot(slot):
    """
    Проверяет, что slot находится
    внутри границ слайда.
    """

    if slot["x"] < 0:
        return False

    if slot["y"] < 0:
        return False

    if slot["x"] + slot["w"] > SLIDE_WIDTH_CM:
        return False

    if slot["y"] + slot["h"] > SLIDE_HEIGHT_CM:
        return False

    return True


# ============================================================
# ДОБАВЛЕНИЕ ФОНА-КАРТИНКИ
# ============================================================

def add_background_image(
    slide,
    image_path,
):
    """
    Добавляет изображение на весь слайд.

    Картинка добавляется первой фигурой,
    поэтому весь остальной контент будет
    находиться поверх неё.
    """

    if not os.path.exists(image_path):

        raise FileNotFoundError(
            "\n"
            "Не найден файл фонового изображения:\n"
            f"{image_path}"
        )

    slide.shapes.add_picture(
        image_path,
        0,
        0,
        width=Cm(SLIDE_WIDTH_CM),
        height=Cm(SLIDE_HEIGHT_CM),
    )


# ============================================================
# ДОБАВЛЕНИЕ ТЕКСТА
# ============================================================

def add_textbox(
    slide,
    text,
    slot,
    font_config,
    color,
    align=PP_ALIGN.LEFT,
    max_lines=1,
):
    """
    Добавляет текст в slot.
    """

    # --------------------------------------------------------
    # Проверяем положение блока
    # --------------------------------------------------------

    if not validate_slot(slot):

        raise ValueError(
            "\n"
            "Слот выходит за границы слайда!\n"
            f"Слот: {slot}"
        )

    # --------------------------------------------------------
    # Проверяем количество строк
    # --------------------------------------------------------

    if len(text.split("\n")) > max_lines:

        raise ValueError(
            "\n"
            "Слишком много строк!\n"
            f"Разрешено: {max_lines}\n"
            f"Текст: {text}"
        )

    # --------------------------------------------------------
    # Проверяем ширину текста
    # --------------------------------------------------------

    if not text_fits(
        text,
        slot,
        font_config,
    ):

        raise ValueError(
            "\n"
            "Текст не помещается в шаблон!\n\n"
            f"Текст:\n{text}\n\n"
            f"Шрифт:\n"
            f"{font_config['family']} "
            f"{font_config['size']} pt\n\n"
            f"Ширина блока:\n"
            f"{slot['w']} см\n\n"
            "Сделай текст короче."
        )

    # --------------------------------------------------------
    # Создаём textbox
    # --------------------------------------------------------

    textbox = slide.shapes.add_textbox(
        Cm(slot["x"]),
        Cm(slot["y"]),
        Cm(slot["w"]),
        Cm(slot["h"]),
    )

    text_frame = textbox.text_frame

    text_frame.clear()

    # --------------------------------------------------------
    # ВЕРТИКАЛЬНОЕ ВЫРАВНИВАНИЕ
    # --------------------------------------------------------

    vertical_align = slot.get(
        "vertical_align",
        "middle"
    )

    if vertical_align == "top":

        text_frame.vertical_anchor = MSO_ANCHOR.TOP

    elif vertical_align == "bottom":

        text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM

    else:

        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --------------------------------------------------------
    # ВНУТРЕННИЕ ОТСТУПЫ
    # --------------------------------------------------------

    margins = TEMPLATE["text_box_defaults"]["margins"]

    text_frame.margin_left = Cm(
        margins["left"]
    )

    text_frame.margin_right = Cm(
        margins["right"]
    )

    text_frame.margin_top = Cm(
        margins["top"]
    )

    text_frame.margin_bottom = Cm(
        margins["bottom"]
    )

    # --------------------------------------------------------
    # ТЕКСТ
    # --------------------------------------------------------

    paragraph = text_frame.paragraphs[0]

    paragraph.alignment = align

    run = paragraph.add_run()

    run.text = text

    # --------------------------------------------------------
    # ШРИФТ
    # --------------------------------------------------------

    run.font.name = font_config["family"]

    run.font.size = Pt(
        font_config["size"]
    )

    run.font.bold = (
        font_config["weight"] == "bold"
    )

    run.font.italic = False

    # --------------------------------------------------------
    # ЦВЕТ
    # --------------------------------------------------------

    run.font.color.rgb = hex_to_rgb(
        color
    )

    return textbox


# ============================================================
# ДОБАВЛЕНИЕ КАРТИНКИ
# ============================================================

def add_image(
    slide,
    image_path,
    slot,
):
    """
    Добавляет картинку по координатам
    x/y/w/h.
    """

    if not validate_slot(slot):

        raise ValueError(
            "\n"
            "Изображение выходит за границы слайда!\n"
            f"Слот: {slot}"
        )

    if not os.path.exists(image_path):

        raise FileNotFoundError(
            f"Не найден файл изображения:\n"
            f"{image_path}"
        )

    slide.shapes.add_picture(
        image_path,
        Cm(slot["x"]),
        Cm(slot["y"]),
        Cm(slot["w"]),
        Cm(slot["h"]),
    )


# ============================================================
# ПОИСК ВАРИАНТА
# ============================================================

def get_variant(variant_id):

    for variant in TEMPLATE["variants"]:

        if variant["id"] == variant_id:
            return variant

    raise ValueError(
        f"В шаблоне нет варианта: {variant_id}"
    )


# ============================================================
# СОЗДАНИЕ TITLE-СЛАЙДА
# ============================================================

def create_title_slide(
    prs,
    title,
    subtitle=None,
    text=None,
    text_2=None,
    variant="title_subtitle_center",
):
    """
    Создаёт title-слайд
    по выбранному варианту шаблона.
    """

    # --------------------------------------------------------
    # СОЗДАЁМ ПУСТОЙ СЛАЙД
    # --------------------------------------------------------

    slide = prs.slides.add_slide(
        prs.slide_layouts[6]
    )

    # --------------------------------------------------------
    # ФОН-КАРТИНКА
    # --------------------------------------------------------
    # ВАЖНО:
    # Добавляем её первой.
    # Поэтому все остальные элементы
    # будут находиться поверх неё.
    # --------------------------------------------------------

    add_background_image(
        slide=slide,
        image_path=TEMPLATE["assets"]["background_image"],
    )

    # --------------------------------------------------------
    # ВАРИАНТ
    # --------------------------------------------------------

    variant_data = get_variant(
        variant
    )

    slots = variant_data["slots"]

    # --------------------------------------------------------
    # ЛОГОТИП
    # --------------------------------------------------------

    if "logo" in slots:

        logo = slots["logo"]

        add_image(
            slide=slide,
            image_path=logo["src"],
            slot=logo,
        )

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    if "title" in slots:

        if variant == "title_subtitle_center":

            title_align = PP_ALIGN.CENTER

        else:

            title_align = PP_ALIGN.LEFT

        add_textbox(
            slide=slide,
            text=title,
            slot=slots["title"],
            font_config=TEMPLATE["fonts"]["title"],
            color=TEMPLATE["colors"]["text_primary"],
            align=title_align,
            max_lines=variant_data["constraints"]["title_max_lines"],
        )

    # --------------------------------------------------------
    # SUBTITLE
    # --------------------------------------------------------

    if subtitle and "subtitle" in slots:

        if variant == "title_subtitle_center":

            subtitle_align = PP_ALIGN.CENTER

        else:

            subtitle_align = PP_ALIGN.LEFT

        add_textbox(
            slide=slide,
            text=subtitle,
            slot=slots["subtitle"],
            font_config=TEMPLATE["fonts"]["subtitle"],
            color=TEMPLATE["colors"]["text_primary"],
            align=subtitle_align,
            max_lines=variant_data["constraints"]["subtitle_max_lines"],
        )

    # --------------------------------------------------------
    # TEXT
    # --------------------------------------------------------

    if text and "text" in slots:

        add_textbox(
            slide=slide,
            text=text,
            slot=slots["text"],
            font_config=TEMPLATE["fonts"]["subtitle"],
            color=TEMPLATE["colors"]["text_primary"],
            align=PP_ALIGN.LEFT,
            max_lines=1,
        )

    # --------------------------------------------------------
    # TEXT 2
    # --------------------------------------------------------

    if text_2 and "text_2" in slots:

        add_textbox(
            slide=slide,
            text=text_2,
            slot=slots["text_2"],
            font_config=TEMPLATE["fonts"]["subtitle"],
            color=TEMPLATE["colors"]["text_primary"],
            align=PP_ALIGN.LEFT,
            max_lines=1,
        )

    return slide


# ============================================================
# ТЕСТОВАЯ ПРЕЗЕНТАЦИЯ
# ============================================================

def create_test_presentation(
    output_file=PROJECT_ROOT / "output" / "manual_example.pptx"
):

    prs = Presentation()

    # --------------------------------------------------------
    # РАЗМЕР СЛАЙДА ИЗ ШАБЛОНА
    # --------------------------------------------------------

    prs.slide_width = Cm(
        TEMPLATE["slide_size"]["width"]
    )

    prs.slide_height = Cm(
        TEMPLATE["slide_size"]["height"]
    )

    # --------------------------------------------------------
    # СЛАЙД 1
    # --------------------------------------------------------

    create_title_slide(
        prs=prs,
        title="ИИИИИИИИ",
        subtitle="технологии, которые меняют будущее",
        variant="title_subtitle_center",
    )

    # --------------------------------------------------------
    # СЛАЙД 2
    # --------------------------------------------------------

    create_title_slide(
        prs=prs,
        title="Создание презент",
        subtitle="автоматизация подготовки визуального контента",
        text="ааааа",
        text_2="bbbbb",
        variant="title_subtitle_left",
    )

    # --------------------------------------------------------
    # СЛАЙД 3
    # --------------------------------------------------------

    create_title_slide(
        prs=prs,
        title="VK Tech",
        subtitle="тестирование шаблона презентации",
        text="ааааа",
        variant="title_subtitle_left_2",
    )

    # --------------------------------------------------------
    # СОХРАНЕНИЕ
    # --------------------------------------------------------

    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_file)

    print(
        f"Презентация создана: {output_file}"
    )


# ============================================================
# ЗАПУСК
# ============================================================

if __name__ == "__main__":

    try:

        create_test_presentation()

    except ValueError as error:

        print("\n❌ ОШИБКА:")
        print(error)

    except FileNotFoundError as error:

        print("\n❌ ОШИБКА:")
        print(error)
