# ============================================================
# modules/xls_writer.py — прямая запись .xls (BIFF8) через xlwt
#
# ЗАМЕНЯЕТ старую цепочку openpyxl(.xlsx) -> win32com(Excel) -> .xls:
# раньше временный .xlsx собирался через openpyxl, а затем открывался
# живым процессом MS Excel через win32com и пересохранялся в .xls.
# Это работало только на Windows с установленным Excel, было на
# порядок медленнее (запуск Excel, COM-вызовы) и падало с
# RPC_E_CALL_REJECTED при конкурентном доступе к Excel.
#
# xlwt пишет бинарный .xls (BIFF8) напрямую и кроссплатформенно, без
# зависимости от Excel. Стили (шрифт/заливка/границы/выравнивание)
# переносятся поячеечно из исходных openpyxl-ячеек в XlsStyleBuilder,
# с кэшированием XF-стилей и цветовой палитры — иначе на каждую ячейку
# создавался бы новый XF-объект, а BIFF8 допускает не более ~4094
# уникальных стилей на книгу, и это было бы заметно медленнее.
#
# ВАЖНО: цветовая палитра .xls назначается на уровне КНИГИ (workbook),
# поэтому один экземпляр XlsStyleBuilder привязан к одной xlwt.Workbook
# и должен создаваться заново для каждого нового файла.
# ============================================================
import xlwt

# ── границы: openpyxl style-name -> код границы xlwt ──────────
_BORDER_STYLE_MAP = {
    "thin": xlwt.Borders.THIN,
    "medium": xlwt.Borders.MEDIUM,
    "thick": xlwt.Borders.THICK,
    "double": xlwt.Borders.DOUBLE,
    "hair": xlwt.Borders.HAIR,
    "dashed": xlwt.Borders.DASHED,
    "dotted": xlwt.Borders.DOTTED,
    "mediumDashed": xlwt.Borders.MEDIUM_DASHED,
    "dashDot": xlwt.Borders.THIN_DASH_DOTTED,
    "mediumDashDot": xlwt.Borders.MEDIUM_DASH_DOTTED,
    "dashDotDot": xlwt.Borders.THIN_DASH_DOT_DOTTED,
    "mediumDashDotDot": xlwt.Borders.MEDIUM_DASH_DOT_DOTTED,
    "slantDashDot": xlwt.Borders.MEDIUM_DASH_DOTTED,
}

_HORZ_MAP = {
    "left": xlwt.Alignment.HORZ_LEFT,
    "center": xlwt.Alignment.HORZ_CENTER,
    "centerContinuous": xlwt.Alignment.HORZ_CENTER_ACROSS_SEL,
    "right": xlwt.Alignment.HORZ_RIGHT,
    "justify": xlwt.Alignment.HORZ_JUSTIFIED,
    "distributed": xlwt.Alignment.HORZ_DISTRIBUTED,
    "fill": xlwt.Alignment.HORZ_FILLED,
    "general": xlwt.Alignment.HORZ_GENERAL,
}
_VERT_MAP = {
    "top": xlwt.Alignment.VERT_TOP,
    "center": xlwt.Alignment.VERT_CENTER,
    "bottom": xlwt.Alignment.VERT_BOTTOM,
    "justify": xlwt.Alignment.VERT_JUSTIFIED,
    "distributed": xlwt.Alignment.VERT_DISTRIBUTED,
}

# «Пользовательская» часть палитры BIFF8 (0x08..0x3F), безопасная для
# переопределения через set_colour_RGB. 0x08/0x09 сразу занимаем под
# чёрный/белый — самые частые цвета в таблицах, чтобы не тратить на них
# отдельные слоты.
_FIRST_CUSTOM_INDEX = 0x0A
_LAST_CUSTOM_INDEX = 0x3F


class XlsStyleBuilder:
    """Конвертер стилей openpyxl -> xlwt для ОДНОЙ книги xlwt.Workbook.
    Создавать заново для каждого выходного .xls файла."""

    def __init__(self, wb: xlwt.Workbook):
        self.wb = wb
        self._color_index = {"000000": 0x08, "FFFFFF": 0x09}
        self._next_index = _FIRST_CUSTOM_INDEX
        self._style_cache = {}

    # ── регистрация/поиск цвета в палитре книги ───────────────
    def _color_idx(self, hex_color):
        if not hex_color:
            return None
        hex_color = str(hex_color).upper()
        if len(hex_color) == 8:  # ARGB -> RGB
            hex_color = hex_color[2:]
        if len(hex_color) != 6:
            return None
        idx = self._color_index.get(hex_color)
        if idx is not None:
            return idx
        if self._next_index > _LAST_CUSTOM_INDEX:
            # Палитра исчерпана (в проекте используется небольшой набор
            # фирменных цветов сетей — на практике этот предел не
            # достигается). Переиспользуем последний назначенный цвет,
            # чтобы не падать с ошибкой.
            return self._next_index - 1
        try:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            self.wb.set_colour_RGB(self._next_index, r, g, b)
        except Exception:
            return None
        idx = self._next_index
        self._color_index[hex_color] = idx
        self._next_index += 1
        return idx

    @staticmethod
    def _rgb_of(color_obj):
        if color_obj is None:
            return None
        rgb = getattr(color_obj, "rgb", None)
        if not rgb or not isinstance(rgb, str):
            return None
        return rgb

    def style_for_cell(self, cell, number_format_override: str = None) -> xlwt.XFStyle:
        """Строит (или берёт из кэша) xlwt.XFStyle, соответствующий
        фактическому стилю ячейки openpyxl.

        number_format_override, если задан, участвует в ключе кэша и
        применяется вместо number_format исходной ячейки. Это нужно,
        например, для формулы "ИТОГО" (нужен формат "0", а не формат
        исходной ячейки) — НЕЛЬЗЯ просто присвоить .num_format_str на
        уже полученном объекте: он общий (закэширован) и может
        совпадать с объектом других ячеек с тем же сигнатурным стилем,
        так что прямая мутация тихо испортила бы их формат тоже.
        """
        if not getattr(cell, "has_style", False):
            xf = xlwt.XFStyle()
            if number_format_override:
                xf.num_format_str = number_format_override
            return xf

        font = cell.font
        fill = cell.fill
        border = cell.border
        align = cell.alignment

        fg = None
        if fill is not None and fill.patternType and fill.fgColor:
            fg = self._rgb_of(fill.fgColor)
        font_color = self._rgb_of(font.color) if font and font.color else None

        sides = ("left", "right", "top", "bottom")
        border_sig = None
        if border is not None:
            border_sig = tuple(
                (
                    getattr(border, side).style if getattr(border, side) else None,
                    self._rgb_of(getattr(border, side).color) if getattr(border, side) and getattr(border, side).color else None,
                )
                for side in sides
            )

        sig = (
            font.name if font else None,
            round(font.size) if font and font.size else 10,
            bool(font.bold) if font else False,
            bool(font.italic) if font else False,
            font_color,
            fg,
            border_sig,
            align.horizontal if align else None,
            align.vertical if align else None,
            bool(align.wrap_text) if align else False,
            number_format_override or (str(cell.number_format) if cell.number_format else "General"),
        )
        cached = self._style_cache.get(sig)
        if cached is not None:
            return cached

        xf = xlwt.XFStyle()

        # ── шрифт ───────────────────────────────────────────
        xf_font = xlwt.Font()
        xf_font.name = (font.name if font and font.name else "Arial")[:31]
        size_pt = font.size if font and font.size else 10
        xf_font.height = int(round(size_pt * 20))  # twips
        xf_font.bold = bool(font.bold) if font else False
        xf_font.italic = bool(font.italic) if font else False
        if font_color:
            idx = self._color_idx(font_color)
            if idx is not None:
                xf_font.colour_index = idx
        xf.font = xf_font

        # ── заливка ─────────────────────────────────────────
        if fg:
            idx = self._color_idx(fg)
            if idx is not None:
                pattern = xlwt.Pattern()
                pattern.pattern = xlwt.Pattern.SOLID_PATTERN
                pattern.pattern_fore_colour = idx
                pattern.pattern_back_colour = idx
                xf.pattern = pattern

        # ── границы ─────────────────────────────────────────
        if border is not None:
            borders = xlwt.Borders()
            for side in sides:
                side_obj = getattr(border, side)
                style_name = side_obj.style if side_obj else None
                code = _BORDER_STYLE_MAP.get(style_name, xlwt.Borders.NO_LINE)
                color_hex = self._rgb_of(side_obj.color) if side_obj and side_obj.color else None
                color_idx = self._color_idx(color_hex) if color_hex else 0x08
                setattr(borders, side, code)
                setattr(borders, f"{side}_colour", color_idx if color_idx is not None else 0x08)
            xf.borders = borders

        # ── выравнивание ────────────────────────────────────
        if align is not None:
            xf_align = xlwt.Alignment()
            xf_align.horz = _HORZ_MAP.get(align.horizontal, xlwt.Alignment.HORZ_GENERAL)
            xf_align.vert = _VERT_MAP.get(align.vertical, xlwt.Alignment.VERT_CENTER)
            xf_align.wrap = 1 if align.wrap_text else 0
            xf.alignment = xf_align

        # ── числовой формат ─────────────────────────────────
        effective_format = number_format_override or cell.number_format
        if effective_format and effective_format != "General":
            try:
                xf.num_format_str = effective_format
            except Exception:
                pass

        self._style_cache[sig] = xf
        return xf


def col_width_xls(chars: float) -> int:
    """Переводит ширину колонки openpyxl (в 'символах') в единицы xlwt
    (1/256 ширины символа '0')."""
    if not chars:
        chars = 10
    return int(max(chars, 1) * 256)


def row_height_xls(points: float) -> int:
    """Переводит высоту строки в points (openpyxl) в twips (1/20 точки —
    единица высоты строки в xlwt)."""
    if not points:
        points = 15
    return int(max(points, 1) * 20)
