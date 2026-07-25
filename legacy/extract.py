def extract_text(self, shape, ns) -> list[tuple[str, bool]]:
    """
    Возвращает список фрагментов (text, is_bold),
    сохраняя переносы строк как отдельные фрагменты '\n'.
    """
    fragments: list[tuple[str, bool]] = []

    for p in shape.findall(".//a:p", ns):
        paragraph_has_text = False

        for r in p.findall(".//a:r", ns):
            t = r.find("a:t", ns)
            if t is None or not t.text:
                continue

            rpr = r.find("a:rPr", ns)
            is_bold = False
            if rpr is not None:
                bold = rpr.attrib.get("b")
                if bold in ("1", "true", "True"):
                    is_bold = True

            fragments.append((t.text, is_bold))
            paragraph_has_text = True

        # если в абзаце был текст — фиксируем перенос строки
        if paragraph_has_text:
            fragments.append(("\n", False))

    # убираем последний лишний перенос строки
    if fragments and fragments[-1][0] == "\n":
        fragments.pop()

    return fragments
