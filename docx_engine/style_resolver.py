# -*- coding: utf-8 -*-
"""
Style Resolver — parse styles.xml and compute effective formatting by
walking the style inheritance chain (basedOn).
"""
from typing import Dict, List, Optional, Any
from lxml import etree

from .xml_parser import NSMAP, _q


class StyleResolver:
    """Resolves effective paragraph and run formatting from style inheritance."""

    def __init__(self, styles_root: Optional[etree._Element] = None):
        self.styles: Dict[str, Dict[str, Any]] = {}
        self._style_chain_cache: Dict[str, Dict[str, Any]] = {}
        if styles_root is not None:
            self.load_styles(styles_root)

    def load_styles(self, styles_root: etree._Element):
        """Parse styles.xml into internal dictionary."""
        for style in styles_root.findall(_q("style")):
            style_id = style.get(_q("styleId"), "")
            if not style_id:
                continue

            style_type = style.get(_q("type"), "paragraph")
            default_attr = style.get(_q("default"), "0")
            is_default = default_attr == "1"

            entry = {
                "styleId": style_id,
                "type": style_type,
                "default": is_default,
                "basedOn": None,
                "name": "",
                "pPr": {},
                "rPr": {},
            }

            # Style name
            name_el = style.find(_q("name"))
            if name_el is not None:
                entry["name"] = name_el.get(_q("val"), "")

            # BasedOn (inheritance)
            based_on = style.find(_q("basedOn"))
            if based_on is not None:
                entry["basedOn"] = based_on.get(_q("val"), "")

            # Paragraph properties
            ppr = style.find(_q("pPr"))
            if ppr is not None:
                entry["pPr"] = self._parse_ppr(ppr)

            # Run properties
            rpr = style.find(_q("rPr"))
            if rpr is not None:
                entry["rPr"] = self._parse_rpr(rpr)

            self.styles[style_id] = entry

    def _parse_ppr(self, ppr: etree._Element) -> Dict[str, Any]:
        props = {}

        jc = ppr.find(_q("jc"))
        if jc is not None:
            props["alignment"] = jc.get(_q("val"), "left")

        ind = ppr.find(_q("ind"))
        if ind is not None:
            props["first_line_indent"] = self._twips_to_pt(ind.get(_q("firstLine")))
            props["left_indent"] = self._twips_to_pt(ind.get(_q("left")))
            props["right_indent"] = self._twips_to_pt(ind.get(_q("right")))
            props["hanging"] = self._twips_to_pt(ind.get(_q("hanging")))

        spacing = ppr.find(_q("spacing"))
        if spacing is not None:
            props["space_before"] = self._twips_to_pt(spacing.get(_q("before")))
            props["space_after"] = self._twips_to_pt(spacing.get(_q("after")))
            line = spacing.get(_q("line"))
            line_rule = spacing.get(_q("lineRule"))
            if line:
                if line_rule == "auto":
                    props["line_spacing"] = int(line) / 240
                else:
                    props["line_spacing"] = self._twips_to_pt(line)
            props["line_rule"] = line_rule

        outline = ppr.find(_q("outlineLvl"))
        if outline is not None:
            props["outline_level"] = int(outline.get(_q("val"), "9"))

        numpr = ppr.find(_q("numPr"))
        if numpr is not None:
            ilvl = numpr.find(_q("ilvl"))
            numId = numpr.find(_q("numId"))
            props["numPr"] = {
                "ilvl": int(ilvl.get(_q("val"), "0")) if ilvl is not None else 0,
                "numId": int(numId.get(_q("val"), "0")) if numId is not None else 0,
            }

        return props

    def _parse_rpr(self, rpr: etree._Element) -> Dict[str, Any]:
        props = {}

        rfonts = rpr.find(_q("rFonts"))
        if rfonts is not None:
            props["font"] = rfonts.get(_q("ascii")) or rfonts.get(_q("eastAsia")) or rfonts.get(_q("hAnsi"))

        sz = rpr.find(_q("sz"))
        if sz is not None:
            props["size"] = int(sz.get(_q("val"), "0")) / 2

        szCs = rpr.find(_q("szCs"))
        if szCs is not None and "size" not in props:
            props["size"] = int(szCs.get(_q("val"), "0")) / 2

        b = rpr.find(_q("b"))
        if b is not None:
            props["bold"] = b.get(_q("val"), "true") != "false"

        i = rpr.find(_q("i"))
        if i is not None:
            props["italic"] = i.get(_q("val"), "true") != "false"

        u = rpr.find(_q("u"))
        if u is not None:
            props["underline"] = u.get(_q("val"), "single")

        color = rpr.find(_q("color"))
        if color is not None:
            props["color"] = color.get(_q("val"), "auto")

        highlight = rpr.find(_q("highlight"))
        if highlight is not None:
            props["highlight"] = highlight.get(_q("val"), "yellow")

        strike = rpr.find(_q("strike"))
        if strike is not None:
            props["strike"] = strike.get(_q("val"), "true") != "false"

        vert_align = rpr.find(_q("vertAlign"))
        if vert_align is not None:
            val = vert_align.get(_q("val"), "baseline")
            if val == "superscript":
                props["superscript"] = True
            elif val == "subscript":
                props["subscript"] = True

        return props

    def _twips_to_pt(self, val: Optional[str]) -> Optional[float]:
        if val is None:
            return None
        try:
            return int(val) / 20
        except ValueError:
            return None

    def resolve_paragraph_format(self, style_id: Optional[str]) -> Dict[str, Any]:
        """Resolve effective paragraph format for a style by walking inheritance chain."""
        if not style_id or style_id not in self.styles:
            return {}

        if style_id in self._style_chain_cache:
            cached = self._style_chain_cache[style_id]
            return cached.get("pPr", {})

        chain = self._walk_chain(style_id)
        effective = {}
        for s in reversed(chain):
            effective.update(s.get("pPr", {}))

        return effective

    def resolve_run_format(self, style_id: Optional[str]) -> Dict[str, Any]:
        """Resolve effective run format for a style by walking inheritance chain."""
        if not style_id or style_id not in self.styles:
            return {}

        if style_id in self._style_chain_cache:
            cached = self._style_chain_cache[style_id]
            return cached.get("rPr", {})

        chain = self._walk_chain(style_id)
        effective = {}
        for s in reversed(chain):
            effective.update(s.get("rPr", {}))

        return effective

    def resolve_full_format(self, style_id: Optional[str]) -> Dict[str, Any]:
        """Resolve both paragraph and run formats."""
        if not style_id or style_id not in self.styles:
            return {"pPr": {}, "rPr": {}}

        if style_id in self._style_chain_cache:
            return self._style_chain_cache[style_id]

        chain = self._walk_chain(style_id)
        ppr = {}
        rpr = {}
        for s in reversed(chain):
            ppr.update(s.get("pPr", {}))
            rpr.update(s.get("rPr", {}))

        result = {"pPr": ppr, "rPr": rpr}
        self._style_chain_cache[style_id] = result
        return result

    def _walk_chain(self, style_id: str, visited: Optional[set] = None) -> List[Dict]:
        """Walk the style inheritance chain."""
        if visited is None:
            visited = set()
        if style_id in visited or style_id not in self.styles:
            return []
        visited.add(style_id)

        style = self.styles[style_id]
        chain = [style]

        if style.get("basedOn"):
            chain = self._walk_chain(style["basedOn"], visited) + chain

        return chain

    def get_style_names(self) -> Dict[str, str]:
        """Return mapping of styleId -> display name."""
        return {sid: info.get("name", sid) for sid, info in self.styles.items()}

    def get_styles_by_type(self, style_type: str) -> Dict[str, Dict]:
        """Return all styles of a given type ('paragraph', 'character', 'table')."""
        return {sid: info for sid, info in self.styles.items() if info.get("type") == style_type}
