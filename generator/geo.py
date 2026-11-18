# -*- coding: utf-8 -*-
"""Koordinata Qashqadaryo viloyati hududida ekanini tekshirish (tashqi kutubxonasiz)."""
import json
import math
from pathlib import Path

import config


class RegionChecker(object):
    """Viloyat chegarasi (GeoJSON Polygon) ichida ekanini tekshiradi.

    Tezlik uchun poligon qirralari kenglik bo'yicha tasmalarga bo'lingan, shuning uchun
    bitta nuqta uchun faqat bir necha qirra tekshiriladi. `tolerance_km` - chegara yaqinidagi
    nuqtalarga beriladigan yo'l qo'yilgan xato (soddalashtirilgan chegara uchun).
    """

    _STRIP = 0.02  # daraja

    def __init__(self, geojson_path=None, tolerance_km=None):
        path = Path(geojson_path or config.REGION_GEOJSON)
        data = json.loads(path.read_text(encoding="utf-8"))
        geom = data["geometry"] if data.get("type") == "Feature" else data
        polys = [geom["coordinates"]] if geom["type"] == "Polygon" else geom["coordinates"]
        self._edges = []
        xs, ys = [], []
        for poly in polys:
            for ring in poly:  # tashqi va ichki (teshik) halqalar even-odd qoidasi bilan
                for i in range(len(ring) - 1):
                    (x1, y1), (x2, y2) = ring[i][:2], ring[i + 1][:2]
                    if y1 != y2:
                        self._edges.append((x1, y1, x2, y2))
                xs += [p[0] for p in ring]
                ys += [p[1] for p in ring]
        self.min_lon, self.max_lon, self.min_lat, self.max_lat = min(xs), max(xs), min(ys), max(ys)
        self._tol = config.BORDER_TOLERANCE_KM if tolerance_km is None else tolerance_km
        self._y0 = self.min_lat
        n = int((self.max_lat - self.min_lat) / self._STRIP) + 2
        self._strips = [[] for _ in range(n)]
        for e in self._edges:
            lo, hi = sorted((e[1], e[3]))
            for s in range(int((lo - self._y0) / self._STRIP), int((hi - self._y0) / self._STRIP) + 1):
                if 0 <= s < n:
                    self._strips[s].append(e)
        self._cache = {}

    def _inside_exact(self, lat, lon):
        s = int((lat - self._y0) / self._STRIP)
        if s < 0 or s >= len(self._strips):
            return False
        inside = False
        for x1, y1, x2, y2 in self._strips[s]:
            if (y1 > lat) != (y2 > lat):
                if lon < (x2 - x1) * (lat - y1) / (y2 - y1) + x1:
                    inside = not inside
        return inside

    def contains(self, lat, lon):
        """True: nuqta viloyat ichida (yoki chegaradan tolerance_km ichida)."""
        if lat is None or lon is None:
            return False
        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            return False
        d = self._tol / 111.0
        dlon = d / max(math.cos(math.radians(lat)), 0.1)
        if not (self.min_lat - d <= lat <= self.max_lat + d and self.min_lon - dlon <= lon <= self.max_lon + dlon):
            return False
        if self._inside_exact(lat, lon):
            return True
        if self._tol <= 0:
            return False
        return any(self._inside_exact(lat + a * d, lon + b * dlon)
                   for a, b in ((1, 0), (-1, 0), (0, 1), (0, -1), (1, 1), (1, -1), (-1, 1), (-1, -1)))
