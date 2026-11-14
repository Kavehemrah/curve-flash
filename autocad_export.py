"""Draw Curve Flash critical points and dimensions in the active AutoCAD drawing.

Excel columns are mapped as: A=point number, B=Northing/Y, C=Easting/X.
Only connects to an already-running AutoCAD instance; it never launches AutoCAD.
"""
from __future__ import annotations

from math import hypot


LAYER_NAME = "FLASH_CHECK"


def _xyz(x: float, y: float) -> tuple[float, float, float]:
    return float(x), float(y), 0.0


def draw_flash_results(points, results, layer_name: str = LAYER_NAME) -> dict:
    """Draw one red marker, perpendicular and aligned dimension per chord result."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "کتابخانه pywin32 نصب نیست. دستور python -m pip install pywin32==312 را اجرا کنید."
        ) from exc

    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True

        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception as exc:
            raise RuntimeError(
                "به AutoCAD متصل نشدم. ابتدا AutoCAD را اجرا کنید و یک نقشه را باز نگه دارید."
            ) from exc

        doc = acad.ActiveDocument
        if doc is None:
            raise RuntimeError("در AutoCAD نقشه فعالی باز نیست.")

        try:
            layer = doc.Layers.Item(layer_name)
        except Exception:
            layer = doc.Layers.Add(layer_name)

        layer.Color = 1  # AutoCAD Color Index: red
        layer.LayerOn = True
        layer.Lock = False

        if not points or not results:
            raise RuntimeError("داده‌ای برای ترسیم وجود ندارد.")

        by_number = {p.number: p for p in points}
        xs = [p.easting for p in points]
        ys = [p.northing for p in points]
        extent = max(max(xs) - min(xs), max(ys) - min(ys), 1.0)
        marker_radius = max(extent * 0.001, 0.01)
        text_height = marker_radius * 2.5
        model = doc.ModelSpace
        drawn = 0
        skipped = []

        for start_no, end_no, critical_no, signed_flash, abs_flash, _count in results:
            if start_no not in by_number or end_no not in by_number or critical_no not in by_number:
                skipped.append((start_no, end_no, critical_no))
                continue

            a, b, p = by_number[start_no], by_number[end_no], by_number[critical_no]
            dx = b.easting - a.easting
            dy = b.northing - a.northing
            length_sq = dx * dx + dy * dy
            if length_sq <= 1e-20:
                skipped.append((start_no, end_no, critical_no))
                continue

            # Projection on the infinite chord line, consistent with the calculation.
            t = ((p.easting - a.easting) * dx + (p.northing - a.northing) * dy) / length_sq
            foot_x = a.easting + t * dx
            foot_y = a.northing + t * dy
            point_xyz = _xyz(p.easting, p.northing)
            foot_xyz = _xyz(foot_x, foot_y)

            circle = model.AddCircle(point_xyz, marker_radius)
            circle.Layer = layer_name
            circle.Color = 1

            # Cross marker makes the critical point visible at different zoom levels.
            for x1, y1, x2, y2 in (
                (p.easting - marker_radius * 1.5, p.northing - marker_radius * 1.5,
                 p.easting + marker_radius * 1.5, p.northing + marker_radius * 1.5),
                (p.easting - marker_radius * 1.5, p.northing + marker_radius * 1.5,
                 p.easting + marker_radius * 1.5, p.northing - marker_radius * 1.5),
            ):
                mark = model.AddLine(_xyz(x1, y1), _xyz(x2, y2))
                mark.Layer = layer_name
                mark.Color = 1

            perpendicular = model.AddLine(point_xyz, foot_xyz)
            perpendicular.Layer = layer_name
            perpendicular.Color = 1

            foot_mark = model.AddCircle(foot_xyz, marker_radius * 0.45)
            foot_mark.Layer = layer_name
            foot_mark.Color = 1

            vx, vy = p.easting - foot_x, p.northing - foot_y
            distance = hypot(vx, vy)
            if distance > 1e-12:
                dim_x = (p.easting + foot_x) / 2 + (vx / distance) * marker_radius * 4
                dim_y = (p.northing + foot_y) / 2 + (vy / distance) * marker_radius * 4
            else:
                dim_x = (p.easting + foot_x) / 2
                dim_y = (p.northing + foot_y) / 2 + marker_radius * 4

            dimension = model.AddDimAligned(point_xyz, foot_xyz, _xyz(dim_x, dim_y))
            dimension.Layer = layer_name
            dimension.Color = 1
            dimension.TextOverride = f"{abs_flash:.3f}"

            label = model.AddText(
                f"P{critical_no}",
                _xyz(p.easting + marker_radius * 2, p.northing + marker_radius * 2),
                text_height,
            )
            label.Layer = layer_name
            label.Color = 1
            drawn += 1

        doc.Regen(1)
        return {
            "drawing": str(doc.Name),
            "layer": layer_name,
            "dimensions": drawn,
            "skipped": len(skipped),
        }
    finally:
        if initialized:
            pythoncom.CoUninitialize()
