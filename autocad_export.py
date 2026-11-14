"""Draw one midpoint flash point and its dimension in active AutoCAD.

Excel columns: A=point number, B=Northing/Y, C=Easting/X.
Only connects to an already-running AutoCAD instance.
"""
from __future__ import annotations


LAYER_NAME = "FLASH_CHECK"


def _xyz(x: float, y: float, pythoncom, win32com):
    """Build an AutoCAD-compatible SAFEARRAY of three doubles."""
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (float(x), float(y), 0.0),
    )


def draw_flash_results(points, results, layer_name: str = LAYER_NAME) -> dict:
    """Draw exactly one red point and one aligned dimension per chord."""
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

        layer.Color = 1
        layer.LayerOn = True
        layer.Lock = False

        if not points or not results:
            raise RuntimeError("داده‌ای برای ترسیم وجود ندارد.")

        by_number = {p.number: p for p in points}
        model = doc.ModelSpace
        drawn = 0
        skipped = []

        for start_no, end_no, midpoint_no, signed_flash, abs_flash, _count in results:
            if start_no not in by_number or end_no not in by_number or midpoint_no not in by_number:
                skipped.append((start_no, end_no, midpoint_no))
                continue

            a = by_number[start_no]
            b = by_number[end_no]
            midpoint = by_number[midpoint_no]
            dx = b.easting - a.easting
            dy = b.northing - a.northing
            length_sq = dx * dx + dy * dy
            if length_sq <= 1e-20:
                skipped.append((start_no, end_no, midpoint_no))
                continue

            # Foot of the perpendicular from the single midpoint to the chord.
            t = ((midpoint.easting - a.easting) * dx +
                 (midpoint.northing - a.northing) * dy) / length_sq
            foot_x = a.easting + t * dx
            foot_y = a.northing + t * dy

            point_xyz = _xyz(midpoint.easting, midpoint.northing, pythoncom, win32com)
            foot_xyz = _xyz(foot_x, foot_y, pythoncom, win32com)

            # Exactly one point entity for the midpoint.
            point_entity = model.AddPoint(point_xyz)
            point_entity.Layer = layer_name
            point_entity.Color = 1

            # The aligned dimension itself displays the perpendicular flash.
            # Offset the dimension line slightly from the perpendicular midpoint.
            vx = midpoint.easting - foot_x
            vy = midpoint.northing - foot_y
            distance = (vx * vx + vy * vy) ** 0.5
            if distance > 1e-12:
                dim_x = (midpoint.easting + foot_x) / 2 + (vx / distance) * max(distance * 0.15, 0.01)
                dim_y = (midpoint.northing + foot_y) / 2 + (vy / distance) * max(distance * 0.15, 0.01)
            else:
                # A zero flash cannot form a non-zero aligned dimension.
                # Keep the point and report it as skipped rather than adding
                # misleading or overlapping dimension geometry.
                skipped.append((start_no, end_no, midpoint_no))
                continue

            dimension = model.AddDimAligned(
                point_xyz,
                foot_xyz,
                _xyz(dim_x, dim_y, pythoncom, win32com),
            )
            dimension.Layer = layer_name
            dimension.Color = 1
            dimension.TextOverride = f"{abs_flash:.3f}"
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
