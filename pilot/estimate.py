"""Demo quantity rules for the local website pilot.

These factors scale a typed built-up area. They are not a published schedule of rates.
Rupee amounts stay empty until a vendor has typed a price for that material.
"""

RULES = {
    "JOB-NEW": [
        ("Cement", "bags", 0.08),
        ("Sand", "cubic metre", 0.012),
        ("Bricks", "numbers", 5),
        ("Steel", "kg", 0.4),
        ("Floor tiles", "sq ft", 1),
    ],
    "JOB-EXT": [
        ("Cement", "bags", 0.06),
        ("Sand", "cubic metre", 0.008),
        ("Bricks", "numbers", 4),
        ("Steel", "kg", 0.3),
    ],
    "JOB-FIT": [
        ("Floor tiles", "sq ft", 1),
        ("Paint", "litre", 0.02),
        ("Cement", "bags", 0.01),
    ],
}

LABELS = {
    "JOB-NEW": "New building",
    "JOB-EXT": "Extension or extra floor",
    "JOB-FIT": "Interior fit-out",
}


def quantity_lines(job_type, area_sqft):
    rules = RULES.get(job_type)
    if not rules:
        return []
    area = float(area_sqft)
    lines = []
    for material, unit, per_sqft in rules:
        qty = round(per_sqft * area, 2)
        lines.append({"material": material, "unit": unit, "quantity": qty, "rate": None, "amount": None, "source": "RATE NEEDED"})
    return lines
