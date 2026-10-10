"""Render the Mistri AI Stage 2 system-context diagram with exact IDs."""

from PIL import Image, ImageDraw, ImageFont

W, H = 2400, 1320
BG = "#F4F1EC"
INK = "#1C1917"
MUTED = "#57534E"
WHITE = "#FFFFFF"

USER = "#0F766E"
USER_FILL = "#CCFBF1"
CHAN = "#1D4ED8"
CHAN_FILL = "#DBEAFE"
CLOUD = "#0F172A"
MOD_LINE = "#0369A1"
DATA = "#9A3412"
DATA_FILL = "#FFEDD5"
EXT = "#6D28D9"
EXT_FILL = "#EDE9FE"
OUT = "#78716C"
OUT_FILL = "#E7E5E4"
ARROW = "#44403C"
LANG = "#115E59"

SANS = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

font_title = ImageFont.truetype(SANS_B, 38)
font_sub = ImageFont.truetype(SANS, 20)
font_id = ImageFont.truetype(SANS_B, 17)
font_name = ImageFont.truetype(SANS, 15)
font_small = ImageFont.truetype(SANS, 14)
font_tiny = ImageFont.truetype(SANS, 13)
font_head = ImageFont.truetype(SANS_B, 16)
font_arrow = ImageFont.truetype(SANS, 13)


def text_size(draw, text, font):
    box = draw.textbbox((0, 0), text, font=font)
    return box[2] - box[0], box[3] - box[1]


def round_rect(draw, xy, radius, fill, outline, width=2):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def center_text(draw, xy, text, font, fill):
    x0, y0, x1, y1 = xy
    tw, th = text_size(draw, text, font)
    draw.text(
        (x0 + (x1 - x0 - tw) / 2, y0 + (y1 - y0 - th) / 2 - 1),
        text,
        font=font,
        fill=fill,
    )


def wrap(draw, text, font, max_w):
    words = text.split()
    lines, current = [], ""
    for word in words:
        trial = word if not current else current + " " + word
        if text_size(draw, trial, font)[0] <= max_w:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def card(draw, xy, id_text, name, fill, outline):
    round_rect(draw, xy, 12, fill, outline, 2)
    x0, y0, x1, y1 = xy
    center_text(draw, (x0, y0 + 6, x1, y0 + 30), id_text, font_id, outline)
    lines = wrap(draw, name, font_name, x1 - x0 - 16)
    total_h = len(lines) * 18
    y = y0 + 34 + max(0, (y1 - y0 - 42 - total_h) / 2)
    for line in lines:
        tw, _ = text_size(draw, line, font_name)
        draw.text((x0 + (x1 - x0 - tw) / 2, y), line, font=font_name, fill=INK)
        y += 18


def arrow_head(draw, start, end):
    ex, ey = end
    sx, sy = start
    if abs(ex - sx) >= abs(ey - sy):
        sign = 1 if ex > sx else -1
        head = [(ex, ey), (ex - sign * 12, ey - 6), (ex - sign * 12, ey + 6)]
    else:
        sign = 1 if ey > sy else -1
        head = [(ex, ey), (ex - 6, ey - sign * 12), (ex + 6, ey - sign * 12)]
    draw.polygon(head, fill=ARROW)


def labeled_arrow(draw, start, end, label=None, label_at=0.5):
    draw.line([start, end], fill=ARROW, width=3)
    arrow_head(draw, start, end)
    if not label:
        return
    sx, sy = start
    ex, ey = end
    mx = sx + (ex - sx) * label_at
    my = sy + (ey - sy) * label_at
    tw, th = text_size(draw, label, font_arrow)
    pad = 3
    draw.rectangle(
        (mx - tw / 2 - pad, my - th / 2 - pad, mx + tw / 2 + pad, my + th / 2 + pad),
        fill=BG,
    )
    draw.text((mx - tw / 2, my - th / 2 - 1), label, font=font_arrow, fill=MUTED)


def main():
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    draw.text((40, 28), "Mistri AI — system context", font=font_title, fill=INK)
    draw.text(
        (40, 76),
        "Delhi pilot   ·   Stage 2 design   ·   every ID is the trace key for later quantities and costs",
        font=font_sub,
        fill=MUTED,
    )

    # Vertical bands. Users and channels are centred on the cloud.
    cloud = (700, 150, 1680, 690)
    band_top, band_bot = cloud[1], cloud[3]
    band_h = band_bot - band_top

    draw.text((40, 124), "USERS", font=font_head, fill=USER)
    draw.text((360, 124), "CHANNELS", font=font_head, fill=CHAN)
    draw.text((700, 124), "APPLICATION", font=font_head, fill=CLOUD)
    draw.text((1720, 124), "DATA STORES", font=font_head, fill=DATA)

    user_h, user_w = 120, 230
    user_gap = 28
    user_block = 3 * user_h + 2 * user_gap
    user_y0 = band_top + (band_h - user_block) / 2
    users = [
        ("USR-W", "Construction worker"),
        ("USR-C", "Contractor"),
        ("USR-V", "Vendor"),
    ]
    user_boxes = []
    for i, (uid, name) in enumerate(users):
        y = user_y0 + i * (user_h + user_gap)
        box = (40, y, 40 + user_w, y + user_h)
        card(draw, box, uid, name, USER_FILL, USER)
        user_boxes.append(box)

    chan_x, chan_w = 360, 220
    chan_boxes = []
    channels = [
        ("CH-AND", "Android app"),
        ("CH-WA", "WhatsApp"),
        ("CH-WEB", "Website"),
    ]
    for i, (uid, name) in enumerate(channels):
        y = user_y0 + i * (user_h + user_gap)
        box = (chan_x, y, chan_x + chan_w, y + user_h)
        card(draw, box, uid, name, CHAN_FILL, CHAN)
        chan_boxes.append(box)

    # One arrow for the whole user group: every role uses every channel.
    # Draw a bus on the right of the users and the left of the channels.
    bus_x = 312
    top_y = (user_boxes[0][1] + user_boxes[0][3]) / 2
    bot_y = (user_boxes[2][1] + user_boxes[2][3]) / 2
    draw.line([(bus_x, top_y), (bus_x, bot_y)], fill=ARROW, width=3)
    for box in user_boxes:
        mid = (box[1] + box[3]) / 2
        draw.line([(box[2], mid), (bus_x, mid)], fill=ARROW, width=3)
    for box in chan_boxes:
        mid = (box[1] + box[3]) / 2
        labeled_arrow(draw, (bus_x, mid), (box[0] - 1, mid))

    # Channels into the cloud, one arrow each, so no role is tied to one channel.
    for box in chan_boxes:
        mid = (box[1] + box[3]) / 2
        labeled_arrow(draw, (box[2] + 2, mid), (cloud[0] - 2, mid))
    draw.text((460, 124), "every role uses all three", font=font_arrow, fill=MUTED)

    round_rect(draw, cloud, 16, "#F8FAFC", CLOUD, 3)
    draw.text((724, 164), "PLT-CLOUD", font=font_id, fill=CLOUD)
    draw.text((860, 164), "Public cloud hosted in India", font=font_name, fill=INK)
    lang = (724, 196, 1656, 228)
    round_rect(draw, lang, 8, "#CCFBF1", LANG, 1)
    center_text(draw, lang, "PLT-LANG     Hindi and English on Android, WhatsApp, and the website", font_small, LANG)

    modules = [
        ("MOD-AUTH", "Phone OTP sign-in"),
        ("MOD-PROF", "Role and profile"),
        ("MOD-HOME", "Home and my jobs"),
        ("MOD-JOB", "Job intake: text, audio, photo"),
        ("MOD-EST", "Estimate and quantities"),
        ("MOD-CAT", "Vendor catalogue"),
        ("MOD-MATCH", "Material match by locality"),
        ("MOD-BOOK", "Find and book by trade"),
        ("MOD-SITE", "Site progress and attendance"),
        ("MOD-INV", "Internal invoice status"),
    ]
    left, top = 724, 248
    gap_x, gap_y = 12, 12
    mw, mh = 180, 190
    module_boxes = {}
    for i, (uid, name) in enumerate(modules):
        c, r = i % 5, i // 5
        x = left + c * (mw + gap_x)
        y = top + r * (mh + gap_y)
        box = (x, y, x + mw, y + mh)
        card(draw, box, uid, name, WHITE, MOD_LINE)
        module_boxes[uid] = box

    stores = [
        ("DAT-ACC", "Accounts, role, trade"),
        ("DAT-JOB", "Jobs, audio, photos"),
        ("DAT-EST", "Estimates, quantities"),
        ("DAT-CAT", "Prices, stock, locality"),
        ("DAT-BOOK", "Bookings"),
        ("DAT-SITE", "Attendance and location"),
        ("DAT-INV", "Invoices, paid or unpaid"),
    ]
    # Two columns so the store group matches the cloud height.
    store_w, store_h, store_gap = 280, 118, 14
    store_x = [1788, 2084]
    for i, (uid, name) in enumerate(stores):
        col = 0 if i < 4 else 1
        row = i if i < 4 else i - 4
        # second column of 3, centred against the four
        y_off = 0 if col == 0 else (store_h + store_gap) / 2
        y = cloud[1] + y_off + row * (store_h + store_gap)
        card(draw, (store_x[col], y, store_x[col] + store_w, y + store_h), uid, name, DATA_FILL, DATA)

    mid_y = (cloud[1] + cloud[3]) / 2
    labeled_arrow(draw, (cloud[2] + 2, mid_y), (store_x[0] - 4, mid_y))
    rw = "read / write"
    tw, th = text_size(draw, rw, font_arrow)
    rx = (cloud[2] + store_x[0]) / 2 - tw / 2
    draw.rectangle((rx - 3, mid_y - th - 8, rx + tw + 3, mid_y - 4), fill=BG)
    draw.text((rx, mid_y - th - 6), rw, font=font_arrow, fill=MUTED)

    # External systems
    ext_top = 760
    draw.text((40, ext_top), "EXTERNAL SYSTEMS", font=font_head, fill=EXT)
    otp = (40, ext_top + 32, 520, ext_top + 168)
    waba = (548, ext_top + 32, 1028, ext_top + 168)
    card(draw, otp, "EXT-OTP", "SMS one-time-password gateway. Called by MOD-AUTH.", EXT_FILL, EXT)
    card(draw, waba, "EXT-WABA", "WhatsApp Business API. Carries CH-WA.", EXT_FILL, EXT)

    out = (1056, ext_top + 32, 2356, ext_top + 168)
    round_rect(draw, out, 12, OUT_FILL, OUT, 2)
    center_text(draw, (1056, ext_top + 44, 2356, ext_top + 80), "NOT CONNECTED IN THIS VERSION", font_id, OUT)
    center_text(
        draw,
        (1056, ext_top + 84, 2356, ext_top + 150),
        "UPI and payment apps    ·    GST invoicing    ·    Map services",
        font_name,
        INK,
    )

    # MOD-AUTH down to EXT-OTP, routed in the left gutter of the cloud
    auth = module_boxes["MOD-AUTH"]
    auth_mid_x = (auth[0] + auth[2]) / 2
    # route: down from auth to just inside cloud, left, out, down to otp
    draw.line(
        [
            (auth_mid_x, auth[3]),
            (auth_mid_x, cloud[3] - 16),
            (cloud[0] + 16, cloud[3] - 16),
            (cloud[0] + 16, cloud[3]),
            ((otp[0] + otp[2]) / 2, cloud[3]),
            ((otp[0] + otp[2]) / 2, otp[1]),
        ],
        fill=ARROW,
        width=3,
    )
    arrow_head(draw, ((otp[0] + otp[2]) / 2, otp[1] - 16), ((otp[0] + otp[2]) / 2, otp[1]))
    draw.text((80, cloud[3] + 8), "sends OTP", font=font_arrow, fill=MUTED)

    # Reading notes
    notes_y = 970
    draw.text((40, notes_y), "HOW A JOB MOVES", font=font_head, fill=INK)
    notes = [
        "USR-W, USR-C, and USR-V each sign in on CH-AND, CH-WA, or CH-WEB. The same phone number is the same account.",
        "MOD-JOB stores the description as text, audio, or a photo in DAT-JOB. MOD-EST writes the quantity list into DAT-EST.",
        "MOD-CAT holds the price, stock, and Delhi locality the vendor types into DAT-CAT. MOD-MATCH uses that locality. It does not call a map.",
        "MOD-BOOK writes DAT-BOOK. MOD-SITE writes attendance, photos, and the saved location into DAT-SITE. MOD-INV writes paid or unpaid into DAT-INV.",
        "Screens SCR-01 to SCR-13 and WhatsApp flows SCR-W1 to SCR-W7 sit on these modules. The design table names every screen.",
    ]
    y = notes_y + 30
    for note in notes:
        draw.ellipse((48, y + 4, 60, y + 16), fill=CLOUD)
        draw.text((72, y), note, font=font_small, fill=INK)
        y += 26

    out_path = "/workspace/design/mistri-ai-system-context.png"
    img.save(out_path, "PNG")
    print(out_path)


if __name__ == "__main__":
    main()
