"""Mistri AI local website pilot. SMS, WhatsApp, and Android are not connected."""

import os
import secrets
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory, session
from werkzeug.utils import secure_filename

from estimate import LABELS, quantity_lines

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
UPLOADS = DATA / "uploads"
DB_PATH = DATA / "mistri.db"
DATA.mkdir(exist_ok=True)
UPLOADS.mkdir(exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = os.environ.get("MISTRI_SECRET", "mistri-local-pilot")
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024


def now():
    return datetime.now(timezone.utc)


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY,
                phone TEXT UNIQUE NOT NULL,
                name TEXT,
                role TEXT,
                trade TEXT,
                locality TEXT
            );
            CREATE TABLE IF NOT EXISTS otps (
                phone TEXT PRIMARY KEY,
                code TEXT NOT NULL,
                expires TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS jobs (
                id INTEGER PRIMARY KEY,
                owner_id INTEGER NOT NULL,
                job_type TEXT NOT NULL,
                area_sqft REAL NOT NULL,
                description TEXT NOT NULL,
                audio_path TEXT,
                photo_path TEXT,
                created TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS estimates (
                job_id INTEGER NOT NULL,
                material TEXT NOT NULL,
                unit TEXT NOT NULL,
                quantity REAL NOT NULL
            );
            CREATE TABLE IF NOT EXISTS catalogue (
                id INTEGER PRIMARY KEY,
                vendor_id INTEGER NOT NULL,
                material TEXT NOT NULL,
                unit TEXT NOT NULL,
                price REAL NOT NULL,
                stock TEXT NOT NULL,
                locality TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS bookings (
                id INTEGER PRIMARY KEY,
                job_id INTEGER NOT NULL,
                requester_id INTEGER NOT NULL,
                person_id INTEGER NOT NULL,
                created TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS site_updates (
                id INTEGER PRIMARY KEY,
                job_id INTEGER NOT NULL,
                author_id INTEGER NOT NULL,
                note TEXT NOT NULL,
                attendance TEXT,
                photo_path TEXT,
                created TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY,
                job_id INTEGER NOT NULL,
                author_id INTEGER NOT NULL,
                vendor_id INTEGER,
                note TEXT NOT NULL,
                status TEXT NOT NULL,
                created TEXT NOT NULL
            );
            """
        )


def current_user():
    uid = session.get("uid")
    if not uid:
        return None
    with connect() as conn:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (uid,)).fetchone()
    return dict(row) if row else None


def require_user():
    user = current_user()
    if not user:
        return None, (jsonify({"error": "Sign in required"}), 401)
    return user, None


def save_upload(file_storage):
    if not file_storage or not file_storage.filename:
        return None
    name = secure_filename(file_storage.filename) or "upload"
    stored = f"{secrets.token_hex(8)}_{name}"
    file_storage.save(UPLOADS / stored)
    return stored


def job_locality(conn, job):
    owner = conn.execute("SELECT locality FROM users WHERE id = ?", (job["owner_id"],)).fetchone()
    return owner["locality"] if owner else ""


def match_lines(conn, job):
    lines = conn.execute("SELECT * FROM estimates WHERE job_id = ?", (job["id"],)).fetchall()
    locality = (job_locality(conn, job) or "").strip().casefold()
    catalogue = conn.execute(
        """
        SELECT c.*, u.name AS vendor_name, u.phone AS vendor_phone
        FROM catalogue c JOIN users u ON u.id = c.vendor_id
        """
    ).fetchall()
    matched = []
    for line in lines:
        material = line["material"].casefold()
        offers = []
        for item in catalogue:
            item_loc = (item["locality"] or "").strip().casefold()
            same_place = locality and item_loc and (locality == item_loc or locality in item_loc or item_loc in locality)
            same_material = material == item["material"].casefold() or material in item["material"].casefold() or item["material"].casefold() in material
            if same_place and same_material:
                amount = round(line["quantity"] * item["price"], 2)
                offers.append(
                    {
                        "vendor_id": item["vendor_id"],
                        "vendor_name": item["vendor_name"],
                        "vendor_phone": item["vendor_phone"],
                        "price": item["price"],
                        "unit": item["unit"],
                        "stock": item["stock"],
                        "locality": item["locality"],
                        "amount": amount,
                        "source": "Price typed by the vendor",
                    }
                )
        matched.append(
            {
                "material": line["material"],
                "unit": line["unit"],
                "quantity": line["quantity"],
                "offers": offers,
                "rate": offers[0]["price"] if offers else None,
                "amount": offers[0]["amount"] if offers else None,
                "source": offers[0]["source"] if offers else "RATE NEEDED",
            }
        )
    return matched


@app.get("/")
def index():
    return send_from_directory(ROOT / "templates", "index.html")


@app.get("/api/health")
def health():
    return jsonify({"ok": True, "service": "Mistri AI local website pilot"})


@app.post("/api/login/start")
def login_start():
    phone = "".join(ch for ch in (request.json or {}).get("phone", "") if ch.isdigit())
    if len(phone) < 10:
        return jsonify({"error": "Enter a 10-digit mobile number"}), 400
    phone = phone[-10:]
    code = f"{secrets.randbelow(1000000):06d}"
    expires = (now() + timedelta(minutes=10)).isoformat()
    with connect() as conn:
        conn.execute(
            "INSERT INTO otps (phone, code, expires) VALUES (?, ?, ?) ON CONFLICT(phone) DO UPDATE SET code = excluded.code, expires = excluded.expires",
            (phone, code, expires),
        )
        conn.execute("INSERT OR IGNORE INTO users (phone) VALUES (?)", (phone,))
    return jsonify({"ok": True, "phone": phone, "demo_code": code, "note": "Shown on screen. No SMS is sent in this local pilot."})


@app.post("/api/login/verify")
def login_verify():
    body = request.json or {}
    phone = "".join(ch for ch in body.get("phone", "") if ch.isdigit())[-10:]
    code = "".join(ch for ch in str(body.get("code", "")) if ch.isdigit())
    with connect() as conn:
        row = conn.execute("SELECT * FROM otps WHERE phone = ?", (phone,)).fetchone()
        if not row or row["code"] != code or row["expires"] < now().isoformat():
            return jsonify({"error": "That code does not match"}), 400
        user = conn.execute("SELECT * FROM users WHERE phone = ?", (phone,)).fetchone()
        conn.execute("DELETE FROM otps WHERE phone = ?", (phone,))
    session["uid"] = user["id"]
    return jsonify({"user": dict(user)})


@app.post("/api/logout")
def logout():
    session.clear()
    return jsonify({"ok": True})


@app.get("/api/me")
def me():
    user = current_user()
    if not user:
        return jsonify({"user": None})
    return jsonify({"user": user})


@app.post("/api/profile")
def profile():
    user, error = require_user()
    if error:
        return error
    body = request.json or {}
    role = body.get("role")
    name = (body.get("name") or "").strip()
    trade = (body.get("trade") or "").strip()
    locality = (body.get("locality") or "").strip()
    if role not in ("worker", "contractor", "vendor") or not name or not trade or not locality:
        return jsonify({"error": "Name, role, trade, and Delhi locality are required"}), 400
    with connect() as conn:
        conn.execute(
            "UPDATE users SET name = ?, role = ?, trade = ?, locality = ? WHERE id = ?",
            (name, role, trade, locality, user["id"]),
        )
        user = dict(conn.execute("SELECT * FROM users WHERE id = ?", (user["id"],)).fetchone())
    return jsonify({"user": user})


@app.post("/api/jobs")
def create_job():
    user, error = require_user()
    if error:
        return error
    if user["role"] not in ("worker", "contractor"):
        return jsonify({"error": "A vendor keeps a catalogue. A worker or contractor opens a job."}), 403
    job_type = request.form.get("job_type", "")
    description = (request.form.get("description") or "").strip()
    try:
        area = float(request.form.get("area_sqft") or 0)
    except ValueError:
        area = 0
    if job_type not in LABELS or not description or area <= 0:
        return jsonify({"error": "Choose a job type, describe it, and enter the built-up area."}), 400
    audio = save_upload(request.files.get("audio"))
    photo = save_upload(request.files.get("photo"))
    lines = quantity_lines(job_type, area)
    with connect() as conn:
        cur = conn.execute(
            "INSERT INTO jobs (owner_id, job_type, area_sqft, description, audio_path, photo_path, created) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (user["id"], job_type, area, description, audio, photo, now().isoformat()),
        )
        job_id = cur.lastrowid
        conn.executemany(
            "INSERT INTO estimates (job_id, material, unit, quantity) VALUES (?, ?, ?, ?)",
            [(job_id, line["material"], line["unit"], line["quantity"]) for line in lines],
        )
    return jsonify({"id": job_id})


@app.get("/api/jobs")
def list_jobs():
    user, error = require_user()
    if error:
        return error
    with connect() as conn:
        if user["role"] == "vendor":
            rows = []
        else:
            rows = conn.execute(
                "SELECT * FROM jobs WHERE owner_id = ? ORDER BY id DESC",
                (user["id"],),
            ).fetchall()
    return jsonify({"jobs": [dict(row) | {"type_label": LABELS.get(row["job_type"], row["job_type"])} for row in rows]})


@app.get("/api/jobs/<int:job_id>")
def get_job(job_id):
    user, error = require_user()
    if error:
        return error
    with connect() as conn:
        job = conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()
        if not job:
            return jsonify({"error": "Job not found"}), 404
        if user["role"] != "vendor" and job["owner_id"] != user["id"]:
            booked = conn.execute(
                "SELECT 1 FROM bookings WHERE job_id = ? AND person_id = ?",
                (job_id, user["id"]),
            ).fetchone()
            if not booked:
                return jsonify({"error": "This job is not yours"}), 403
        payload = dict(job)
        payload["type_label"] = LABELS.get(job["job_type"], job["job_type"])
        payload["lines"] = match_lines(conn, job)
        payload["updates"] = [dict(r) for r in conn.execute("SELECT * FROM site_updates WHERE job_id = ? ORDER BY id DESC", (job_id,))]
        payload["invoice"] = None
        inv = conn.execute("SELECT * FROM invoices WHERE job_id = ? ORDER BY id DESC LIMIT 1", (job_id,)).fetchone()
        if inv:
            payload["invoice"] = dict(inv)
    return jsonify({"job": payload})


@app.post("/api/catalogue")
def add_catalogue():
    user, error = require_user()
    if error:
        return error
    if user["role"] != "vendor":
        return jsonify({"error": "Only a vendor edits the catalogue"}), 403
    body = request.json or {}
    material = (body.get("material") or "").strip()
    unit = (body.get("unit") or "").strip()
    locality = (body.get("locality") or user.get("locality") or "").strip()
    stock = (body.get("stock") or "").strip()
    try:
        price = float(body.get("price"))
    except (TypeError, ValueError):
        return jsonify({"error": "Enter the price you charge"}), 400
    if not material or not unit or not locality or not stock or price < 0:
        return jsonify({"error": "Material, unit, price, stock, and locality are required"}), 400
    with connect() as conn:
        conn.execute(
            "INSERT INTO catalogue (vendor_id, material, unit, price, stock, locality) VALUES (?, ?, ?, ?, ?, ?)",
            (user["id"], material, unit, price, stock, locality),
        )
    return jsonify({"ok": True})


@app.get("/api/catalogue")
def get_catalogue():
    user, error = require_user()
    if error:
        return error
    with connect() as conn:
        if user["role"] == "vendor":
            rows = conn.execute("SELECT * FROM catalogue WHERE vendor_id = ? ORDER BY id DESC", (user["id"],)).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT c.*, u.name AS vendor_name FROM catalogue c
                JOIN users u ON u.id = c.vendor_id ORDER BY c.id DESC
                """
            ).fetchall()
    return jsonify({"items": [dict(r) for r in rows]})


@app.get("/api/people")
def people():
    user, error = require_user()
    if error:
        return error
    trade = (request.args.get("trade") or "").strip()
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT id, name, phone, role, trade, locality FROM users
            WHERE role IN ('worker', 'contractor') AND id != ? AND name IS NOT NULL
            ORDER BY name
            """,
            (user["id"],),
        ).fetchall()
    people_rows = [dict(r) for r in rows]
    if trade:
        needle = trade.casefold()
        people_rows = [p for p in people_rows if needle in (p["trade"] or "").casefold()]
    return jsonify({"people": people_rows})


@app.post("/api/bookings")
def book():
    user, error = require_user()
    if error:
        return error
    if user["role"] not in ("worker", "contractor"):
        return jsonify({"error": "A worker or contractor makes a booking"}), 403
    body = request.json or {}
    with connect() as conn:
        job = conn.execute("SELECT * FROM jobs WHERE id = ? AND owner_id = ?", (body.get("job_id"), user["id"])).fetchone()
        person = conn.execute("SELECT * FROM users WHERE id = ?", (body.get("person_id"),)).fetchone()
        if not job or not person or person["role"] not in ("worker", "contractor"):
            return jsonify({"error": "Choose your job and a worker or contractor"}), 400
        conn.execute(
            "INSERT INTO bookings (job_id, requester_id, person_id, created) VALUES (?, ?, ?, ?)",
            (job["id"], user["id"], person["id"], now().isoformat()),
        )
    return jsonify({"ok": True})


@app.get("/api/bookings")
def list_bookings():
    user, error = require_user()
    if error:
        return error
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT b.*, j.description, j.job_type, u.name AS person_name, u.trade, u.phone AS person_phone
            FROM bookings b
            JOIN jobs j ON j.id = b.job_id
            JOIN users u ON u.id = b.person_id
            WHERE b.requester_id = ? OR b.person_id = ?
            ORDER BY b.id DESC
            """,
            (user["id"], user["id"]),
        ).fetchall()
    return jsonify({"bookings": [dict(r) | {"type_label": LABELS.get(r["job_type"], r["job_type"])} for r in rows]})


@app.post("/api/site")
def site():
    user, error = require_user()
    if error:
        return error
    if user["role"] != "contractor":
        return jsonify({"error": "The contractor writes the site update. The worker can view it."}), 403
    job_id = request.form.get("job_id")
    note = (request.form.get("note") or "").strip()
    attendance = (request.form.get("attendance") or "").strip()
    if not note:
        return jsonify({"error": "Write what happened on site"}), 400
    photo = save_upload(request.files.get("photo"))
    with connect() as conn:
        job = conn.execute("SELECT id FROM jobs WHERE id = ? AND owner_id = ?", (job_id, user["id"])).fetchone()
        if not job:
            return jsonify({"error": "That job is not yours"}), 404
        conn.execute(
            "INSERT INTO site_updates (job_id, author_id, note, attendance, photo_path, created) VALUES (?, ?, ?, ?, ?, ?)",
            (job_id, user["id"], note, attendance, photo, now().isoformat()),
        )
    return jsonify({"ok": True})


@app.post("/api/invoices")
def create_invoice():
    user, error = require_user()
    if error:
        return error
    if user["role"] != "contractor":
        return jsonify({"error": "The contractor raises the invoice"}), 403
    body = request.json or {}
    note = (body.get("note") or "").strip()
    if not note:
        return jsonify({"error": "Describe what the invoice records"}), 400
    vendor_id = body.get("vendor_id") or None
    with connect() as conn:
        job = conn.execute("SELECT id FROM jobs WHERE id = ? AND owner_id = ?", (body.get("job_id"), user["id"])).fetchone()
        if not job:
            return jsonify({"error": "That job is not yours"}), 404
        conn.execute(
            "INSERT INTO invoices (job_id, author_id, vendor_id, note, status, created) VALUES (?, ?, ?, ?, 'unpaid', ?)",
            (job["id"], user["id"], vendor_id, note, now().isoformat()),
        )
    return jsonify({"ok": True})


@app.post("/api/invoices/<int:invoice_id>/paid")
def mark_paid(invoice_id):
    user, error = require_user()
    if error:
        return error
    if user["role"] != "contractor":
        return jsonify({"error": "The contractor marks the invoice paid"}), 403
    with connect() as conn:
        row = conn.execute("SELECT * FROM invoices WHERE id = ? AND author_id = ?", (invoice_id, user["id"])).fetchone()
        if not row:
            return jsonify({"error": "Invoice not found"}), 404
        conn.execute("UPDATE invoices SET status = 'paid' WHERE id = ?", (invoice_id,))
    return jsonify({"ok": True})


@app.get("/api/invoices")
def list_invoices():
    user, error = require_user()
    if error:
        return error
    with connect() as conn:
        if user["role"] == "vendor":
            rows = conn.execute(
                "SELECT i.*, j.description FROM invoices i JOIN jobs j ON j.id = i.job_id WHERE i.vendor_id = ? ORDER BY i.id DESC",
                (user["id"],),
            ).fetchall()
        elif user["role"] == "worker":
            rows = conn.execute(
                """
                SELECT i.*, j.description FROM invoices i
                JOIN jobs j ON j.id = i.job_id
                JOIN bookings b ON b.job_id = j.id AND b.person_id = ?
                ORDER BY i.id DESC
                """,
                (user["id"],),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT i.*, j.description FROM invoices i JOIN jobs j ON j.id = i.job_id WHERE i.author_id = ? ORDER BY i.id DESC",
                (user["id"],),
            ).fetchall()
    return jsonify({"invoices": [dict(r) for r in rows]})


@app.get("/uploads/<path:name>")
def uploads(name):
    return send_from_directory(UPLOADS, Path(name).name)


init_db()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8080")), debug=False)
