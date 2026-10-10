const state = { phone: "", user: null, lang: "en" };
const text = {
  en: {
    tagline: "Delhi pilot · website on this computer",
    banner: "This is the website channel only. The one-time password appears on this page. No text message is sent. WhatsApp, the Android app, UPI, GST, and maps are not connected.",
    signin: "Sign in with your mobile number",
    phone: "Mobile number",
    send: "Send code",
    signout: "Sign out",
    lang: "हिंदी",
  },
  hi: {
    tagline: "दिल्ली पायलट · इस कंप्यूटर पर वेबसाइट",
    banner: "यह केवल वेबसाइट है। कोड इसी पेज पर दिखेगा। संदेश नहीं जाएगा। व्हाट्सऐप, एंड्रॉइड ऐप, यूपीआई, जीएसटी और नक्शा जुड़े नहीं हैं।",
    signin: "मोबाइल नंबर से प्रवेश",
    phone: "मोबाइल नंबर",
    send: "कोड भेजें",
    signout: "बाहर जाएँ",
    lang: "English",
  },
};

function t(key) { return text[state.lang][key] || key; }

function show(id) {
  for (const view of ["view-phone", "view-code", "view-profile", "view-home"]) {
    document.getElementById(view).classList.toggle("hidden", view !== id);
  }
}

function say(message) {
  document.getElementById("error").textContent = message || "";
}

async function api(url, options = {}) {
  const response = await fetch(url, {
    headers: options.body instanceof FormData ? undefined : { "Content-Type": "application/json" },
    ...options,
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || "Something went wrong");
  return data;
}

function applyLang() {
  document.getElementById("tagline").textContent = t("tagline");
  document.getElementById("banner").textContent = t("banner");
  document.getElementById("h-phone").textContent = t("signin");
  document.getElementById("l-phone").textContent = t("phone");
  document.getElementById("send-code").textContent = t("send");
  document.getElementById("signout").textContent = t("signout");
  document.getElementById("lang").textContent = t("lang");
}

function roleLabel(role) {
  return { worker: "Construction worker", contractor: "Contractor", vendor: "Vendor" }[role] || "";
}

async function enter() {
  const me = await api("/api/me");
  state.user = me.user;
  document.getElementById("signout").classList.toggle("hidden", !state.user);
  if (!state.user) return show("view-phone");
  if (!state.user.role) return show("view-profile");
  show("view-home");
  renderHome();
}

function renderHome() {
  const user = state.user;
  document.getElementById("hello").textContent = user.name;
  document.getElementById("who").textContent = `${roleLabel(user.role)} · ${user.trade} · ${user.locality}`;
  const actions = document.getElementById("home-actions");
  const buttons = [];
  if (user.role === "vendor") {
    buttons.push(["Catalogue", () => panelCatalogue()], ["Invoices", () => panelInvoices()]);
  } else {
    buttons.push(["New job", () => panelJob()], ["My jobs", () => panelJobs()], ["Find people", () => panelPeople()], ["Bookings", () => panelBookings()]);
    if (user.role === "contractor") buttons.push(["Invoices", () => panelInvoices()]);
  }
  actions.innerHTML = "";
  for (const [label, fn] of buttons) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "secondary";
    button.textContent = label;
    button.onclick = fn;
    actions.appendChild(button);
  }
  if (user.role === "vendor") panelCatalogue();
  else panelJobs();
}

function panel(html) {
  document.getElementById("panel").innerHTML = html;
}

async function panelJob() {
  panel(`
    <form class="card" id="job-form">
      <h2>New job</h2>
      <p class="muted">Quantities use the demo rules for this pilot. They are not a published rate schedule. Rupees appear only after a vendor types a price.</p>
      <label>Job type</label>
      <select name="job_type">
        <option value="JOB-NEW">New building</option>
        <option value="JOB-EXT">Extension or extra floor</option>
        <option value="JOB-FIT">Interior fit-out</option>
      </select>
      <label>Built-up area (sq ft)</label>
      <input name="area_sqft" type="number" min="1" step="1" required>
      <label>Description</label>
      <textarea name="description" required placeholder="Two rooms, plaster and floor, Laxmi Nagar"></textarea>
      <label>Photo</label>
      <input name="photo" type="file" accept="image/*">
      <label>Audio</label>
      <input name="audio" type="file" accept="audio/*">
      <button type="submit">Make quantity list</button>
    </form>`);
  document.getElementById("job-form").onsubmit = async (event) => {
    event.preventDefault();
    say("");
    try {
      const created = await api("/api/jobs", { method: "POST", body: new FormData(event.target) });
      await openJob(created.id);
    } catch (error) { say(error.message); }
  };
}

async function panelJobs() {
  const data = await api("/api/jobs");
  const rows = data.jobs.map((job) => `
    <tr class="job" data-id="${job.id}">
      <td><span class="pill">${job.type_label}</span></td>
      <td>${job.area_sqft} sq ft</td>
      <td>${escapeHtml(job.description)}</td>
    </tr>`).join("");
  panel(`<div class="card"><h2>My jobs</h2>${rows ? `<table><tr><th>Type</th><th>Area</th><th>Description</th></tr>${rows}</table>` : `<p class="muted">No jobs yet.</p>`}</div>`);
  document.querySelectorAll(".job").forEach((row) => { row.onclick = () => openJob(row.dataset.id); });
}

async function openJob(id) {
  const data = await api(`/api/jobs/${id}`);
  const job = data.job;
  const lines = job.lines.map((line) => {
    const offer = line.offers[0];
    const vendor = offer ? `${escapeHtml(offer.vendor_name)} · ${offer.locality} · stock ${escapeHtml(offer.stock)}` : "No vendor in this locality has typed this material";
    const rate = line.rate == null ? "RATE NEEDED" : `₹${line.rate}`;
    const amount = line.amount == null ? "" : `₹${line.amount}`;
    return `<tr><td>${escapeHtml(line.material)}</td><td>${line.quantity} ${escapeHtml(line.unit)}</td><td>${rate}</td><td>${amount}</td><td>${vendor}<br><span class="muted">${escapeHtml(line.source)}</span></td></tr>`;
  }).join("");
  const contractor = state.user.role === "contractor";
  panel(`
    <div class="card">
      <h2>${job.type_label}</h2>
      <p>${job.area_sqft} sq ft · ${escapeHtml(job.description)}</p>
      ${job.photo_path ? `<p><img alt="Job photo" src="/uploads/${job.photo_path}" style="max-width:240px;border-radius:8px"></p>` : ""}
      <table>
        <tr><th>Material</th><th>Quantity</th><th>Rate</th><th>Amount</th><th>Vendor</th></tr>
        ${lines}
      </table>
    </div>
    ${contractor ? siteForm(job.id) + invoiceForm(job.id, job.invoice, vendorsOn(job)) : ""}
    <div class="card"><h2>Site record</h2>${job.updates.length ? job.updates.map((item) => `<p>${escapeHtml(item.note)} ${item.attendance ? "· attendance " + escapeHtml(item.attendance) : ""}</p>`).join("") : `<p class="muted">No site update yet.</p>`}</div>
  `);
  const site = document.getElementById("site-form");
  if (site) site.onsubmit = async (event) => {
    event.preventDefault();
    try {
      await api("/api/site", { method: "POST", body: new FormData(event.target) });
      await openJob(id);
    } catch (error) { say(error.message); }
  };
  const invoice = document.getElementById("invoice-form");
  if (invoice) invoice.onsubmit = async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    try {
      await api("/api/invoices", { method: "POST", body: JSON.stringify({ job_id: id, note: form.get("note"), vendor_id: form.get("vendor_id") || null }) });
      await openJob(id);
    } catch (error) { say(error.message); }
  };
  const paid = document.getElementById("mark-paid");
  if (paid) paid.onclick = async () => {
    await api(`/api/invoices/${job.invoice.id}/paid`, { method: "POST", body: "{}" });
    await openJob(id);
  };
}

function siteForm(jobId) {
  return `<form class="card" id="site-form"><h2>Site update</h2><input type="hidden" name="job_id" value="${jobId}"><label>What happened</label><textarea name="note" required></textarea><label>Attendance</label><input name="attendance" placeholder="2 masons present"><button type="submit">Save update</button></form>`;
}

function vendorsOn(job) {
  const map = new Map();
  for (const line of job.lines) {
    for (const offer of line.offers) map.set(String(offer.vendor_id), offer.vendor_name);
  }
  return [...map.entries()];
}

function invoiceForm(jobId, invoice, vendors) {
  if (invoice) {
    return `<div class="card"><h2>Invoice</h2><p>${escapeHtml(invoice.note)}</p><p class="pill">${invoice.status}</p>${invoice.status === "unpaid" ? `<button type="button" id="mark-paid">Mark paid</button>` : ""}</div>`;
  }
  const options = [`<option value="">No vendor named</option>`].concat(vendors.map(([id, name]) => `<option value="${id}">${escapeHtml(name)}</option>`)).join("");
  return `<form class="card" id="invoice-form"><h2>Internal invoice</h2><p class="muted">This records unpaid or paid. It does not collect money and it is not a GST invoice.</p><label>What it records</label><textarea name="note" required></textarea><label>Vendor named on this record</label><select name="vendor_id">${options}</select><button type="submit">Raise invoice</button></form>`;
}

async function panelCatalogue() {
  const data = await api("/api/catalogue");
  const rows = data.items.map((item) => `<tr><td>${escapeHtml(item.material)}</td><td>₹${item.price} / ${escapeHtml(item.unit)}</td><td>${escapeHtml(item.stock)}</td><td>${escapeHtml(item.locality)}</td></tr>`).join("");
  const form = state.user.role === "vendor" ? `
    <form id="cat-form">
      <label>Material</label><input name="material" required placeholder="Cement">
      <label>Unit</label><input name="unit" required placeholder="bags">
      <label>Your price (rupees)</label><input name="price" type="number" min="0" step="0.01" required>
      <label>Stock</label><input name="stock" required placeholder="40 bags">
      <label>Locality you serve</label><input name="locality" value="${escapeHtml(state.user.locality)}" required>
      <button type="submit">Add to catalogue</button>
    </form>` : "";
  panel(`<div class="card"><h2>Catalogue</h2>${form}${rows ? `<table><tr><th>Material</th><th>Price</th><th>Stock</th><th>Locality</th></tr>${rows}</table>` : `<p class="muted">Nothing typed in yet.</p>`}</div>`);
  const cat = document.getElementById("cat-form");
  if (cat) cat.onsubmit = async (event) => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const body = Object.fromEntries(formData.entries());
    try {
      await api("/api/catalogue", { method: "POST", body: JSON.stringify(body) });
      await panelCatalogue();
    } catch (error) { say(error.message); }
  };
}

async function panelPeople() {
  panel(`<div class="card"><h2>Find people by trade</h2><label>Trade</label><input id="trade-q" placeholder="mason"><button type="button" id="trade-go">Search</button><div id="people"></div></div>`);
  document.getElementById("trade-go").onclick = loadPeople;
}

async function loadPeople() {
  const trade = document.getElementById("trade-q").value;
  const jobs = await api("/api/jobs");
  const data = await api(`/api/people?trade=${encodeURIComponent(trade)}`);
  const options = jobs.jobs.map((job) => `<option value="${job.id}">${job.type_label}: ${escapeHtml(job.description)}</option>`).join("");
  document.getElementById("people").innerHTML = data.people.length ? data.people.map((person) => `
    <div class="card">
      <strong>${escapeHtml(person.name)}</strong>
      <p class="muted">${roleLabel(person.role)} · ${escapeHtml(person.trade)} · ${escapeHtml(person.locality)} · ${person.phone}</p>
      <label>Book onto</label>
      <select id="job-for-${person.id}">${options || `<option value="">Open a job first</option>`}</select>
      <button type="button" data-person="${person.id}">Book</button>
    </div>`).join("") : `<p class="muted">No one with that trade has signed up.</p>`;
  document.querySelectorAll("[data-person]").forEach((button) => {
    button.onclick = async () => {
      try {
        await api("/api/bookings", { method: "POST", body: JSON.stringify({ job_id: document.getElementById(`job-for-${button.dataset.person}`).value, person_id: Number(button.dataset.person) }) });
        say("");
        await panelBookings();
      } catch (error) { say(error.message); }
    };
  });
}

async function panelBookings() {
  const data = await api("/api/bookings");
  panel(`<div class="card"><h2>Bookings</h2>${data.bookings.length ? data.bookings.map((item) => `<p><strong>${escapeHtml(item.person_name)}</strong> · ${escapeHtml(item.trade)} · ${item.type_label}<br><span class="muted">${escapeHtml(item.description)}</span></p>`).join("") : `<p class="muted">No bookings yet.</p>`}</div>`);
}

async function panelInvoices() {
  const data = await api("/api/invoices");
  panel(`<div class="card"><h2>Invoices</h2>${data.invoices.length ? data.invoices.map((item) => `<p>${escapeHtml(item.note)} · <span class="pill">${item.status}</span><br><span class="muted">${escapeHtml(item.description || "")}</span></p>`).join("") : `<p class="muted">No invoices yet.</p>`}</div>`);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
}

document.getElementById("lang").onclick = () => {
  state.lang = state.lang === "en" ? "hi" : "en";
  applyLang();
};
document.getElementById("send-code").onclick = async () => {
  say("");
  try {
    const data = await api("/api/login/start", { method: "POST", body: JSON.stringify({ phone: document.getElementById("phone").value }) });
    state.phone = data.phone;
    document.getElementById("demo-code").textContent = data.demo_code;
    show("view-code");
  } catch (error) { say(error.message); }
};
document.getElementById("verify").onclick = async () => {
  say("");
  try {
    const data = await api("/api/login/verify", { method: "POST", body: JSON.stringify({ phone: state.phone, code: document.getElementById("code").value }) });
    state.user = data.user;
    await enter();
  } catch (error) { say(error.message); }
};
document.getElementById("save-profile").onclick = async () => {
  say("");
  try {
    await api("/api/profile", { method: "POST", body: JSON.stringify({
      name: document.getElementById("name").value,
      role: document.getElementById("role").value,
      trade: document.getElementById("trade").value,
      locality: document.getElementById("locality").value,
    }) });
    await enter();
  } catch (error) { say(error.message); }
};
document.getElementById("signout").onclick = async () => {
  await api("/api/logout", { method: "POST", body: "{}" });
  state.user = null;
  await enter();
};

applyLang();
enter().catch((error) => say(error.message));
