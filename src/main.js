const form = document.querySelector("#review-form");
const status = document.querySelector("#status");
const results = document.querySelector("#results");
const download = document.querySelector("#download");
const runButton = document.querySelector("#run");
const select = form.elements.namedItem("tab");
const exceptions = document.querySelector("#exceptions");
const matchedStatus = "Matched journal entry";
const fileInputs = ["schedule", "chart", "register"].map(name => form.elements.namedItem(name));

let worker;
let sequence = 0;
let displayedRows = [];
let downloadUrl;
let running = false;
const pending = new Map();

function request(body, buffers) {
  if (!worker) {
    worker = new Worker("/worker.js", { type: "module" });
    worker.onmessage = ({ data }) => {
      const task = pending.get(data.id);
      if (!task) return;
      pending.delete(data.id);
      if (data.error) task.reject(new Error(data.error));
      else task.resolve(data);
    };
    worker.onerror = () => {
      for (const task of pending.values()) task.reject(new Error("Could not load the review engine. Reload and try again."));
      pending.clear();
      worker.terminate();
      worker = undefined;
    };
  }
  const id = ++sequence;
  return new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject });
    worker.postMessage({ id, ...body }, buffers);
  });
}

function report(message, error = false) {
  status.textContent = message;
  status.toggleAttribute("data-error", error);
}

function clearResults() {
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = undefined;
  download.removeAttribute("href");
  download.hidden = true;
  results.hidden = true;
  displayedRows = [];
  document.querySelector("#rows").replaceChildren();
}

form.addEventListener("change", () => {
  if (!running) {
    clearResults();
    report("");
  }
});

fileInputs[0].addEventListener("change", async () => {
  const file = fileInputs[0].files[0];
  select.replaceChildren(new Option("Automatic by month", ""));
  if (!file) return;
  try {
    const buffer = await file.arrayBuffer();
    const result = await request({ task: "tabs", buffer }, [buffer]);
    if (fileInputs[0].files[0] === file) {
      for (const name of result.tabs) select.add(new Option(name, name));
    }
  } catch (error) {
    if (!running && fileInputs[0].files[0] === file) report(error.message, true);
  }
});

function renderRows() {
  const fragment = document.createDocumentFragment();
  for (const row of displayedRows) {
    const exception = row[5] !== matchedStatus;
    if (exceptions.checked && !exception) continue;
    const tr = document.createElement("tr");
    tr.toggleAttribute("data-exception", exception);
    for (const value of row) {
      const td = document.createElement("td");
      td.textContent = value ?? "—";
      tr.append(td);
    }
    fragment.append(tr);
  }
  document.querySelector("#rows").replaceChildren(fragment);
}

exceptions.addEventListener("change", renderRows);
form.addEventListener("submit", async event => {
  event.preventDefault();
  if (running) return;
  const files = fileInputs.map(input => input.files[0]);
  if (files.some(file => !file || !file.name.toLowerCase().endsWith(".xlsx"))) {
    report("Select three .xlsx workbooks.", true);
    return;
  }
  const month = form.elements.namedItem("month").value;
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
    report("Enter the review month as YYYY-MM.", true);
    return;
  }
  const tab = select.value;
  const prior = form.elements.namedItem("prior").checked;
  clearResults();
  running = true;
  const controls = [...form.querySelectorAll("input, select, button")];
  controls.forEach(control => { control.disabled = true; });
  runButton.textContent = "Reviewing…";
  report("Reading workbooks and comparing journal activity…");
  try {
    const buffers = await Promise.all(files.map(file => file.arrayBuffer()));
    const result = await request({ task: "review", buffers, names: files.map(file => file.name),
      month, tab, prior }, buffers);
    const { details, rows, zip } = result;
    displayedRows = rows.sort((a, b) => Number(a[5] === matchedStatus) - Number(b[5] === matchedStatus));
    const matched = rows.filter(row => row[5] === matchedStatus).length;
    results.hidden = false;
    exceptions.checked = false;
    document.querySelector("#source").textContent =
      `Schedule tab ${details["Schedule tab"]} · ${details["Schedule source"]} · ${details["Register coverage"]}`;
    report(`${month} · ${matched} matched · ${rows.length - matched} to review · ${details["Draft entries"]} drafts` +
      (details["Full review month in register"] === "Yes" ? "" : " · Partial or unverified register coverage"));
    renderRows();
    downloadUrl = URL.createObjectURL(new Blob([zip], { type: "application/zip" }));
    download.href = downloadUrl;
    download.download = `Recurring_Review_${month}.zip`;
    download.hidden = false;
  } catch (error) {
    report(error.message, true);
  } finally {
    controls.forEach(control => { control.disabled = false; });
    runButton.textContent = "Run review";
    running = false;
  }
});
