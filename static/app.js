const tabs = document.querySelectorAll(".tab");
const panels = document.querySelectorAll(".panel");
const log = document.querySelector("#log");
const jobStatus = document.querySelector("#job-status");
const mergeFilesInput = document.querySelector("#merge-files");
const mergeFilesLabel = document.querySelector("#merge-files-label");
const printerSelect = document.querySelector("#printer");
const singlePrinterSelect = document.querySelector("#single-printer");
const geostudioPrinterSelect = document.querySelector("#geostudio-printer");
const printerError = document.querySelector("#printer-error");

function setStatus(text) {
  jobStatus.textContent = text;
}

function writeLog(text) {
  log.textContent = text || "";
  log.scrollTop = log.scrollHeight;
}

function appendLog(lines) {
  if (!lines || !lines.length) return;
  const prefix = log.textContent && log.textContent !== "Ready." ? "\n" : "";
  log.textContent += prefix + lines.join("\n");
  log.scrollTop = log.scrollHeight;
}

function isPdfPrinter(printer) {
  const value = (printer || "").toLowerCase();
  return value.includes("pdf");
}

function formDataFrom(form) {
  return new FormData(form);
}

async function postForm(url, formData) {
  const response = await fetch(url, { method: "POST", body: formData });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || "Request failed.");
  return payload;
}

function watchJob(jobId) {
  writeLog("");
  setStatus("Running");
  const events = new EventSource(`/api/jobs/${jobId}/events`);
  events.onmessage = (event) => {
    const payload = JSON.parse(event.data);
    appendLog(payload.lines);
    setStatus(payload.status);
    if (payload.status === "done" || payload.status === "failed") {
      events.close();
      appendLog([`Finished: ${payload.status}`]);
    }
  };
  events.onerror = () => {
    events.close();
    setStatus("Disconnected");
  };
}

async function loadPrinters() {
  try {
    const response = await fetch("/api/printers");
    const payload = await response.json();
    printerSelect.innerHTML = "";
    singlePrinterSelect.innerHTML = "";
    geostudioPrinterSelect.innerHTML = "";
    if (payload.printers && payload.printers.length) {
      payload.printers.forEach((printer) => {
        const option = document.createElement("option");
        option.value = printer;
        option.textContent = printer;
        printerSelect.appendChild(option);
        singlePrinterSelect.appendChild(option.cloneNode(true));
        geostudioPrinterSelect.appendChild(option.cloneNode(true));
      });
    } else {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "No printers found";
      printerSelect.appendChild(option);
      singlePrinterSelect.appendChild(option.cloneNode(true));
      geostudioPrinterSelect.appendChild(option.cloneNode(true));
    }
    printerError.textContent = payload.error || "";
    updatePrinterMode(document.querySelector("#print-form"), printerSelect.value);
    updatePrinterMode(document.querySelector("#print-one-form"), singlePrinterSelect.value);
  } catch (error) {
    printerSelect.innerHTML = '<option value="">Could not load printers</option>';
    singlePrinterSelect.innerHTML = '<option value="">Could not load printers</option>';
    geostudioPrinterSelect.innerHTML = '<option value="">Could not load printers</option>';
    printerError.textContent = error.message;
  }
}

function renderFilePreview(containerId, countId, excludeId, files) {
  const container = document.querySelector(containerId);
  const count = document.querySelector(countId);
  const excludeInput = document.querySelector(excludeId);
  const summary = document.querySelector(`${excludeId}-summary`);
  container.innerHTML = "";
  excludeInput.value = "";
  if (summary) summary.textContent = "No files excluded";
  count.textContent = files.length ? `${files.length} file(s)` : "No matching files";

  files.forEach((file) => {
    const row = document.createElement("label");
    row.className = "file-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.dataset.index = file.index;
    checkbox.addEventListener("change", () => {
      const excluded = Array.from(container.querySelectorAll("input:checked"))
        .map((item) => item.dataset.index)
        .join(",");
      excludeInput.value = excluded;
      if (summary) {
        summary.textContent = excluded ? `Excluded file number(s): ${excluded}` : "No files excluded";
      }
    });

    const index = document.createElement("span");
    index.className = "file-index";
    index.textContent = file.index;

    const name = document.createElement("span");
    name.className = "file-name";
    name.textContent = file.label;

    row.append(checkbox, index, name);
    container.appendChild(row);
  });
}

function updatePrinterMode(form, printer) {
  const pdfMode = isPdfPrinter(printer);
  form.querySelectorAll(".paper-option").forEach((element) => {
    element.classList.toggle("hidden", pdfMode);
  });
  form.querySelectorAll(".pdf-output").forEach((element) => {
    element.classList.toggle("hidden", !pdfMode);
  });
}

function wirePrinterMode(select, formSelector) {
  const form = document.querySelector(formSelector);
  const update = () => updatePrinterMode(form, select.value);
  select.addEventListener("change", update);
  update();
}

tabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    tabs.forEach((item) => item.classList.remove("active"));
    panels.forEach((panel) => panel.classList.remove("active"));
    tab.classList.add("active");
    document.querySelector(`#${tab.dataset.tab}`).classList.add("active");
  });
});

document.querySelectorAll("[data-pick-folder]").forEach((button) => {
  button.addEventListener("click", async () => {
    setStatus("Opening picker");
    const response = await fetch("/api/pick-folder", { method: "POST" });
    const payload = await response.json();
    if (payload.path) {
      document.querySelector(`#${button.dataset.pickFolder}`).value = payload.path;
    }
    setStatus("Idle");
  });
});

document.querySelector("#pick-merge-files").addEventListener("click", async () => {
  setStatus("Opening picker");
  const response = await fetch("/api/pick-merge-files", { method: "POST" });
  const payload = await response.json();
  const files = payload.files || [];
  mergeFilesInput.value = JSON.stringify(files);
  mergeFilesLabel.value = files.length ? `${files.length} file(s) selected` : "";
  setStatus("Idle");
});

document.querySelector("#pick-print-file").addEventListener("click", async () => {
  setStatus("Opening picker");
  const response = await fetch("/api/pick-print-file", { method: "POST" });
  const payload = await response.json();
  if (payload.path) {
    document.querySelector("#single-file").value = payload.path;
  }
  setStatus("Idle");
});

document.querySelector("#pick-split-pdf").addEventListener("click", async () => {
  setStatus("Opening picker");
  const response = await fetch("/api/pick-pdf-file", { method: "POST" });
  const payload = await response.json();
  if (payload.path) {
    document.querySelector("#split-input-pdf").value = payload.path;
  }
  setStatus("Idle");
});

document.querySelector("#pick-geostudio-file").addEventListener("click", async () => {
  setStatus("Opening picker");
  const response = await fetch("/api/pick-geostudio-file", { method: "POST" });
  const payload = await response.json();
  if (payload.path) {
    document.querySelector("#geostudio-project").value = payload.path;
  }
  setStatus("Idle");
});

document.querySelector("#load-geostudio-analyses").addEventListener("click", async () => {
  try {
    const payload = await postForm("/api/geostudio-analyses", formDataFrom(document.querySelector("#geostudio-form")));
    const container = document.querySelector("#geostudio-analysis-preview");
    const count = document.querySelector("#geostudio-analysis-count");
    const input = document.querySelector("#geostudio-analyses");
    container.innerHTML = "";
    input.value = "";
    const analyses = payload.analyses || [];
    count.textContent = analyses.length ? `${analyses.length} analysis/analyses` : "No analyses found";

    analyses.forEach((analysis, idx) => {
      const row = document.createElement("label");
      row.className = "file-row";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.dataset.name = analysis;
      checkbox.addEventListener("change", () => {
        input.value = Array.from(container.querySelectorAll("input:checked"))
          .map((item) => item.dataset.name)
          .join(", ");
      });

      const index = document.createElement("span");
      index.className = "file-index";
      index.textContent = idx + 1;

      const name = document.createElement("span");
      name.className = "file-name";
      name.textContent = analysis;

      row.append(checkbox, index, name);
      container.appendChild(row);
    });
    setStatus("Analyses loaded");
  } catch (error) {
    writeLog(`ERROR: ${error.message}`);
    setStatus("Error");
  }
});

document.querySelector("#geostudio-print-pdf").addEventListener("change", (event) => {
  const enabled = event.currentTarget.checked;
  document.querySelectorAll(".geostudio-print-option").forEach((element) => {
    element.classList.toggle("hidden", !enabled);
  });
  if (enabled) {
    document.querySelector("#geostudio-form input[name='pdf']").checked = true;
  }
});

document.querySelector("#preview-print-files").addEventListener("click", async () => {
  try {
    const payload = await postForm("/api/preview-print", formDataFrom(document.querySelector("#print-form")));
    renderFilePreview("#print-file-preview", "#print-preview-count", "#exclude", payload.files || []);
    setStatus("Preview ready");
  } catch (error) {
    writeLog(`ERROR: ${error.message}`);
    setStatus("Error");
  }
});

document.querySelector("#preview-merge-files").addEventListener("click", async () => {
  try {
    const payload = await postForm("/api/preview-merge", formDataFrom(document.querySelector("#merge-form")));
    renderFilePreview("#merge-file-preview", "#merge-preview-count", "#merge-exclude", payload.files || []);
    setStatus("Preview ready");
  } catch (error) {
    writeLog(`ERROR: ${error.message}`);
    setStatus("Error");
  }
});

document.querySelector("#print-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const payload = await postForm("/api/print", formDataFrom(event.currentTarget));
    watchJob(payload.job_id);
  } catch (error) {
    writeLog(`ERROR: ${error.message}`);
    setStatus("Error");
  }
});

document.querySelector("#print-one-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const payload = await postForm("/api/print-one", formDataFrom(event.currentTarget));
    watchJob(payload.job_id);
  } catch (error) {
    writeLog(`ERROR: ${error.message}`);
    setStatus("Error");
  }
});

document.querySelector("#merge-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const payload = await postForm("/api/merge", formDataFrom(event.currentTarget));
    watchJob(payload.job_id);
  } catch (error) {
    writeLog(`ERROR: ${error.message}`);
    setStatus("Error");
  }
});

document.querySelector("#split-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const payload = await postForm("/api/split", formDataFrom(event.currentTarget));
    watchJob(payload.job_id);
  } catch (error) {
    writeLog(`ERROR: ${error.message}`);
    setStatus("Error");
  }
});

document.querySelector("#geostudio-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const payload = await postForm("/api/geostudio", formDataFrom(event.currentTarget));
    watchJob(payload.job_id);
  } catch (error) {
    writeLog(`ERROR: ${error.message}`);
    setStatus("Error");
  }
});

wirePrinterMode(printerSelect, "#print-form");
wirePrinterMode(singlePrinterSelect, "#print-one-form");
loadPrinters();
