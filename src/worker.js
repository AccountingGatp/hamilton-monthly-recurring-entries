import { loadPyodide } from "/runtime/pyodide.mjs";

let ready;
async function start() {
  const py = await loadPyodide({ indexURL: "/runtime/" });
  const files = [
    "/vendor/et_xmlfile-2.0.0-py3-none-any.whl",
    "/vendor/openpyxl-3.1.5-py2.py3-none-any.whl",
    "/engine/review.py", "/engine/exports.py", "/engine/browser_review.py"
  ];
  await Promise.all(files.map(async url => {
    const response = await fetch(url);
    if (!response.ok) throw new Error("The review engine could not load. Reload the page and retry.");
    py.FS.writeFile(url.split("/").pop(), new Uint8Array(await response.arrayBuffer()));
  }));
  py.runPython(`
import sys
sys.path.extend(["et_xmlfile-2.0.0-py3-none-any.whl", "openpyxl-3.1.5-py2.py3-none-any.whl"])
from browser_review import tabs, review
`);
  return { tabs: py.globals.get("tabs"), review: py.globals.get("review") };
}

self.onmessage = async ({ data }) => {
  const { id, task } = data;
  try {
    const engine = await (ready ??= start());
    if (task === "tabs") {
      self.postMessage({ id, tabs: JSON.parse(engine.tabs(new Uint8Array(data.buffer))) });
      return;
    }
    const result = engine.review(...data.buffers.map(buffer => new Uint8Array(buffer)),
      JSON.stringify(data.names), data.start, data.end, data.tab, data.prior);
    try {
      const [summary, zip] = result.toJs();
      self.postMessage({ id, ...JSON.parse(summary), zip }, [zip.buffer]);
    } finally {
      result.destroy();
    }
  } catch (error) {
    const lines = String(error.message).trim().split("\n");
    const message = lines.at(-1).replace(/^(ValueError|FileNotFoundError):\s*/, "");
    self.postMessage({ id, error: message });
  }
};
