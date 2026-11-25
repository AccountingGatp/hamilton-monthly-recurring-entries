import { cp, mkdir, rm } from "node:fs/promises";

await rm("dist", { recursive: true, force: true });
await cp("src", "dist", { recursive: true });
await cp("engine", "dist/engine", { recursive: true, filter: name => !name.includes("__pycache__") });
await cp("vendor", "dist/vendor", { recursive: true });
await mkdir("dist/runtime");
const runtime = ["pyodide.mjs", "pyodide.asm.js", "pyodide.asm.wasm", "python_stdlib.zip", "pyodide-lock.json"];
await Promise.all(runtime.map(file => cp(`node_modules/pyodide/${file}`, `dist/runtime/${file}`)));
