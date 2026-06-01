// Local dev launcher: vendors sb, starts the Python validator (:8090), then runs
// `next dev` pointed at it. One command, two long-running processes, cleaned up
// together. The vendor step is a one-shot and must NOT trigger shutdown.
import { spawn } from "node:child_process";

const PORT = process.env.VALIDATOR_PORT ?? "8090";
const daemons = [];

function shutdown() {
  for (const p of daemons) { try { p.kill(); } catch { /* noop */ } }
}
process.on("SIGINT", () => { shutdown(); process.exit(0); });
process.on("SIGTERM", () => { shutdown(); process.exit(0); });

// long-running; if one dies, take the whole dev session down
function daemon(cmd, args, env = {}) {
  const p = spawn(cmd, args, { stdio: "inherit", env: { ...process.env, ...env } });
  daemons.push(p);
  p.on("exit", (code) => { shutdown(); process.exit(code ?? 0); });
}

// one-shot vendor, then launch the daemons
const vendor = spawn("python3", ["scripts/vendor_sb.py"], { stdio: "inherit" });
vendor.on("exit", () => {
  daemon("python3", ["scripts/validator_server.py", PORT]);
  daemon("npx", ["next", "dev"], { NEXT_PUBLIC_VALIDATOR_BASE: `http://localhost:${PORT}` });
});
