// Spawns and manages the local Python FastAPI backend (server.py).
const path = require('path');
const { spawn } = require('child_process');

const REPO_ROOT = path.join(__dirname, '..', '..');
const PORT = 8765;
const HEALTH_URL = `http://127.0.0.1:${PORT}/api/health`;

function venvPythonPath() {
  if (process.platform === 'win32') {
    return path.join(REPO_ROOT, '.venv', 'Scripts', 'python.exe');
  }
  return path.join(REPO_ROOT, '.venv', 'bin', 'python3');
}

let pythonProcess = null;

function start() {
  if (process.env.SKIP_PYTHON_SPAWN) {
    console.log('[pythonBackend] SKIP_PYTHON_SPAWN set — expecting server.py already running');
    return;
  }
  pythonProcess = spawn(venvPythonPath(), ['server.py'], { cwd: REPO_ROOT });
  pythonProcess.stdout.on('data', (data) => process.stdout.write(`[python] ${data}`));
  pythonProcess.stderr.on('data', (data) => process.stderr.write(`[python] ${data}`));
  pythonProcess.on('exit', (code) => console.log(`[pythonBackend] process exited with code ${code}`));
}

async function waitUntilHealthy(timeoutMs = 15000, intervalMs = 300) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(HEALTH_URL);
      if (res.ok) return true;
    } catch {
      // backend not up yet — keep polling
    }
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
  throw new Error(`Python backend did not become healthy within ${timeoutMs}ms`);
}

function stop() {
  if (!pythonProcess) return;
  pythonProcess.kill('SIGTERM');
  const proc = pythonProcess;
  setTimeout(() => {
    if (!proc.killed) proc.kill('SIGKILL');
  }, 3000);
  pythonProcess = null;
}

module.exports = { start, stop, waitUntilHealthy, PORT };
