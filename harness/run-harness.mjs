import { spawn } from "node:child_process";
import { promises as fs } from "node:fs";
import path from "node:path";

function parseArgs(argv) {
  const args = {
    evalFile: "harness/evals/smoke.eval.json",
    attempts: 1,
    timeoutMs: 10000,
    reportDir: "harness/reports"
  };

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === "--eval") {
      args.evalFile = argv[i + 1];
      i += 1;
      continue;
    }
    if (token === "--attempts") {
      args.attempts = Number.parseInt(argv[i + 1], 10);
      i += 1;
      continue;
    }
    if (token === "--timeout") {
      args.timeoutMs = Number.parseInt(argv[i + 1], 10);
      i += 1;
      continue;
    }
    if (token === "--report-dir") {
      args.reportDir = argv[i + 1];
      i += 1;
    }
  }

  if (!Number.isInteger(args.attempts) || args.attempts < 1) {
    throw new Error("--attempts must be an integer >= 1");
  }
  if (!Number.isInteger(args.timeoutMs) || args.timeoutMs < 1) {
    throw new Error("--timeout must be an integer >= 1");
  }

  return args;
}

function makeTimestamp(date = new Date()) {
  const pad = (n) => String(n).padStart(2, "0");
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate())
  ].join("") + "-" + [pad(date.getHours()), pad(date.getMinutes()), pad(date.getSeconds())].join("");
}

function ensureArray(value, fieldName) {
  if (value === undefined) return [];
  if (!Array.isArray(value)) {
    throw new Error(`${fieldName} must be an array`);
  }
  return value;
}

function evaluateCaseResult(result, expect = {}) {
  const failures = [];

  const expectedExit = expect.exitCode ?? 0;
  if (result.exitCode !== expectedExit) {
    failures.push(`exitCode expected=${expectedExit} actual=${result.exitCode}`);
  }

  for (const token of ensureArray(expect.stdoutIncludes, "stdoutIncludes")) {
    if (!result.stdout.includes(token)) {
      failures.push(`stdout missing token: ${JSON.stringify(token)}`);
    }
  }
  for (const token of ensureArray(expect.stdoutNotIncludes, "stdoutNotIncludes")) {
    if (result.stdout.includes(token)) {
      failures.push(`stdout unexpectedly contained token: ${JSON.stringify(token)}`);
    }
  }
  for (const token of ensureArray(expect.stderrIncludes, "stderrIncludes")) {
    if (!result.stderr.includes(token)) {
      failures.push(`stderr missing token: ${JSON.stringify(token)}`);
    }
  }
  for (const token of ensureArray(expect.stderrNotIncludes, "stderrNotIncludes")) {
    if (result.stderr.includes(token)) {
      failures.push(`stderr unexpectedly contained token: ${JSON.stringify(token)}`);
    }
  }

  return {
    pass: failures.length === 0,
    failures
  };
}

function runCommand(command, { timeoutMs }) {
  return new Promise((resolve) => {
    const startedAt = Date.now();
    const isArray = Array.isArray(command);
    const spawnArgs = isArray
      ? { cmd: command[0], args: command.slice(1), shell: false }
      : { cmd: command, args: [], shell: true };

    const child = spawn(spawnArgs.cmd, spawnArgs.args, {
      shell: spawnArgs.shell,
      windowsHide: true
    });

    let stdout = "";
    let stderr = "";
    let timedOut = false;

    const timer = setTimeout(() => {
      timedOut = true;
      child.kill("SIGTERM");
    }, timeoutMs);

    child.stdout.on("data", (chunk) => {
      stdout += String(chunk);
    });
    child.stderr.on("data", (chunk) => {
      stderr += String(chunk);
    });

    child.on("close", (exitCode, signal) => {
      clearTimeout(timer);
      resolve({
        command,
        stdout: stdout.trim(),
        stderr: stderr.trim(),
        exitCode: exitCode ?? (timedOut ? 124 : 1),
        signal: signal ?? null,
        timedOut,
        durationMs: Date.now() - startedAt
      });
    });

    child.on("error", (error) => {
      clearTimeout(timer);
      resolve({
        command,
        stdout: stdout.trim(),
        stderr: `${stderr}\n${error.message}`.trim(),
        exitCode: 1,
        signal: null,
        timedOut,
        durationMs: Date.now() - startedAt
      });
    });
  });
}

async function readEvalFile(evalPath) {
  const raw = await fs.readFile(evalPath, "utf8");
  const parsed = JSON.parse(raw);
  if (!parsed || typeof parsed !== "object") {
    throw new Error("Eval file must be a JSON object");
  }
  if (!Array.isArray(parsed.cases) || parsed.cases.length === 0) {
    throw new Error("Eval file must include non-empty 'cases' array");
  }
  return parsed;
}

function validateCase(testCase, index) {
  if (!testCase || typeof testCase !== "object") {
    throw new Error(`Case #${index + 1} must be an object`);
  }
  if (!testCase.id || typeof testCase.id !== "string") {
    throw new Error(`Case #${index + 1} must include string 'id'`);
  }
  if (typeof testCase.command !== "string" && !Array.isArray(testCase.command)) {
    throw new Error(`Case '${testCase.id}' command must be a string or array`);
  }
  if (Array.isArray(testCase.command) && testCase.command.length === 0) {
    throw new Error(`Case '${testCase.id}' command array cannot be empty`);
  }
}

function pct(part, whole) {
  if (whole === 0) return "0.0%";
  return `${((part / whole) * 100).toFixed(1)}%`;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const evalPath = path.resolve(process.cwd(), args.evalFile);
  const reportDir = path.resolve(process.cwd(), args.reportDir);
  const evalDoc = await readEvalFile(evalPath);
  const evalName = evalDoc.name ?? path.basename(evalPath, ".json");

  for (let i = 0; i < evalDoc.cases.length; i += 1) {
    validateCase(evalDoc.cases[i], i);
  }

  const caseResults = [];
  console.log(`Running eval: ${evalName}`);
  console.log(`Cases: ${evalDoc.cases.length}, Attempts per case: ${args.attempts}`);

  for (const testCase of evalDoc.cases) {
    const attempts = [];
    for (let i = 0; i < args.attempts; i += 1) {
      const execution = await runCommand(testCase.command, { timeoutMs: args.timeoutMs });
      const grading = evaluateCaseResult(execution, testCase.expect);
      attempts.push({
        attempt: i + 1,
        ...execution,
        ...grading
      });
    }

    const passAtK = attempts.some((a) => a.pass);
    const passHatK = attempts.every((a) => a.pass);
    caseResults.push({
      id: testCase.id,
      description: testCase.description ?? "",
      passAtK,
      passHatK,
      attempts
    });

    const status = passAtK ? "PASS" : "FAIL";
    console.log(`- ${testCase.id}: ${status} (pass@${args.attempts}=${passAtK}, pass^${args.attempts}=${passHatK})`);
  }

  const passAtKCount = caseResults.filter((r) => r.passAtK).length;
  const passHatKCount = caseResults.filter((r) => r.passHatK).length;

  const report = {
    evalName,
    evalPath,
    attemptsPerCase: args.attempts,
    timeoutMs: args.timeoutMs,
    startedAt: new Date().toISOString(),
    summary: {
      totalCases: caseResults.length,
      passAtKCount,
      passHatKCount,
      passAtKRate: pct(passAtKCount, caseResults.length),
      passHatKRate: pct(passHatKCount, caseResults.length)
    },
    cases: caseResults
  };

  await fs.mkdir(reportDir, { recursive: true });
  const reportPath = path.join(reportDir, `${evalName}-${makeTimestamp()}.json`);
  await fs.writeFile(reportPath, JSON.stringify(report, null, 2), "utf8");

  console.log("");
  console.log(`Summary: pass@${args.attempts} ${passAtKCount}/${caseResults.length} (${report.summary.passAtKRate})`);
  console.log(`Summary: pass^${args.attempts} ${passHatKCount}/${caseResults.length} (${report.summary.passHatKRate})`);
  console.log(`Report: ${path.relative(process.cwd(), reportPath)}`);

  process.exitCode = passAtKCount === caseResults.length ? 0 : 1;
}

main().catch((error) => {
  console.error(`Harness failed: ${error.message}`);
  process.exitCode = 1;
});
