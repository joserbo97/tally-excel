// Offline tests for the agent loop (empty turns, Stop), with a scripted stand-in for Ollama.
//   node scripts/test-agent-loop.mjs
import { createWorkbook } from "./lib/mock-excel.mjs";

const { Excel, api } = createWorkbook({ sheets: { Sheet1: [["a", 1]] } });
globalThis.Excel = Excel;
globalThis.OLLAMA_API_BASE = "http://stub/api";

let script = [];
const requests = [];
let streamedChunks = 0;
globalThis.fetch = async (url, init) => {
  requests.push(JSON.parse(init.body));
  const msg = script.shift() || { content: "out of script" };
  if (msg.thinkingChunks) {
    // Stream thinking 1,000 chars at a time, stopping when the request is aborted (as Ollama does).
    let i = 0;
    const enc = new TextEncoder();
    const body = new ReadableStream({
      pull(ctrl) {
        if (init.signal?.aborted) return ctrl.error(new DOMException("Aborted", "AbortError"));
        if (i++ < msg.thinkingChunks) { streamedChunks++; return ctrl.enqueue(enc.encode(JSON.stringify({ message: { role: "assistant", thinking: "x".repeat(1000) } }) + "\n")); }
        ctrl.enqueue(enc.encode(JSON.stringify({ message: { role: "assistant", content: msg.content || "" }, done: true, eval_count: 5 }) + "\n"));
        ctrl.close();
      },
    });
    return new Response(body, { status: 200 });
  }
  const line = JSON.stringify({ message: { role: "assistant", ...msg }, done: true, done_reason: "stop", eval_count: 5 }) + "\n";
  return new Response(line, { status: 200 });
};

const { runTurn } = await import("../public/js/agent.js");
const noop = () => {};
function ui(empties) {
  return {
    step: () => ({ thinking: noop, content: noop, done: noop, empty: (i) => empties.push(i), toolStart: () => ({ approve: async () => true, end: noop }) }),
    stats: noop,
  };
}
async function run(steps) {
  script = steps;
  requests.length = 0;
  const empties = [];
  const history = [{ role: "user", content: "task" }];
  const answer = await runTurn({ history, settings: { model: "m", think: "low", numCtx: 1024, temperature: 0, maxSteps: 10, approval: "none", web: false }, signal: new AbortController().signal, ui: ui(empties), undo: [] });
  return { answer, empties, history };
}

let bad = 0;
const ok = (cond, name) => { console.log(`${cond ? "✓" : "✗"} ${name}`); if (!cond) bad++; };

// 1. Empty turn, then a real answer: the agent nudges and returns the answer.
let r = await run([{ content: "", thinking: "hmm" }, { content: "Done." }]);
ok(r.answer === "Done.", "empty turn is nudged, then the real answer is returned");
ok(r.empties.length === 1 && r.empties[0].thinkingTail === "hmm", "empty turn is reported with the thinking tail");
ok(requests[1].messages.at(-1).role === "user" && /no reply and no tool call/.test(requests[1].messages.at(-1).content), "the nudge is sent to the model");

// 2. Tool call written inside the thinking: it is recovered and executed.
r = await run([
  { content: "", thinking: 'I will write. <tool_call>{"name":"write_range","arguments":{"address":"Sheet1!C1","values":[[42]]}}</tool_call>' },
  { content: "Wrote 42." },
]);
ok(api.value("Sheet1!C1") === 42, "a tool call inside the thinking is recovered and run");
ok(r.empties.length === 0, "no nudge needed when the call was recovered");

// 3. The model stays empty: at most 2 nudges, then the (empty) answer is returned instead of looping.
r = await run([{ content: "" }, { content: "" }, { content: "" }, { content: "never reached" }]);
ok(r.empties.length === 2 && r.answer === "", "gives up after 2 nudges");

// 4. Normal answer: untouched.
r = await run([{ content: "Hello" }]);
ok(r.answer === "Hello" && r.empties.length === 0 && requests.length === 1, "a normal answer is unaffected");

// 5. Long thinking is never cut off: the whole stream is read and the step still thinks.
streamedChunks = 0;
r = await run([{ thinkingChunks: 200, content: "Acted." }]);
ok(r.answer === "Acted." && streamedChunks === 200 && requests.length === 1 && requests[0].think === "low", `long thinking runs to the end (${streamedChunks}K chars read)`);

// 6. The user pressing Stop mid-thinking aborts the turn.
{
  script = [{ thinkingChunks: 200 }];
  const ctl = new AbortController();
  const stopUi = { step: () => ({ ...ui([]).step(), thinking: () => ctl.abort() }), stats: noop };
  let err = null;
  try { await runTurn({ history: [{ role: "user", content: "t" }], settings: { model: "m", think: "low", numCtx: 1024, temperature: 0, maxSteps: 3, approval: "none", web: false }, signal: ctl.signal, ui: stopUi, undo: [] }); } catch (e) { err = e; }
  ok(err?.name === "AbortError", "Stop by the user still aborts the turn");
}

console.log(bad ? `\n${bad} failed` : "\nall passed");
process.exit(bad ? 1 : 0);
