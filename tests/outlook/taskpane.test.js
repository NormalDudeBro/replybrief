// ReplyBrief Outlook task pane tests.
//
//   bun test tests/outlook
//
// No dependencies. The real taskpane.js runs against a fake Office runtime and a minimal DOM built
// here, so the behaviour the .NET tests cannot reach is covered without a browser or a mailbox.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { beforeEach, describe, expect, test } from "bun:test";

const SOURCE = fileURLToPath(new URL("../../integrations/outlook/taskpane.js", import.meta.url));
const SOURCE_TEXT = readFileSync(SOURCE, "utf8");

const ACCEPTED = { ok: true, status: 202, text: async () => '{"status":"accepted"}' };

function makeElement(id) {
  return {
    id,
    value: "",
    hidden: false,
    disabled: false,
    textContent: "",
    className: "",
    listeners: {},
    addEventListener(name, handler) {
      (this.listeners[name] ||= []).push(handler);
    },
    dispatch(name) {
      (this.listeners[name] || []).forEach((handler) => handler());
    },
    focus() {},
  };
}

const SENDER = { displayName: "Dana Example", emailAddress: "dana@example.com" };

/**
 * Installs the globals the task pane touches, loads the real script and returns handles for driving
 * it. `setItem` installs the message the reading pane currently has selected and returns the array
 * that body-read callbacks will land in.
 */
function boot({ fetchResponse = ACCEPTED, gate = false } = {}) {
  const reads = [];
  const elements = new Map(
    ["pairing-code", "forget-button", "preview-subject", "preview-from", "capture-button", "status"].map((id) => [id, makeElement(id)]),
  );

  const domListeners = {};
  globalThis.document = {
    addEventListener: (name, handler) => {
      domListeners[name] = handler;
    },
    getElementById: (id) => elements.get(id),
  };

  const storageWrites = [];
  globalThis.localStorage = { setItem: () => storageWrites.push("localStorage") };
  globalThis.sessionStorage = { setItem: () => storageWrites.push("sessionStorage") };
  globalThis.window = { location: { protocol: "https:", origin: "https://localhost:47831" } };

  const office = {
    item: null,
    handlers: {},
    onReady: null,
    refuseItemChanged: false,
    context: {
      mailbox: {
        get item() {
          return office.item;
        },
        addHandlerAsync(type, handler) {
          // Office only allows this while the pane is pinned to a message.
          if (office.refuseItemChanged) throw new Error("not pinned");
          office.handlers[type] = handler;
        },
      },
    },
    EventType: { ItemChanged: "itemChanged" },
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded" },
    MailboxEnums: { ItemType: { Message: "message" } },
    onReady(callback) {
      office.onReady = callback;
      return Promise.resolve();
    },
  };
  globalThis.Office = office;

  const sent = [];
  let held = [];
  globalThis.fetch = (path, init) => {
    sent.push({ path, init });
    const settle = () => Promise.resolve(fetchResponse);
    if (!gate) return settle();
    return new Promise((resolve) => held.push(() => resolve(settle())));
  };

  // Evaluate a fresh copy of the real script against the globals just installed, so each test gets
  // its own module instance rather than sharing cached module state.
  new Function(SOURCE_TEXT)();

  return {
    office,
    sent,
    storageWrites,
    ui: {
      pairing: elements.get("pairing-code"),
      forget: elements.get("forget-button"),
      subject: elements.get("preview-subject"),
      from: elements.get("preview-from"),
      capture: elements.get("capture-button"),
      status: elements.get("status"),
    },

    /** Selects a message. Returns the array that body-read callbacks will be pushed onto. */
    setItem({ subject = "Project sync", sender = SENDER, body = true } = {}) {
      office.item = body
        ? {
            subject,
            from: sender,
            body: {
              getAsync(_coercion, callback) {
                reads.push(callback);
              },
            },
          }
        : null;
      reads.length = 0;
      return reads;
    },

    /** Moves the reading pane to another message, exactly as Office would report it. */
    switchTo(next) {
      this.setItem(next);
      office.handlers.itemChanged?.();
    },

    releaseHeldRequests: () => held.splice(0).forEach((release) => release()),

    start: async () => {
      domListeners.DOMContentLoaded();
      await office.onReady();
    },
  };
}

const tick = () => new Promise((resolve) => setTimeout(resolve, 5));

async function capture(pane, reads, body = "Body of the message") {
  pane.ui.pairing.value = "pairing-code-1";
  pane.ui.capture.dispatch("click");
  expect(reads.length).toBeGreaterThan(0);
  reads.forEach((callback) => callback({ status: "succeeded", value: body }));
  await tick();
  await tick();
}

describe("capturing only on request", () => {
  let pane;

  beforeEach(async () => {
    pane = boot();
    pane.setItem();
    await pane.start();
  });

  test("sends nothing while the pane is idle", async () => {
    expect(pane.sent).toHaveLength(0);
    expect(pane.ui.capture.disabled).toBe(false);
  });

  test("sends subject, sender and plain text body when the button is pressed", async () => {
    const reads = pane.setItem();
    await capture(pane, reads);

    expect(pane.sent).toHaveLength(1);
    expect(pane.sent[0].path).toBe("/capture");
    expect(pane.sent[0].init.method).toBe("POST");
    expect(pane.sent[0].init.headers["Content-Type"]).toBe("application/json");
    expect(pane.sent[0].init.headers["X-ReplyBrief-Pairing"]).toBe("pairing-code-1");
    expect(pane.sent[0].init.cache).toBe("no-store");

    const payload = JSON.parse(pane.sent[0].init.body);
    expect(payload.source).toBe("outlook");
    expect(payload.title).toBe("Project sync");
    expect(payload.participants).toBe("Dana Example <dana@example.com>");
    expect(payload.content).toBe("Body of the message");
    expect(payload.contextUsed.length).toBeGreaterThan(0);

    expect(pane.ui.status.textContent).toContain("Sent");
    expect(pane.ui.status.textContent).toContain("Project sync");
    expect(pane.ui.capture.disabled).toBe(false);
  });

  test("falls back to item.sender when item.from is absent", async () => {
    const reads = pane.setItem();
    pane.office.item.from = null;
    pane.office.item.sender = { displayName: "Lee Example", emailAddress: "lee@example.com" };

    await capture(pane, reads);

    expect(JSON.parse(pane.sent[0].init.body).participants).toBe("Lee Example <lee@example.com>");
  });
});

describe("the pairing code", () => {
  test("blocks the capture until it is supplied and is never persisted", async () => {
    const pane = boot();
    const reads = pane.setItem();
    await pane.start();

    pane.ui.capture.dispatch("click");
    reads.forEach((callback) => callback({ status: "succeeded", value: "Body" }));
    await tick();
    await tick();

    expect(pane.sent).toHaveLength(0);
    expect(pane.ui.status.textContent).toContain("Paste the pairing code");
    expect(pane.storageWrites).toEqual([]);
  });

  test("is trimmed and can be forgotten from the pane", async () => {
    const pane = boot();
    pane.setItem();
    await pane.start();

    pane.ui.pairing.value = "  spaced-code  ";
    pane.ui.capture.dispatch("click");
    pane.ui.forget.dispatch("click");

    expect(pane.ui.pairing.value).toBe("");
    expect(pane.ui.forget.hidden).toBe(true);
    expect(pane.ui.status.textContent).toContain("cleared");
    expect(pane.storageWrites).toEqual([]);
  });
});

describe("switching the selected message", () => {
  test("before the request is sent, the stale capture is dropped", async () => {
    const pane = boot();
    const reads = pane.setItem({ subject: "Second message" });
    await pane.start();
    pane.ui.pairing.value = "code";

    pane.ui.capture.dispatch("click");
    expect(reads).toHaveLength(1);

    // The reading pane moves on while the body read is still outstanding.
    pane.switchTo({ subject: "Third message", sender: { displayName: "Kim", emailAddress: "kim@example.com" } });
    reads.forEach((callback) => callback({ status: "succeeded", value: "Body of the second message" }));
    await tick();
    await tick();

    expect(pane.sent).toHaveLength(0);
    expect(pane.ui.status.textContent).toContain("changed");
    expect(pane.ui.subject.textContent).toBe("Third message");
  });

  test("after the request is sent, the send is reported truthfully", async () => {
    const pane = boot({ gate: true });
    const reads = pane.setItem({ subject: "Sixth message", sender: { displayName: "Kim", emailAddress: "kim@example.com" } });
    await pane.start();
    pane.ui.pairing.value = "code";

    pane.ui.capture.dispatch("click");
    reads.forEach((callback) => callback({ status: "succeeded", value: "Body of the sixth message" }));
    await tick();

    pane.switchTo({ subject: "Seventh message", sender: { displayName: "Lee", emailAddress: "lee@example.com" } });
    await tick();

    pane.releaseHeldRequests();
    await tick();
    await tick();

    expect(pane.sent).toHaveLength(1);
    expect(JSON.parse(pane.sent[0].init.body).title).toBe("Sixth message");
    expect(pane.ui.status.textContent).toContain("Sent");
    expect(pane.ui.status.textContent).not.toContain("nothing was sent");
    expect(pane.ui.status.textContent).toContain("different message");
  });

  test("a fresh capture after a switch uses the newly selected message", async () => {
    const pane = boot();
    await pane.start();

    pane.switchTo({ subject: "Fourth message", sender: { displayName: "Kim", emailAddress: "kim@example.com" } });
    const reads = pane.setItem({ subject: "Fourth message", sender: { displayName: "Kim", emailAddress: "kim@example.com" } });
    await capture(pane, reads);

    expect(pane.sent).toHaveLength(1);
    expect(JSON.parse(pane.sent[0].init.body).title).toBe("Fourth message");
  });
});

describe("failures", () => {
  test("no message open sends nothing and explains what to do", async () => {
    const pane = boot();
    await pane.start();
    pane.setItem({ body: false });
    pane.ui.pairing.value = "code";

    pane.ui.capture.dispatch("click");
    await tick();

    expect(pane.sent).toHaveLength(0);
    expect(pane.ui.status.textContent).toContain("reading pane");
  });

  test("a rejected pairing code is reported and cleared", async () => {
    const pane = boot({ fetchResponse: { ok: false, status: 401, text: async () => '{"error":"nope"}' } });
    const reads = pane.setItem();
    await pane.start();

    await capture(pane, reads);

    expect(pane.ui.status.textContent).toContain("pairing code");
    expect(pane.ui.pairing.value).toBe("");
    expect(pane.ui.forget.hidden).toBe(true);
  });

  test("a busy desktop is reported as busy rather than as a failure", async () => {
    const pane = boot({ fetchResponse: { ok: false, status: 429, text: async () => '{"error":"busy"}' } });
    const reads = pane.setItem();
    await pane.start();

    await capture(pane, reads);

    expect(pane.ui.status.textContent).toContain("busy");
  });

  test("a payload too large for the desktop is explained", async () => {
    const pane = boot({ fetchResponse: { ok: false, status: 413, text: async () => '{"error":"too large"}' } });
    const reads = pane.setItem();
    await pane.start();

    await capture(pane, reads);

    expect(pane.ui.status.textContent).toContain("too large");
  });

  test("an oversized message is refused rather than truncated", async () => {
    const pane = boot();
    const reads = pane.setItem({ subject: "Huge" });
    await pane.start();

    await capture(pane, reads, "x".repeat(60001));

    expect(pane.sent).toHaveLength(0);
    expect(pane.ui.status.textContent).toContain("60,001");
  });

  test("an empty body is refused", async () => {
    const pane = boot();
    const reads = pane.setItem();
    await pane.start();

    await capture(pane, reads, "");

    expect(pane.sent).toHaveLength(0);
    expect(pane.ui.status.textContent).toContain("readable body");
  });

  test("a body read failure never leaks the Office error text", async () => {
    const pane = boot();
    const reads = pane.setItem();
    await pane.start();
    pane.ui.pairing.value = "code";

    const warnings = [];
    const originalWarn = console.warn;
    console.warn = (value) => warnings.push(String(value));

    try {
      pane.ui.capture.dispatch("click");
      reads.forEach((callback) => callback({ status: "failed", error: { message: "SECRET server detail 12345" } }));
      await tick();
      await tick();
    } finally {
      console.warn = originalWarn;
    }

    expect(pane.sent).toHaveLength(0);
    expect(pane.ui.status.textContent).toContain("would not hand over");
    expect(pane.ui.status.textContent).not.toContain("SECRET");
    expect(warnings.join(" ")).not.toContain("SECRET");
  });

  test("an unreachable desktop points at the origin", async () => {
    const pane = boot();
    const reads = pane.setItem();
    await pane.start();

    globalThis.fetch = () => Promise.reject(new TypeError("Failed to fetch"));
    await capture(pane, reads);

    expect(pane.ui.status.textContent).toContain("https://localhost:47831");
  });
});

describe("rendering message metadata", () => {
  test("hostile subjects and bodies are inserted as text, never as markup", async () => {
    const pane = boot();
    const hostile = "<img src=x onerror=alert(1)>";
    const reads = pane.setItem({ subject: hostile, sender: { displayName: hostile, emailAddress: "a@b.c" } });
    await pane.start();

    // The preview goes through textContent, so there is no innerHTML for markup to land in.
    expect(pane.ui.subject.textContent).toBe(hostile);
    expect(pane.ui.subject.innerHTML).toBeUndefined();

    await capture(pane, reads, hostile);

    const payload = JSON.parse(pane.sent[0].init.body);
    expect(payload.title).toBe(hostile);
    expect(payload.content).toBe(hostile);
  });

  test("an over-long subject is truncated to the core limit", async () => {
    const pane = boot();
    const reads = pane.setItem({ subject: "s".repeat(900) });
    await pane.start();

    await capture(pane, reads);

    expect(JSON.parse(pane.sent[0].init.body).title.length).toBeLessThanOrEqual(500);
  });
});

describe("startup", () => {
  test("renders even when the item-change handler cannot be registered", async () => {
    const pane = boot();
    pane.setItem();
    pane.office.refuseItemChanged = true;
    await pane.start();

    expect(pane.ui.subject.textContent).toBe("Project sync");
    expect(pane.ui.capture.disabled).toBe(false);
  });
});