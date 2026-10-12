"use strict";

/*
 * ReplyBrief Outlook task pane.
 *
 * The pane is intentionally passive: it renders the selected message, waits for a human to press
 * the capture button, and forwards that one message to the ReplyBrief desktop app on this machine.
 * It never captures on its own, never stores anything durably, and never talks to an AI service.
 */

(function () {
  var CAPTURE_PATH = "/capture";
  var PAIRING_HEADER = "X-ReplyBrief-Pairing";
  var MAX_CONTENT_LENGTH = 60000;
  var MAX_TITLE_LENGTH = 500;
  var MAX_PARTICIPANTS_LENGTH = 2000;
  var CONTEXT_USED = "Outlook reading pane: the selected message's subject, sender and body text.";

  var ui = {};

  /*
   * The pairing code is held in a JavaScript variable for the lifetime of this task pane and in a
   * password field. It is never written to localStorage, sessionStorage, IndexedDB, a cookie, a URL
   * fragment or the document title.
   */
  var pairingCode = "";

  /*
   * Bumped every time Outlook tells us a different item is selected. A capture remembers the value
   * it started with and refuses to deliver anything once it has moved, which is what stops a slow
   * body read from attaching the previous message to a click that was meant for the next one.
   */
  var generation = 0;
  var capturing = false;

  document.addEventListener("DOMContentLoaded", function () {
    ui = {
      pairing: document.getElementById("pairing-code"),
      forget: document.getElementById("forget-button"),
      subject: document.getElementById("preview-subject"),
      from: document.getElementById("preview-from"),
      capture: document.getElementById("capture-button"),
      status: document.getElementById("status"),
    };

    ui.capture.addEventListener("click", function () {
      captureSelectedItem();
    });

    ui.forget.addEventListener("click", function () {
      pairingCode = "";
      ui.pairing.value = "";
      ui.forget.hidden = true;
      setStatus("Pairing code cleared from this task pane.", "info");
      ui.pairing.focus();
    });

    ui.pairing.addEventListener("input", function () {
      if (ui.pairing.value.trim() === "") pairingCode = "";
      ui.forget.hidden = pairingCode === "";
    });

    if (typeof Office === "undefined" || !Office.context) {
      setStatus("The Outlook runtime did not load, so this pane cannot read the message.", "error");
      ui.capture.disabled = true;
      return;
    }

    if (window.location.protocol !== "https:") {
      setStatus("This pane must be served over HTTPS. Ask your administrator to check the ReplyBrief desktop app.", "error");
      ui.capture.disabled = true;
      return;
    }

    Office.onReady(function () {
      renderSelectedItem();
      watchForItemChanges();

      if (pairingCode === "") ui.pairing.focus();
    }).catch(function (error) {
      reportUnexpected("Outlook could not finish starting the task pane.");
    });
  });

  /* Outlook only exposes a body for message and appointment items, and the manifest puts this
   * command on the Message tab, so an item with a body here is the selected message. */
  function selectedItem() {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    return item && item.body ? item : null;
  }

  function subjectOf(item) {
    try {
      return truncate(item.subject || "(no subject)", MAX_TITLE_LENGTH);
    } catch (error) {
      return "(no subject)";
    }
  }

  function participantsOf(item) {
    var candidates = [safe(() => item.from), safe(() => item.sender)];
    for (var i = 0; i < candidates.length; i += 1) {
      var person = candidates[i];
      if (!person) continue;

      var address = person.emailAddress || "";
      var name = person.displayName || "";
      if (!address && !name) continue;

      return truncate(name && address ? name + " <" + address + ">" : name || address, MAX_PARTICIPANTS_LENGTH);
    }

    return truncate(item.displayName || "", MAX_PARTICIPANTS_LENGTH);
  }

  function readBody(item) {
    return new Promise(function (resolve, reject) {
      item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          // The Office error text is deliberately not surfaced or logged: it is the only part of
          // this flow that a host controls, and forwarding it verbatim risks leaking content into
          // whatever sink the string ends up in.
          reject(new Failure("Outlook would not hand over the message body."));
          return;
        }

        resolve(typeof asyncResult.value === "string" ? asyncResult.value : "");
      });
    });
  }

  function captureSelectedItem() {
    if (capturing) return;

    var item = selectedItem();
    if (!item) {
      setStatus("Open a message in the reading pane first. ReplyBrief captures one selected message at a time.", "error");
      return;
    }

    var code = readPairingCode();
    if (code === "") {
      setStatus("Paste the pairing code that ReplyBrief shows on this computer.", "error");
      ui.pairing.focus();
      return;
    }

    var startedAt = generation;
    var subject = subjectOf(item);

    capturing = true;
    ui.capture.disabled = true;
    setStatus("Reading the selected message\u2026", "info");

    readBody(item)
      .then(function (content) {
        if (startedAt !== generation) throw staleItem();

        if (content.length === 0) throw new Failure("This message has no readable body text.");
        if (content.length > MAX_CONTENT_LENGTH) {
          throw new Failure("This message is " + content.length.toLocaleString() + " characters long. ReplyBrief does not send part of a message; capture a shorter one.");
        }

        setStatus("Sending to ReplyBrief\u2026", "info");

        return postCapture(
          {
            source: "outlook",
            title: subject,
            participants: participantsOf(item),
            content: content,
            contextUsed: CONTEXT_USED,
          },
          code
        );
      })
      .then(function () {
        // The capture is already on the desktop by this point. If the pane moved on while the
        // request was in flight, say so plainly instead of implying nothing was sent.
        setStatus(
          startedAt === generation
            ? "Sent \u201c" + subject + "\u201d to ReplyBrief."
            : "Sent \u201c" + subject + "\u201d to ReplyBrief. You selected a different message in the meantime, so check the desktop app to see which capture it picked up.",
          "success"
        );
      })
      .catch(function (error) {
        if (error instanceof Failure) {
          // A stale item or a stale pairing code is something the person can act on; it is a warning
          // rather than a failure, and it must look the same whichever path reported it.
          setStatus(error.message, error.reason === "capture" ? "error" : "warn");
          if (error.reason === "pairing") {
            pairingCode = "";
            ui.pairing.value = "";
            ui.forget.hidden = true;
            ui.pairing.focus();
          }
          return;
        }

        reportUnexpected("ReplyBrief could not be reached at " + window.location.origin + ". Check that ReplyBrief is running and that the pairing code is current.");
      })
      .then(function () {
        capturing = false;
        renderSelectedItem();
      });
  }

  function postCapture(payload, code) {
    return fetch(CAPTURE_PATH, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-ReplyBrief-Pairing": code,
      },
      body: JSON.stringify(payload),
      cache: "no-store",
      credentials: "omit",
      referrerPolicy: "no-referrer",
    }).then(function (response) {
      if (response.ok) return;

      if (response.status === 401) throw new Failure("ReplyBrief rejected that pairing code. Paste the code it is showing now.", "pairing");
      if (response.status === 403) throw new Failure("ReplyBrief refused the request because it did not come from its own task pane.");
      if (response.status === 413) throw new Failure("This message is too large for ReplyBrief to accept.");
      if (response.status === 415) throw new Failure("ReplyBrief rejected the request format.");
      if (response.status === 429) throw new Failure("ReplyBrief is busy and already has captures waiting. Give it a moment, then capture this one again.");

      return response.text().then(function (body) {
        var reason = readError(body);
        throw new Failure(reason === null ? "ReplyBrief rejected the capture (HTTP " + response.status + ")." : reason);
      });
    });
  }

  function readError(body) {
    try {
      var parsed = JSON.parse(body);
      return typeof parsed.error === "string" && parsed.error !== "" ? parsed.error : null;
    } catch (error) {
      return null;
    }
  }

  function watchForItemChanges() {
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
        generation += 1;
        renderSelectedItem();

        if (capturing) {
          setStatus("The selected message changed. The capture in progress will not be sent.", "warn");
        }
      });
    } catch (error) {
      /* Only available while the task pane is pinned to a message. When it is not pinned, the
       * reading pane cannot switch items underneath us, so the generation guard has nothing to
       * catch and the button still behaves correctly. */
    }
  }

  function renderSelectedItem() {
    var item = selectedItem();

    if (!item) {
      ui.subject.textContent = "No message selected";
      ui.from.textContent = "Open a message in the reading pane.";
      ui.capture.disabled = capturing;
      return;
    }

    ui.subject.textContent = subjectOf(item);
    ui.from.textContent = participantsOf(item) || "(unknown sender)";
    ui.capture.disabled = capturing;
  }

  function readPairingCode() {
    var pasted = ui.pairing.value.trim();
    if (pasted !== "") pairingCode = pasted;

    ui.forget.hidden = pairingCode === "";
    return pairingCode;
  }

  function setStatus(message, tone) {
    ui.status.textContent = message;
    ui.status.className = "status" + (tone && tone !== "info" ? " " + tone : "");
  }

  function reportUnexpected(message) {
    setStatus(message, "error");
  }

  function staleItem() {
    return new Failure("The selected message changed while it was being read, so nothing was sent.", "stale");
  }

  function truncate(value, limit) {
    return value.length > limit ? value.slice(0, limit - 1) + "\u2026" : value;
  }

  function safe(read) {
    try {
      return read();
    } catch (error) {
      return null;
    }
  }

  function Failure(message, reason) {
    this.message = message;
    this.reason = reason || "capture";
  }

  Failure.prototype = Object.create(Error.prototype);
  Failure.prototype.constructor = Failure;
})();