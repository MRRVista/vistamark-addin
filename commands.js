/*
 * commands.js -- Outlook add-in command-surface runtime for Vistamark AI.
 *
 * v5.16.0: registers the OnMessageSend handler used for edit-distance
 * telemetry. The handler runs in a SEPARATE Office runtime from the
 * taskpane (taskpane.html) -- they cannot share JavaScript state. The
 * coordination is via Outlook CustomProperties (per-item Outlook-managed
 * storage):
 *
 *   1. Taskpane's Replace Body stashes vistamarkDraftToken (and
 *      vistamarkDraftText, vistamarkDraftInsertedAt) into CustomProperties
 *      at the moment the draft is inserted.
 *   2. This handler reads vistamarkDraftToken at click-Send time, reads
 *      the final-sent body via item.body.getAsync(Text), POSTs both to
 *      /api/inbound/edit-telemetry on the Vercel backend, then calls
 *      event.completed({allowEvent: true}) to release the send.
 *
 * Manifest SendMode is "soft-block" -- if this handler doesn't complete
 * within 5 seconds the send proceeds anyway with no user-visible UI. Our
 * handler is fast in the normal path (~200ms), but we wrap the POST in
 * a Promise.race with a 3-second timeout to keep well under the limit.
 *
 * Telemetry is best-effort. ALL errors are swallowed; we ALWAYS call
 * event.completed({allowEvent: true}) so the partner's send is never
 * blocked by a telemetry failure. This is the most important invariant
 * in this file -- changing it can lock the partner out of sending.
 */

var TELEMETRY_URL = "https://vistarandall.app/api/inbound/edit-telemetry";
var TELEMETRY_TIMEOUT_MS = 3000;

Office.onReady(function() {
  // Office.actions.associate registers the handler name (must match
  // FunctionName in manifest.xml -> LaunchEvent) with the actual function
  // reference. Available in Mailbox 1.10+ (manifest declares 1.12 minimum
  // for the V1_1 LaunchEvent block, so we're safe).
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
});

/**
 * OnMessageSend event handler. Wired via manifest.xml LaunchEvent.
 *
 * Lifecycle:
 *   1. Partner clicks Send in Outlook compose.
 *   2. Outlook fires OnMessageSend, spawns commands.html runtime, loads
 *      this script, calls onMessageSendHandler(event).
 *   3. We do telemetry work (read body, read token, POST diff request).
 *   4. We call event.completed({allowEvent: true}) -- send proceeds.
 *
 * The function is intentionally simple: it kicks off doTelemetry and
 * registers a .then/.catch that always calls event.completed. The
 * actual telemetry logic is in doTelemetry(); errors there are
 * swallowed so the .then path always runs.
 */
function onMessageSendHandler(event) {
  doTelemetry()
    .catch(function() { /* swallow -- telemetry must never block send */ })
    .then(function() {
      try {
        event.completed({ allowEvent: true });
      } catch (e) {
        // Runtime may have already moved on (e.g. soft-block timeout
        // fired). Nothing to do; the send was released either way.
      }
    });
}

/**
 * Read body + token in parallel, then POST with timeout. Returns a
 * promise that resolves when either the POST completes OR the timeout
 * fires (whichever is first). Never rejects -- all errors are swallowed
 * to keep the .then path in the handler reliable.
 */
function doTelemetry() {
  return new Promise(function(resolve) {
    var item = Office.context.mailbox.item;
    if (!item) { resolve(); return; }
    var senderEmail =
      (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || "";

    Promise.all([readBody(item), readDraftToken(item)]).then(
      function(arr) {
        var finalBody = arr[0];
        var draftToken = arr[1];
        if (!draftToken) {
          // No token in CustomProperties -- either Randall wasn't used
          // for this draft, or the taskpane stash failed. Nothing to do.
          resolve();
          return;
        }

        var postPromise = fetch(TELEMETRY_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            draft_token: draftToken,
            final_body: finalBody,
            sender_email: senderEmail
          })
        }).then(
          function() { /* success path -- nothing to do, telemetry is fire-and-forget from the partner's view */ },
          function() { /* error path -- swallow; telemetry must never block send */ }
        );

        var timeoutPromise = new Promise(function(r) {
          setTimeout(r, TELEMETRY_TIMEOUT_MS);
        });

        // Whichever wins, resolve outer promise so the handler's .then
        // fires and event.completed is called.
        Promise.race([postPromise, timeoutPromise]).then(
          function() { resolve(); },
          function() { resolve(); }
        );
      },
      function() {
        // Body or token read rejected -- swallow.
        resolve();
      }
    );
  });
}

/** Read the final-sent body as plain text. Resolves to "" on any error. */
function readBody(item) {
  return new Promise(function(resolve) {
    if (!item.body || !item.body.getAsync) { resolve(""); return; }
    try {
      item.body.getAsync(Office.CoercionType.Text, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          resolve("");
        }
      });
    } catch (e) {
      resolve("");
    }
  });
}

/** Read vistamarkDraftToken from Outlook CustomProperties. Resolves to "" if missing. */
function readDraftToken(item) {
  return new Promise(function(resolve) {
    if (!item.loadCustomPropertiesAsync) { resolve(""); return; }
    try {
      item.loadCustomPropertiesAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) { resolve(""); return; }
        var props = result.value;
        try {
          var token = props.get("vistamarkDraftToken") || "";
          resolve(String(token));
        } catch (e) {
          resolve("");
        }
      });
    } catch (e) {
      resolve("");
    }
  });
}
