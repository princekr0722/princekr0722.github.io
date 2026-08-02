/**
 * ============================================================================
 *  Router — the project's only doPost / doGet
 * ============================================================================
 *
 *  An Apps Script project gets exactly one doPost and one doGet. Every file
 *  shares a single global namespace, so a second definition of either just
 *  silently overrides the first — no error, no warning, one feature quietly
 *  stops working. This file owns both and dispatches to:
 *
 *    2-analytics.js  ->  handleAnalytics(e)   /  analyticsHealth(e)
 *    3-assistant.js  ->  handleChat(e, body)  /  chatHealth()
 *                        handleChatHistory(e, body)
 *
 *  Routing rule: a POST body with  action: "chat"  or  "chat_history"  goes
 *  to the assistant. Everything else goes to analytics — so the logging
 *  already deployed on the site keeps working untouched.
 *
 *  Deploy note: re-deploy as a NEW VERSION after editing any file in the
 *  project ("Manage deployments" -> edit -> Version: New version). Until you
 *  do, the /exec URL keeps serving the old code.
 *
 * ============================================================================
 */

/**
 * Single POST entry point for the whole project.
 *
 * Deliberately holds no lock. handleAnalytics takes its own script lock for
 * the few milliseconds it needs to append a row; a chat turn waits on the
 * model for seconds, and locking here would queue every analytics write
 * behind it.
 */
function doPost(e) {
  let body = null;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    // Leave body null and fall through to analytics, which reports the
    // parse failure in its own error format.
  }

  // An `action` field means the caller wants the assistant. An unrecognised
  // one is an error, NOT analytics — otherwise a typo'd or not-yet-deployed
  // action silently lands in a "General_Logs" tab and looks like it worked.
  if (body && body.action) {
    if (body.action === "chat") return handleChat(e, body);
    if (body.action === "chat_history") return handleChatHistory(e, body);
    return createJSONOutput({
      status: "error",
      message: "Unknown action: " + body.action,
    });
  }

  return handleAnalytics(e);
}

/**
 * Single GET entry point.
 *
 *   /exec             -> analytics health
 *   /exec?page=chat   -> chat health (confirms the API key is readable)
 */
function doGet(e) {
  const page = e && e.parameter ? e.parameter.page : "";
  return page === "chat" ? chatHealth() : analyticsHealth(e);
}
