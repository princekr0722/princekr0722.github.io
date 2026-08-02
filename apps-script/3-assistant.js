/**
 * ============================================================================
 *  Portfolio AI assistant — Google Apps Script backend (DeepSeek)
 * ============================================================================
 *
 *  Lets visitors on princekr0722.github.io chat with an AI that answers
 *  questions about Prince Kumar.
 *
 *  This file is the whole backend. Paste it into an Apps Script project, put
 *  the DeepSeek key in your Sheet, deploy as a web app, POST to the /exec URL.
 *
 *  It is .js here for editor tooling only — Apps Script stores server-side
 *  files as .gs regardless, and clasp converts .js on push. This does not run
 *  in a browser: UrlFetchApp, SpreadsheetApp, CacheService and ContentService
 *  are Apps Script runtime globals.
 *
 *  ── THREE FILES, ONE PROJECT ──────────────────────────────────────────────
 *
 *  This project serves both the analytics logger and this chat assistant, so
 *  it needs all three files (Apps Script allows only one doPost per project;
 *  1-router.js owns it):
 *
 *    1-router.js     doPost / doGet  ->  dispatches on  action: "chat"
 *    2-analytics.js  handleAnalytics(e) / analyticsHealth(e)
 *    3-assistant.js  handleChat(e, body) / chatHealth()   <- this file
 *
 *  Chat requests must include  action: "chat"  in the body. Analytics posts
 *  are untouched, so the logging already live on the site keeps working.
 *
 *  ── SETUP ─────────────────────────────────────────────────────────────────
 *
 *  1. In your Google Sheet, add a tab named exactly:  Config
 *     Put two columns in it (row 1 is a header, ignored):
 *
 *         A                  B
 *     1   name               value
 *     2   DEEPSEEK_API_KEY   sk-...
 *
 *     Anyone who can open this spreadsheet can read that key, so keep the
 *     sheet private — do not share it broadly or publish it to the web.
 *
 *  2. Put the spreadsheet's ID in CONFIG.SPREADSHEET_ID below. It is the long
 *     id in the sheet URL:
 *       docs.google.com/spreadsheets/d/<THIS_PART>/edit
 *     If this script lives inside that spreadsheet (Extensions → Apps Script),
 *     you can leave SPREADSHEET_ID empty and it will use the active one.
 *
 *  3. Deploy → New deployment → type "Web app"
 *         Execute as:        Me
 *         Who has access:    Anyone
 *     Copy the /exec URL — that is your endpoint.
 *
 *  4. Re-deploy after every edit ("Manage deployments" → edit → new version),
 *     or the live URL keeps serving the old code.
 *
 *  ── CALLING IT FROM THE SITE ──────────────────────────────────────────────
 *
 *  Content-Type MUST be text/plain. Apps Script does not answer CORS
 *  preflight requests, and application/json triggers one — the browser would
 *  block the call before it ever reaches Google. text/plain is a
 *  CORS-safelisted type, so no preflight is sent. The body is still JSON;
 *  only the declared type differs.
 *
 *    const res = await fetch(ENDPOINT_URL, {
 *      method: "POST",
 *      headers: { "Content-Type": "text/plain;charset=utf-8" },
 *      body: JSON.stringify({
 *        action: "chat",           // required — tells the router this is chat
 *        message: "What has Prince built with AI?",
 *        visitorId: "uuid-here"    // the site's existing uid; see below
 *      })
 *    });
 *    const data = await res.json();   // { ok, reply, turns } | { ok:false, error }
 *
 *  The client does NOT send conversation history. Each visitor's transcript
 *  is stored server-side, one row per uid, in the Chat_History tab — so it
 *  survives refreshes and returning visits, and cannot be tampered with from
 *  the browser. Pass the same visitorId the analytics logger uses (its
 *  getVisitorId(), persisted in localStorage) and the two line up.
 *
 *  To restore the transcript when the chat panel opens:
 *
 *    body: JSON.stringify({ action: "chat_history", visitorId })
 *    -> { ok: true, history: [ { role: "user"|"assistant", text }, ... ] }
 *
 *  Note this means visitor questions are retained in your spreadsheet.
 *
 *  ── COST / ABUSE ──────────────────────────────────────────────────────────
 *
 *  This endpoint is public and spends your DeepSeek credits. The limits in
 *  CONFIG below are the only thing between a bored visitor and your bill.
 *  Keep them, and set a spend cap in the DeepSeek console too.
 *
 * ============================================================================
 */

// ────────────────────────────────────────────────────────────────────────────
// Config
// ────────────────────────────────────────────────────────────────────────────

const CONFIG = {
  // ── Where the API key lives ──
  SPREADSHEET_ID: '',            // leave '' if the script is bound to the sheet
  CONFIG_SHEET_NAME: 'Config',
  KEY_NAME: 'DEEPSEEK_API_KEY',
  KEY_CACHE_SECONDS: 300,        // avoid a Sheets read on every single message

  // ── Where conversations live: one row per visitor, keyed by uid ──
  HISTORY_SHEET_NAME: 'Chat_History',
  MAX_HISTORY_CHARS: 20000,      // a cell holds 50k; stay well under it

  // ── Model ──
  API_URL: 'https://api.deepseek.com/chat/completions',

  // 'deepseek-chat' is fast and cheap — right for short factual Q&A.
  // 'deepseek-reasoner' thinks first: better on hard questions, slower and
  // pricier, and it ignores TEMPERATURE.
  MODEL: 'deepseek-chat',

  TEMPERATURE: 0.6,              // low-ish: this bot should stick to facts
  MAX_TOKENS: 800,               // replies are meant to be short

  // ── Limits ──
  MAX_MESSAGE_CHARS: 1500,       // per incoming message
  MAX_HISTORY_TURNS: 12,         // older turns are dropped
  MAX_MSGS_PER_VISITOR: 25,      // per window
  RATE_WINDOW_SECONDS: 600,      // 10 minutes
  MAX_MSGS_GLOBAL: 120,          // across all visitors, same window
};

// ────────────────────────────────────────────────────────────────────────────
// Who the assistant is talking about
//
// Deliberately compact for now — richer, structured context (projects, work
// history, resume text, retrieval) comes later. Everything the assistant is
// allowed to say should be in here; it is told not to invent the rest.
// ────────────────────────────────────────────────────────────────────────────

const PROFILE = `
# WHO

Prince Kumar — Software Developer & System Designer, based in India.
Backend-heavy and product-minded: scalable microservices, distributed
systems, and more recently AI systems grounded in a company's real code,
data and logs. Currently open to new opportunities ("available for hire").

Started coding 29 Aug 2022. Professional career started 20 Sep 2023, so his
years of experience should be counted from that date.

# WORK

## DrinkPrime — Software Developer (Sep 2023 – present)
DrinkPrime is a subscription-based water purifier service (drinkprime.in).

### DPatcher — autonomous AI engineer
An AI engineer with deep, grounded knowledge of DrinkPrime's entire system.
It answers from live sources rather than guessing, cross-referencing a
pgvector knowledge base (Confluence + learned resolutions), a code graph of
the actual repositories, read-only production databases (MySQL/Mongo), and
GKE runtime logs.
- A general-purpose engineering teammate the whole org uses: understanding
  how a feature works and why, debugging and root-causing live issues across
  code + data + logs, drafting test cases and edge cases, triaging and
  resolving on-call incidents, and answering "where/how is X handled"
  onboarding questions — through both Slack and a web dashboard, with
  confidentiality controls that keep internal detail off customer-facing
  threads.
- It gets smarter over time: confirmed resolutions fold back into the
  knowledge base so recurring problems become self-answering, reducing repeat
  escalations and dependence on a few subject-matter experts.
- Prince owned the whole product and its deployment: a role-based admin
  dashboard (Kong auth-gateway JWT, read/write access tiers), a private staff
  DM assistant, live-tunable runtime settings, and a Dockerized rollout wired
  to external Postgres, MySQL/Mongo and a code-graph service.
- Stack: Node.js, Fastify, PostgreSQL/pgvector, NVIDIA LLMs, embeddings,
  React/Vite, Docker, GKE.

### DPengram — code-knowledge platform
A self-contained platform that turns Git repositories into a queryable code
graph — symbols, LSP-resolved call/usage edges, semantic embeddings — exposed
through a single MCP endpoint and a REST API. It is the code-graph backbone
behind DPatcher.
- A transparent MCP proxy forwards all upstream engine tools unchanged while
  adding first-party capabilities: repository discovery and ripgrep-powered
  filename/regex search. Any MCP client (Claude Code, Cursor) connects without
  modification and inherits future engine capabilities automatically.
- Reliable sync and deployment: serialized/coalesced indexing, incremental
  re-indexing from remote branch changes, stable project identities that
  self-heal after a repository moves, Docker-based org deployment, and a
  zero-dependency single-file desktop executable with automatic engine
  provisioning, self-updates and a React management UI.
- What it changes day to day: faster development (instantly locates the right
  code, its callers and usages across every repo); quicker root-cause analysis
  (trace a bug's call path and blast radius in seconds instead of grepping);
  fewer hallucinations (answers grounded in real indexed code, semantic and
  full-text); cross-repo questions answered coherently in one go.
- Stack: MCP, code graph, LSP, REST API, embeddings, ripgrep, React, Docker.

### Core platform & backend
- Single-handedly designed and built a centralised security mechanism for
  DrinkPrime: API restrictions, roles and permissions through one single point
  for every microservice (Spring, Kong API Gateway).
- Fixed race conditions that were over-booking technician slots, using
  distributed locking (Redis/Redisson).
- Upgraded RESTful inter-service communication to event-driven systems (Kafka,
  RabbitMQ), and migrated a legacy PHP system to Spring Boot including
  database migrations and backward compatibility.

## Freelance Software Developer — self-employed (Oct 2023 – present)
Alongside full-time work, to keep learning and stay current. Services:
complete backend, frontend, cloud, custom software and e-commerce.

# SKILLS

AI & LLM engineering (his current focus): LLM agents, RAG pipelines, pgvector,
embeddings, MCP servers, code graphs, NVIDIA LLMs, prompt engineering,
Slack AI bots. Production AI systems that answer from real sources — code,
databases, logs — instead of guessing.

Languages: Java, JavaScript, TypeScript, Go, Python, Lua, HTML, CSS.
Frameworks: Spring, Spring Boot, Hibernate, React.js, Node.js, Next.js,
Fastify.
Databases: MySQL, MongoDB, PostgreSQL, MS SQL Server, Redis.
Cloud & DevOps: AWS, Google Cloud, Firebase, Kubernetes, Docker, Kong.
Architecture: Kafka, RabbitMQ, microservices, DSA.

Strongest in Java/Spring Boot backends and, lately, AI engineering.

# EDUCATION

- Bachelor of Computer Applications, Amity University Noida (2026–2029).
  C++, Java, Python, AI/ML, software development. Formalises the engineering
  work he is already doing professionally.
- Full Stack Web Development, Masai School (2022–2023). Intensive program:
  full-stack development, Java backend, Spring Boot, MERN, DSA, agile and
  professional skills; multiple real-world team projects.
- Higher Secondary Education (2020–2022), science and mathematics — physics,
  chemistry, mathematics. Took part in technical competitions and programming
  challenges.

# PROJECTS

Flagship AI work (built at DrinkPrime, not public): DPatcher and DPengram —
described under WORK above.

Personal and academic side projects:
- Insightgram — a social networking platform to connect people globally.
  Java, Spring Boot, microservices, Spring Security, MySQL, WebSocket.
  Live: https://insightgram.netlify.app
  Code: https://github.com/princekr0722/Insightgram_Main_Backend_Service
- Estate Explorer — real estate platform for property listing and management.
  Java, Spring Boot, MySQL.
  Code: https://github.com/dikshant123321/Real-Estate-Broker-Application-
  Presentation:
  https://drive.google.com/file/d/1pgnoh-SSTSRZIqLJsBwh1rLfGPDzKvyX/view
- Bus Buddy — online bus ticket booking platform, "ride in style, book a
  mile". Java.
  Code: https://github.com/princekr0722/inexpensive-cause-3321
  Presentation:
  https://drive.google.com/file/d/1v5XrzAQipbzSQiNrmaTda8DIXs7HrNk_/view
- Licious Clone — online meat, poultry and seafood delivery storefront.
  HTML, CSS, JS.
  Live: https://candid-sorbet-838343.netlify.app
  Code: https://github.com/AditiSharma00/truthful-wing-8761
- PharmEasy Clone — online healthcare and medicine delivery platform.
  HTML, CSS, JS.
  Live: https://steady-jalebi-560beb.netlify.app
  Code: https://github.com/princekr0722/versed-substance-8213

# NUMBERS

- 1500+ DSA problems solved.
- 15+ projects completed.
- 10+ technologies in regular use.
- Codes roughly 5 hours a day since 29 Aug 2022.

# LINKS & CONTACT

Share these freely — a visitor asking how to reach him, see his code, or read
his resume should get the actual link, not a "check the contact section".

- Email: knownasprincekr@gmail.com
- Phone: +91 95234 34209
- LinkedIn: https://www.linkedin.com/in/prince-kumar-7b9194247/
- GitHub: https://github.com/princekr0722
- Portfolio: https://princekr0722.github.io
- Resume (PDF):
  https://drive.google.com/file/d/1IE8nYQJuCAZkeaaYaVQdGL_VD2iI81hD/view
- Current employer: https://drinkprime.in

On the portfolio page itself, the sections are Home, About, Skills,
Experience, Education, Projects, GitHub and Contact — you can point someone
at a section by name if they want to read it themselves.
`.trim();

const SYSTEM_PROMPT = `
You are Pi, the assistant on Prince Kumar's portfolio site. Visitors are
usually recruiters, hiring managers, or engineers working out what he has
built and whether he fits a role. If anyone asks who you are: you are Pi,
and you are here to talk about Prince.

VOICE
You genuinely enjoy this. You find Prince's work interesting and it shows —
warm, upbeat, a little playful, quick with a well-placed aside. Write like a
sharp colleague who likes the guy, not like a brochure.

Contractions always. An exclamation mark or a single emoji is welcome when
something actually deserves it — an AI engineer that debugs production from
live logs is genuinely cool, so say so. Do not sprinkle them everywhere, do
not open every reply the same way, and never manufacture excitement for
something ordinary. Recruiters are reading: delightful and useful beats
bubbly and empty. Cheerful never means vague — specifics are what make the
enthusiasm land.

GROUND RULES
Answer only from the profile below. If something is not in it — salary,
notice period, an employer not listed, anything personal — say so cheerfully
and point them to Prince directly. Never guess at dates, numbers, or
availability. Being fun is not a licence to make things up.

Speak about Prince in the third person. Keep answers to a few sentences
unless asked to go deeper; this renders in a small chat window. **Bold** for
emphasis is fine. No markdown headers, no bullet-point walls. Talk about the
engineering plainly rather than stacking buzzwords.

If someone asks about something unrelated to Prince, deflect with good
humour and steer back.

Share links as plain URLs when they are useful — resume, GitHub, a live demo.

--- PROFILE ---
${PROFILE}
--- END PROFILE ---
`.trim();

/**
 * The system prompt plus today's date.
 *
 * The date goes last on purpose: DeepSeek caches on a prefix, so a value that
 * changes daily would invalidate the whole profile if it sat at the top.
 */
function buildSystemPrompt() {
  const today = Utilities.formatDate(
    new Date(), 'Asia/Kolkata', 'EEEE, d MMMM yyyy'
  );
  return SYSTEM_PROMPT + '\n\nToday is ' + today +
    '. Use it to work out durations — years of experience, how long he has ' +
    'been at DrinkPrime — rather than guessing.';
}

// ────────────────────────────────────────────────────────────────────────────
// Entry points
//
// Deliberately NOT named doPost/doGet — a project gets only one of each, and
// the analytics logger already owns them. A router dispatches to these; see
// the header for the router code.
// ────────────────────────────────────────────────────────────────────────────

/**
 * Health check — hit the /exec URL with ?page=chat to see this.
 *
 * `historyStorage: "sheet"` only exists in this build. If the deployed
 * response lacks that field, the /exec URL is still serving an older version
 * and you need to re-deploy (Manage deployments -> edit -> New version).
 */
function chatHealth() {
  let keyConfigured = false;
  try {
    keyConfigured = Boolean(getApiKey());
  } catch (err) {
    // fall through — reported as false
  }

  let storedVisitors = 'unavailable';
  try {
    storedVisitors = Math.max(getHistorySheet().getLastRow() - 1, 0);
  } catch (err) {
    console.error('history sheet unreachable: ' + err);
  }

  return jsonOut({
    ok: true,
    service: 'portfolio-assistant',
    model: CONFIG.MODEL,
    keyConfigured: keyConfigured,
    historyStorage: 'sheet',
    historySheet: CONFIG.HISTORY_SHEET_NAME,
    storedVisitors: storedVisitors,
  });
}

/**
 * Chat endpoint.
 *
 * @param {Object} e         the Apps Script event object
 * @param {Object} [parsed]  the already-parsed body, if the router has one
 */
function handleChat(e, parsed) {
  try {
    let body = parsed;

    if (!body) {
      if (!e || !e.postData || !e.postData.contents) {
        return jsonOut({ ok: false, error: 'Empty request body.' });
      }
      try {
        body = JSON.parse(e.postData.contents);
      } catch (err) {
        return jsonOut({ ok: false, error: 'Body must be JSON.' });
      }
    }

    const message = String(body.message || '').trim();
    if (!message) {
      return jsonOut({ ok: false, error: 'No message provided.' });
    }
    if (message.length > CONFIG.MAX_MESSAGE_CHARS) {
      return jsonOut({
        ok: false,
        error: 'That message is too long — keep it under ' +
          CONFIG.MAX_MESSAGE_CHARS + ' characters.',
      });
    }

    const visitorId = String(body.visitorId || 'anonymous').slice(0, 64);
    const limit = checkRateLimit(visitorId);
    if (!limit.allowed) {
      return jsonOut({ ok: false, error: limit.reason, rateLimited: true });
    }

    // History is server-side — the client sends only the new message.
    const history = loadHistory(visitorId);

    const messages = [{ role: 'system', content: buildSystemPrompt() }];
    history.forEach(function (turn) {
      messages.push({ role: turn.role, content: turn.text });
    });
    messages.push({ role: 'user', content: message });

    const result = callDeepSeek(messages);
    if (!result.ok) {
      return jsonOut({ ok: false, error: result.error });
    }

    // Re-reads the row under a lock before writing, so a second message from
    // the same visitor mid-flight cannot clobber this one.
    const turns = appendHistory(visitorId, message, result.reply);

    return jsonOut({
      ok: true,
      reply: result.reply,
      turns: turns,
      usage: result.usage,
    });
  } catch (err) {
    console.error('handleChat failed: ' + err);
    return jsonOut({ ok: false, error: 'Something broke on my end. Try again.' });
  }
}

/**
 * Returns a visitor's stored conversation so the widget can restore it when
 * the panel opens. Body: { action: "chat_history", visitorId }
 */
function handleChatHistory(e, parsed) {
  try {
    let body = parsed;
    if (!body) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (err) {
        return jsonOut({ ok: false, error: 'Body must be JSON.' });
      }
    }

    const visitorId = String(body.visitorId || '').slice(0, 64);
    if (!visitorId) return jsonOut({ ok: true, history: [] });

    return jsonOut({ ok: true, history: loadHistory(visitorId) });
  } catch (err) {
    console.error('handleChatHistory failed: ' + err);
    return jsonOut({ ok: true, history: [] });
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Conversation storage — one row per visitor in the Chat_History tab
//
//   A uid | B first_seen | C last_seen | D turns | E history (JSON)
//
// Reads happen unlocked; the read-modify-write on append takes a script lock,
// but only *after* the model call, so a slow answer never blocks anyone else.
// ────────────────────────────────────────────────────────────────────────────

const HISTORY_HEADERS = ['uid', 'first_seen', 'last_seen', 'turns', 'history'];

function getHistorySheet() {
  const ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(CONFIG.HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.HISTORY_SHEET_NAME);
    sheet.appendRow(HISTORY_HEADERS);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 400);
  }
  return sheet;
}

/**
 * Row number for a uid, or 0. The lookup is cached, and the cached row is
 * re-checked against column A before use in case rows were sorted or deleted.
 */
function findHistoryRow(sheet, uid) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'hrow_' + uid;

  const cached = Number(cache.get(cacheKey) || 0);
  if (cached > 1 && cached <= sheet.getLastRow()) {
    if (String(sheet.getRange(cached, 1).getValue()) === uid) return cached;
    cache.remove(cacheKey);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === uid) {
      const row = i + 2;
      cache.put(cacheKey, String(row), 21600); // 6h
      return row;
    }
  }
  return 0;
}

/** Reads and sanitizes a visitor's stored turns. Never throws. */
function loadHistory(uid) {
  try {
    const sheet = getHistorySheet();
    const row = findHistoryRow(sheet, uid);
    if (!row) return [];

    const raw = String(sheet.getRange(row, 5).getValue() || '');
    if (!raw) return [];

    return normalizeHistory(JSON.parse(raw));
  } catch (err) {
    // A corrupt cell should degrade to a fresh conversation, not a 500.
    console.error('loadHistory failed for ' + uid + ': ' + err);
    return [];
  }
}

/** Appends one exchange to the visitor's row. Returns the new turn count. */
function appendHistory(uid, userText, assistantText) {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000);

  try {
    const sheet = getHistorySheet();
    const row = findHistoryRow(sheet, uid);
    const now = new Date();

    // Re-read inside the lock — another request may have written since.
    let history = [];
    if (row) {
      try {
        history = normalizeHistory(JSON.parse(String(sheet.getRange(row, 5).getValue() || '[]')));
      } catch (err) {
        history = [];
      }
    }

    history.push({ role: 'user', text: userText });
    history.push({ role: 'assistant', text: assistantText });
    history = trimHistory(history);

    const json = JSON.stringify(history);

    if (row) {
      sheet.getRange(row, 3, 1, 3).setValues([[now, history.length, json]]);
    } else {
      sheet.appendRow([uid, now, now, history.length, json]);
      CacheService.getScriptCache().put('hrow_' + uid, String(sheet.getLastRow()), 21600);
    }

    return history.length;
  } catch (err) {
    // Losing the transcript must not lose the answer the visitor already has.
    console.error('appendHistory failed for ' + uid + ': ' + err);
    return 0;
  } finally {
    if (gotLock) lock.releaseLock();
  }
}

/** Caps history by turn count and by characters, keeping the most recent. */
function trimHistory(history) {
  let trimmed = history.slice(-CONFIG.MAX_HISTORY_TURNS * 2);

  while (trimmed.length > 2 && JSON.stringify(trimmed).length > CONFIG.MAX_HISTORY_CHARS) {
    trimmed = trimmed.slice(2); // drop the oldest exchange
  }
  return trimmed;
}

// ────────────────────────────────────────────────────────────────────────────
// DeepSeek call
// ────────────────────────────────────────────────────────────────────────────

/**
 * Sends the conversation to DeepSeek and returns { ok, reply } or { ok, error }.
 *
 * DeepSeek speaks the OpenAI chat-completions shape: the system prompt is the
 * first entry in `messages`, and the answer comes back on choices[0].message.
 */
function callDeepSeek(messages) {
  let apiKey;
  try {
    apiKey = getApiKey();
  } catch (err) {
    console.error('Key lookup failed: ' + err);
    return { ok: false, error: 'The assistant is not configured yet.' };
  }

  if (!apiKey) {
    return {
      ok: false,
      error: 'The assistant is not configured yet (missing API key).',
    };
  }

  const payload = {
    model: CONFIG.MODEL,
    messages: messages,
    max_tokens: CONFIG.MAX_TOKENS,
    stream: false,
  };

  // deepseek-reasoner rejects sampling params — only send them for chat models.
  if (CONFIG.MODEL !== 'deepseek-reasoner') {
    payload.temperature = CONFIG.TEMPERATURE;
  }

  const response = UrlFetchApp.fetch(CONFIG.API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const raw = response.getContentText();

  if (status !== 200) {
    console.error('DeepSeek API ' + status + ': ' + raw);
    if (status === 401) return { ok: false, error: 'Assistant credentials are invalid.' };
    if (status === 402) return { ok: false, error: 'The assistant is out of credit.' };
    if (status === 429) return { ok: false, error: 'Busy right now — try again in a moment.' };
    if (status >= 500) return { ok: false, error: 'The AI service is having trouble. Try again shortly.' };
    return { ok: false, error: 'The assistant could not answer that.' };
  }

  let data;
  try {
    data = JSON.parse(raw);
  } catch (err) {
    console.error('Unparseable DeepSeek response: ' + raw);
    return { ok: false, error: 'Got a malformed answer. Try again.' };
  }

  const choice = data.choices && data.choices[0];
  const reply = choice && choice.message
    ? String(choice.message.content || '').trim()
    : '';

  if (!reply) {
    // Empty content with finish_reason 'content_filter' means it was blocked.
    if (choice && choice.finish_reason === 'content_filter') {
      return {
        ok: false,
        error: "I can't help with that one. Ask me about Prince's work instead.",
      };
    }
    return { ok: false, error: 'I came up empty on that — try rephrasing?' };
  }

  return {
    ok: true,
    // finish_reason 'length' means max_tokens cut it off mid-sentence.
    reply: choice.finish_reason === 'length' ? reply + '…' : reply,
    usage: data.usage
      ? { input: data.usage.prompt_tokens, output: data.usage.completion_tokens }
      : null,
  };
}

// ────────────────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────────────────

/**
 * Reads the API key out of the Config tab, cached briefly so a busy chat does
 * not hit the Sheets API once per message.
 */
function getApiKey() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('api_key');
  if (cached) return cached;

  const ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'No spreadsheet. Set CONFIG.SPREADSHEET_ID, or bind this script to the sheet.'
    );
  }

  const sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error('No "' + CONFIG.CONFIG_SHEET_NAME + '" tab in the spreadsheet.');
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === CONFIG.KEY_NAME) {
      const value = String(rows[i][1] || '').trim();
      if (value) cache.put('api_key', value, CONFIG.KEY_CACHE_SECONDS);
      return value;
    }
  }

  return '';
}

/** Call after changing the key in the sheet so the change takes effect now. */
function clearKeyCache() {
  CacheService.getScriptCache().remove('api_key');
  console.log('Key cache cleared.');
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Keeps only well-formed turns and trims to the cap. */
function normalizeHistory(raw) {
  if (!Array.isArray(raw)) return [];

  const cleaned = [];
  for (let i = 0; i < raw.length; i++) {
    const turn = raw[i];
    if (!turn || (turn.role !== 'user' && turn.role !== 'assistant')) continue;
    const text = String(turn.text || '').slice(0, CONFIG.MAX_MESSAGE_CHARS);
    if (!text) continue;
    cleaned.push({ role: turn.role, text: text });
  }

  // The first turn after the system prompt should be from the user.
  while (cleaned.length && cleaned[0].role !== 'user') cleaned.shift();

  return cleaned.slice(-CONFIG.MAX_HISTORY_TURNS * 2);
}

/**
 * Two counters in CacheService: one per visitor, one global. Cache entries
 * expire on their own, so the window slides without any cleanup.
 */
function checkRateLimit(visitorId) {
  const cache = CacheService.getScriptCache();

  const globalKey = 'rl_global';
  const globalCount = Number(cache.get(globalKey) || 0);
  if (globalCount >= CONFIG.MAX_MSGS_GLOBAL) {
    return { allowed: false, reason: 'The assistant is at capacity right now. Try again later.' };
  }

  const visitorKey = 'rl_' + visitorId;
  const visitorCount = Number(cache.get(visitorKey) || 0);
  if (visitorCount >= CONFIG.MAX_MSGS_PER_VISITOR) {
    return { allowed: false, reason: "You've hit the message limit for now — check back in a bit." };
  }

  cache.put(globalKey, String(globalCount + 1), CONFIG.RATE_WINDOW_SECONDS);
  cache.put(visitorKey, String(visitorCount + 1), CONFIG.RATE_WINDOW_SECONDS);
  return { allowed: true };
}

// ────────────────────────────────────────────────────────────────────────────
// Run these from the Apps Script editor
// ────────────────────────────────────────────────────────────────────────────

/** Confirms the sheet lookup works without printing the key. */
function testKeyLookup() {
  const key = getApiKey();
  console.log(key
    ? 'Key found (' + key.length + ' chars, starts "' + key.slice(0, 6) + '…").'
    : 'No key found — check the Config tab.');
}

/**
 * Full-path smoke test: rate limit -> load history -> model -> save history.
 * Run this, then look for the "editor-test" row in the Chat_History tab.
 */
function testChat() {
  const out = handleChat(null, {
    action: 'chat',
    message: 'What has Prince built with AI?',
    visitorId: 'editor-test',
  });
  console.log(out.getContent());
  console.log('Check the "' + CONFIG.HISTORY_SHEET_NAME +
    '" tab for a row with uid "editor-test".');
}

/** Model call only — no history read or write. Use to isolate API problems. */
function testModelOnly() {
  const result = callDeepSeek([
    { role: 'system', content: buildSystemPrompt() },
    { role: 'user', content: 'What has Prince built with AI?' },
  ]);
  console.log(JSON.stringify(result, null, 2));
}

/** Prints what is currently stored for one visitor. */
function testStoredHistory() {
  const history = loadHistory('editor-test');
  console.log(history.length + ' turns stored');
  console.log(JSON.stringify(history, null, 2));
}
