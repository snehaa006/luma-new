const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageBreak, TableLayoutType, VerticalAlign
} = require('docx');
const fs = require('fs');

// Monochrome palette — black, white and greys only.
const NAVY = '000000';   // headings / emphasis
const TEAL = '000000';   // rules, sub-headings
const GREY = '1A1A1A';   // body text
const LIGHT = 'E4E4E4';  // strong highlight band
const BAND = 'F2F2F2';   // alternating row band

const PW = 9360; // usable width in DXA for A4 with 1" margins approx

function P(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align,
    spacing: { before: opts.before ?? 60, after: opts.after ?? 120, line: 276 },
    indent: opts.indent,
    border: opts.border,
    shading: opts.shading,
    children: [new TextRun({
      text,
      size: opts.size ?? 21,
      bold: opts.bold,
      italics: opts.italics,
      color: opts.color ?? GREY,
      font: 'Calibri'
    })]
  });
}

function RICH(runs, opts = {}) {
  return new Paragraph({
    alignment: opts.align,
    spacing: { before: opts.before ?? 60, after: opts.after ?? 120, line: 276 },
    indent: opts.indent,
    children: runs.map(r => new TextRun({
      text: r.t, bold: r.b, italics: r.i, size: r.size ?? 21,
      color: r.c ?? GREY, font: 'Calibri'
    }))
  });
}

function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 4 } },
    children: [new TextRun({ text, size: 28, bold: true, color: NAVY, font: 'Calibri' })]
  });
}

function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, size: 23, bold: true, color: TEAL, font: 'Calibri' })]
  });
}

function H3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, size: 21, bold: true, color: NAVY, font: 'Calibri' })]
  });
}

function BUL(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { before: 40, after: 60, line: 276 },
    children: [new TextRun({ text, size: 21, color: GREY, font: 'Calibri' })]
  });
}

function BULR(runs, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { before: 40, after: 60, line: 276 },
    children: runs.map(r => new TextRun({ text: r.t, bold: r.b, italics: r.i, size: 21, color: r.c ?? GREY, font: 'Calibri' }))
  });
}

function NUM(text) {
  return new Paragraph({
    numbering: { reference: 'nums', level: 0 },
    spacing: { before: 40, after: 60, line: 276 },
    children: [new TextRun({ text, size: 21, color: GREY, font: 'Calibri' })]
  });
}

function cell(text, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: (Array.isArray(text) ? text : [text]).map(t => new Paragraph({
      alignment: o.align,
      spacing: { before: 20, after: 20, line: 260 },
      children: [new TextRun({
        text: t, bold: o.bold, italics: o.italics, size: o.size ?? 19,
        color: o.color ?? (o.head ? 'FFFFFF' : GREY), font: 'Calibri'
      })]
    }))
  });
}

function TBL(widths, rows) {
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: '9A9A9A' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: '9A9A9A' },
      left: { style: BorderStyle.SINGLE, size: 4, color: '9A9A9A' },
      right: { style: BorderStyle.SINGLE, size: 4, color: '9A9A9A' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: '9A9A9A' },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: '9A9A9A' }
    },
    rows
  });
}

function headRow(widths, labels) {
  return new TableRow({
    tableHeader: true,
    children: labels.map((l, i) => cell(l, widths[i], { fill: NAVY, head: true, bold: true }))
  });
}

function bodyRow(widths, vals, o = {}) {
  return new TableRow({
    children: vals.map((v, i) => cell(v, widths[i], { fill: o.fill, bold: i === 0 && o.boldFirst }))
  });
}

const SPACER = new Paragraph({ spacing: { after: 80 }, children: [] });

function calloutRow(title, body) {
  return new Table({
    columnWidths: [PW],
    width: { size: PW, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: TEAL },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL },
      left: { style: BorderStyle.SINGLE, size: 18, color: TEAL },
      right: { style: BorderStyle.SINGLE, size: 4, color: TEAL },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE }
    },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: PW, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: LIGHT, color: 'auto' },
        margins: { top: 140, bottom: 140, left: 180, right: 180 },
        children: [
          new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: title, bold: true, size: 21, color: NAVY, font: 'Calibri' })] }),
          new Paragraph({ spacing: { after: 0, line: 276 }, children: [new TextRun({ text: body, size: 21, color: GREY, font: 'Calibri' })] })
        ]
      })]
    })]
  });
}

/* ============================ CONTENT ============================ */
const children = [];

/* ---------- COVER ---------- */
children.push(
  new Paragraph({ spacing: { before: 1400, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'PROJECT PROPOSAL  |  STAGE 1: OPEN APPLICATIONS', size: 20, bold: true, color: TEAL, font: 'Calibri', characterSpacing: 60 })] }),
  new Paragraph({ spacing: { before: 300, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'LUMA', size: 84, bold: true, color: NAVY, font: 'Calibri' })] }),
  new Paragraph({ spacing: { before: 60, after: 200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Smart Accessibility Suite', size: 40, bold: true, color: TEAL, font: 'Calibri' })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 400 },
    border: { top: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 8 } },
    children: [new TextRun({
      text: 'An AI-Powered Smart Cane and Voice-First Browser Assistant that Restores Independent Navigation — Physical and Digital — for Blind and Visually Impaired Users',
      size: 24, italics: true, color: GREY, font: 'Calibri' })] })
);

const cw = [2600, 6760];
children.push(TBL(cw, [
  bodyRow(cw, ['Project / Product Title', 'Luma — Smart Accessibility Suite (AI Smart Cane + Luma Voice Extension)'], { fill: BAND, boldFirst: true }),
  bodyRow(cw, ['Domain', 'Assistive Technology · IoT · Embedded Systems · Applied AI · Web Accessibility'], { boldFirst: true }),
  bodyRow(cw, ['Primary Beneficiaries', 'Blind, low-vision and mobility-impaired individuals; elderly users; caregivers'], { fill: BAND, boldFirst: true }),
  bodyRow(cw, ['Current Stage', 'Working prototype — Chrome extension (MV3) + ThingSpeak-linked IoT dashboard with live Gemini vision scan'], { boldFirst: true }),
  bodyRow(cw, ['Repository', 'github.com/snehaa006/luma-new'], { fill: BAND, boldFirst: true }),
  bodyRow(cw, ['Licence / Access Model', 'Open source · Free forever · No accounts, no subscriptions, no data collection'], { boldFirst: true })
]));

children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------- 0. EXECUTIVE SUMMARY ---------- */
children.push(H1('Executive Summary'));
children.push(P('Luma is a single accessibility platform that solves two halves of the same problem for a blind or visually impaired person: moving safely through the physical world, and operating independently in the digital one. Existing solutions treat these as separate markets — a white cane or an electronic travel aid on one side, a screen reader on the other — leaving the user to buy, learn and maintain two unrelated systems, neither of which talks to the other.'));
children.push(P('Luma unifies them. The hardware layer is an AI-augmented smart cane: an ultrasonic obstacle sensor, a rain/wet-surface sensor, a GPS module, a haptic/servo-driven feedback actuator, a buzzer and a one-press SOS button, all publishing live telemetry to the cloud. The intelligence layer converts that raw telemetry into spoken, human-meaningful guidance — and when the user asks, a camera frame is sent to a multimodal AI model that returns a structured, clock-face situational-awareness report ("Direct path obstructed: parked scooter at 1 o’clock. Surface: wet tiles. Entrance ahead has three steps, no ramp."). The software layer is the Luma Voice Extension, a Chrome/Chromium extension that gives the same user complete hands-free, eyes-free control of any website through natural voice commands.'));
children.push(P('Both halves already exist as a working prototype in the repository accompanying this proposal. This document sets out the problem, the solution architecture, the deployment plan, the safety and efficiency case, the scaling roadmap, the innovation advantage, and a before/after user-journey map showing the measurable change in a beneficiary’s day.'));

children.push(calloutRow('The one-line thesis',
  'Independence for a blind user is not one product — it is continuity. Luma is the first free, open-source system that carries a single voice-first assistant across the pavement, the doorway and the browser tab without the user ever changing tools.'));

children.push(new Paragraph({ children: [new PageBreak()] }));

/* ---------- 1. TITLE ---------- */
children.push(H1('1. Title of the Project / Product'));
children.push(RICH([
  { t: 'LUMA — Smart Accessibility Suite: ', b: true, c: NAVY, size: 24 },
  { t: 'An AI-Powered Smart Cane with Real-Time Environmental Vision, Integrated with a Voice-First Browser Assistant for Blind and Visually Impaired Users.', size: 24 }
]));
children.push(H2('Sub-systems covered by this title'));
const t1w = [2400, 3400, 3560];
children.push(TBL(t1w, [
  headRow(t1w, ['Sub-system', 'What it is', 'What it delivers to the user']),
  bodyRow(t1w, ['Luma Smart Cane', 'Microcontroller-based cane with ultrasonic, rain, GPS, servo, buzzer and SOS button', 'Detects obstacles and hazards before contact; raises alerts; broadcasts location on emergency'], { fill: BAND }),
  bodyRow(t1w, ['Luma Virtual Eyes', 'Camera + multimodal AI (Gemini vision) scan triggered on demand or on proximity', 'Speaks a ≤50-word structured description of the path, surface and barriers ahead']),
  bodyRow(t1w, ['Luma IoT Dashboard', 'Live web dashboard over ThingSpeak with 7 telemetry channels and history charts', 'Lets caregivers and institutions monitor the user, view SOS location on a map, and audit trends'], { fill: BAND }),
  bodyRow(t1w, ['Luma Voice Extension', 'Chrome Manifest V3 extension using the Web Speech API', 'Full hands-free control of any website: reading, navigation, clicking, media, search'])
]));

/* ---------- 2. PROBLEM ---------- */
children.push(H1('2. Problem Statement'));
children.push(calloutRow('Problem Statement (single statement)',
  'Blind and visually impaired people are denied independence twice over — outdoors, because the traditional white cane detects only what it physically touches at ground level and gives no warning of overhead, waist-height, wet-surface or approaching hazards, and indoors, because the modern web remains effectively unusable without sighted help or an expensive, steep-learning-curve screen reader — and no affordable solution today addresses both halves of that dependence in one continuous, voice-first system.'));

children.push(H2('2.1 Why this problem matters'));
children.push(BUL('Roughly 2.2 billion people worldwide live with a vision impairment; around 43 million are blind, and India alone accounts for a very large share of that population. The overwhelming majority live in low- and middle-income settings where assistive technology is a luxury purchase.'));
children.push(BUL('The white cane — the dominant mobility aid, essentially unchanged for eighty years — is a contact sensor. It finds a kerb only by striking it, and it is blind to anything above knee height: open shutters, low signboards, parked two-wheelers, hanging cables, protruding scaffolding. Head and torso injuries from these are routine and under-reported.'));
children.push(BUL('Wet or slippery ground is invisible to a cane until the user is already on it. Falls are a leading cause of injury and of the loss of confidence that makes people stop going out at all.'));
children.push(BUL('When something goes wrong — a fall, disorientation, a medical event — the user has no fast, hands-free way to tell anyone where they are.'));
children.push(BUL('Digitally, the picture is no better. Commercial screen readers cost as much as a month’s income, take months to master, and still break on dynamic, JavaScript-heavy pages. Ordinary tasks — searching for a video, playing it, adjusting volume, reading an article — become multi-step keyboard ordeals or simply require a sighted person.'));
children.push(BUL('The result is a dependence loop: the user needs a companion to go out, and a companion to go online. Independence is lost not to blindness but to the absence of appropriate tooling.'));

children.push(H2('2.2 Gap in existing solutions'));
const t2w = [2500, 3300, 3560];
children.push(TBL(t2w, [
  headRow(t2w, ['Existing option', 'What it does', 'Where it falls short']),
  bodyRow(t2w, ['Traditional white cane', 'Ground-level contact detection', 'No overhead/waist hazards, no wet-surface warning, no SOS, no location, no digital access'], { fill: BAND }),
  bodyRow(t2w, ['Commercial smart canes (e.g. ultrasonic buzz canes)', 'Vibration on proximity', 'Expensive; a beep tells the user "something is there" but never what it is or where to go; no cloud, no caregiver link']),
  bodyRow(t2w, ['Screen readers (JAWS, NVDA, VoiceOver)', 'Reads screen content', 'Costly or steep learning curve; keyboard-centric; struggles with dynamic content; entirely separate from mobility'], { fill: BAND }),
  bodyRow(t2w, ['Phone AI vision apps', 'Describes a photo on request', 'Requires holding and aiming a phone while also holding a cane; no obstacle sensing, no emergency channel, not integrated with browsing']),
  bodyRow(t2w, ['Luma', 'Cane + Virtual Eyes + caregiver dashboard + voice browser control', 'One system, one wake word, one voice, free and open source'], { fill: LIGHT })
]));

/* ---------- 3. SOLUTION ---------- */
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1('3. Solution Proposed'));
children.push(P('Luma is a two-layer, voice-first accessibility suite bound together by a single conversational identity. The user says "Hey Luma" whether they are standing at a kerb or sitting at a laptop, and receives the same calm, spoken response. Everything below is implemented in the working prototype.'));

children.push(H2('3.1 Layer 1 — The Luma Smart Cane (physical navigation)'));
children.push(H3('Sensing and actuation'));
children.push(BULR([{ t: 'Ultrasonic obstacle sensor: ', b: true, c: NAVY }, { t: 'continuously measures clear distance ahead in centimetres. Below a configurable threshold (20 cm in the current build) the system escalates: buzzer tone, actuator feedback, and a spoken prompt to run an environmental scan.' }]));
children.push(BULR([{ t: 'Rain / wet-surface sensor: ', b: true, c: NAVY }, { t: 'reports an analog moisture value so the user is warned about slippery ground and rain before stepping into it — the single most common cause of falls.' }]));
children.push(BULR([{ t: 'Servo-driven feedback actuator: ', b: true, c: NAVY }, { t: 'provides silent, directional physical feedback in the grip, so guidance still reaches the user in noisy traffic or when they have chosen not to wear earphones.' }]));
children.push(BULR([{ t: 'Buzzer: ', b: true, c: NAVY }, { t: 'graded audible alerting for imminent hazards, independent of the phone or network.' }]));
children.push(BULR([{ t: 'GPS module: ', b: true, c: NAVY }, { t: 'streams latitude and longitude at seven-decimal precision for live location and route history.' }]));
children.push(BULR([{ t: 'One-press SOS button: ', b: true, c: NAVY }, { t: 'a single physical press raises a high-visibility emergency alert on every connected caregiver dashboard, carrying a one-click Google Maps link to the user’s exact coordinates.' }]));

children.push(H3('Connectivity and telemetry'));
children.push(P('The cane’s microcontroller (ESP32-class, Wi-Fi/GSM capable) publishes to a ThingSpeak IoT channel on seven fields, which the Luma dashboard polls every five seconds:'));
const t3w = [1300, 2900, 5160];
children.push(TBL(t3w, [
  headRow(t3w, ['Channel', 'Signal', 'Use in the system']),
  bodyRow(t3w, ['field1', 'Ultrasonic distance (cm)', 'Proximity alert; auto-prompts an AI environmental scan when < 20 cm'], { fill: BAND }),
  bodyRow(t3w, ['field2', 'Rain sensor (analog)', 'Wet / slippery surface warning and weather-aware routing']),
  bodyRow(t3w, ['field3', 'Emergency button state', 'Triggers the SOS banner and location broadcast to caregivers'], { fill: BAND }),
  bodyRow(t3w, ['field4 / field5', 'Latitude / Longitude', 'Live position, map link, and travel-history reconstruction']),
  bodyRow(t3w, ['field6', 'Buzzer state', 'Confirms the alert actually fired — closed-loop verification, not assumption'], { fill: BAND }),
  bodyRow(t3w, ['field7', 'Servo / actuator state', 'Confirms haptic feedback delivery and supports fault diagnosis'])
]));

children.push(H3('Virtual Eyes — on-demand AI environmental scan'));
children.push(P('The defining feature of the cane layer. A camera frame is captured (front or rear facing, switchable) and sent to a multimodal vision model with a purpose-written situational-awareness prompt. The model is instructed to answer in a fixed, safety-oriented structure and in under fifty complete words, so the user gets a usable instruction rather than a paragraph of prose:'));
children.push(NUM('Direct path — "Clear" or "Obstructed", naming the obstacle and its clock-face position (e.g. "cyclist at 2 o’clock").'));
children.push(NUM('Surface — ground type and hazards such as puddles, potholes, kerbs or steps.'));
children.push(NUM('Landmarks and barriers — entrances, and specifically entrances with stairs but no ramp.'));
children.push(P('The reply is spoken aloud immediately through speech synthesis. The implementation cascades through a list of vision models and degrades gracefully — invalid key, quota exhaustion and network failure each produce a distinct spoken message rather than silence, because for a blind user an unexplained silence is itself a safety failure.'));

children.push(H2('3.2 Layer 2 — The Luma Voice Extension (digital navigation)'));
children.push(P('A Chrome Manifest V3 extension (content script + background service worker) that runs on any website and turns the entire browser into a voice interface. It is currently implemented with more than fifty distinct spoken commands.'));
children.push(BULR([{ t: 'Wake-word activation: ', b: true, c: NAVY }, { t: 'a always-on lightweight recognizer listens for "Hey Luma" and a set of fuzzy variants matched by regular expression, so mispronunciation, accent and background noise do not lock the user out.' }]));
children.push(BULR([{ t: 'Multi-command mode: ', b: true, c: NAVY }, { t: 'after waking, Luma stays active for 30 seconds and re-arms the window after each command, so a user can chain "read headings — next heading — click Contact" without repeating the wake word.' }]));
children.push(BULR([{ t: 'Page reading: ', b: true, c: NAVY }, { t: '"read page", "read headings", "read links", "read selection", "summarize", "where am I" — content is extracted, cleaned and spoken.' }]));
children.push(BULR([{ t: 'Structural navigation: ', b: true, c: NAVY }, { t: '"next heading", "previous heading", "next link", "previous link" walk the document’s semantic structure, and "click [text]" activates any element by the words the user actually heard.' }]));
children.push(BULR([{ t: 'Media and YouTube control: ', b: true, c: NAVY }, { t: '"search for cats", "play first video", "pause", "forward 30", "volume up", "mute", "fullscreen" — an entire media session without sight or a mouse.' }]));
children.push(BULR([{ t: 'Browser control: ', b: true, c: NAVY }, { t: 'scrolling, top/bottom, back/forward, refresh, "open [any website]", zoom in/out for low-vision users who retain partial sight, plus time and date.' }]));
children.push(BULR([{ t: 'Interruptibility: ', b: true, c: NAVY }, { t: '"stop" cancels speech instantly and "go to sleep" deactivates the assistant — the user is always in control of the voice, never trapped by it.' }]));
children.push(BULR([{ t: 'Keyboard parity: ', b: true, c: NAVY }, { t: 'Ctrl+Shift+S (Cmd+Shift+S on macOS) toggles the assistant for users who prefer or need a silent trigger.' }]));
children.push(BULR([{ t: 'Accessible status surface: ', b: true, c: NAVY }, { t: 'an on-page status bar carries role="status" and aria-live="polite", so Luma cooperates with an existing screen reader instead of fighting it.' }]));

children.push(H2('3.3 Technology stack'));
const t4w = [2400, 6960];
children.push(TBL(t4w, [
  headRow(t4w, ['Layer', 'Technologies']),
  bodyRow(t4w, ['Embedded / hardware', 'ESP32-class microcontroller (Wi-Fi); HC-SR04 ultrasonic sensor; rain/moisture sensor; NEO-6M GPS; SG90 servo actuator; piezo buzzer; momentary SOS switch; Li-ion power with charge management'], { fill: BAND, boldFirst: true }),
  bodyRow(t4w, ['IoT / cloud', 'ThingSpeak channel API over HTTPS; 5-second polling; 7 telemetry fields; 20-point rolling history buffer'], { boldFirst: true }),
  bodyRow(t4w, ['Applied AI', 'Google Gemini multimodal vision API with a structured situational-awareness prompt, temperature 0.1 for factual determinism, and a multi-model fallback chain'], { fill: BAND, boldFirst: true }),
  bodyRow(t4w, ['Dashboard', 'HTML5, CSS3, vanilla JavaScript, Chart.js time-series charts, MediaDevices getUserMedia camera capture, Geolocation, Google Maps deep-linking, Web Speech Synthesis'], { boldFirst: true }),
  bodyRow(t4w, ['Browser extension', 'Chrome Manifest V3, service worker, content scripts, Web Speech API (SpeechRecognition + SpeechSynthesis), Chrome Commands API, ARIA live regions'], { fill: BAND, boldFirst: true }),
  bodyRow(t4w, ['Privacy', 'On-device browser speech processing; API keys stored only in the user’s local storage; no accounts; no audio retained or transmitted by the extension', ], { boldFirst: true })
]));

children.push(H2('3.4 End-to-end mechanism'));
children.push(NUM('Sensors on the cane sample the environment continuously and publish to the ThingSpeak channel.'));
children.push(NUM('The Luma dashboard polls that channel every five seconds and renders live values plus rolling history for all seven signals.'));
children.push(NUM('When clear distance falls below the safety threshold, the system escalates automatically: buzzer, actuator feedback, and a spoken prompt inviting an environmental scan.'));
children.push(NUM('On scan, a camera frame goes to the vision model, which returns the structured path/surface/barrier report; it is spoken aloud within seconds.'));
children.push(NUM('If the user presses SOS, the emergency banner fires on every caregiver dashboard with a live map link to the user’s coordinates.'));
children.push(NUM('Indoors, the same user says "Hey Luma" to a browser and controls reading, navigation, search and media entirely by voice — the same wake word, the same voice, zero context switching.'));

/* ---------- 4. IMPLEMENTATION ---------- */
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1('4. Implementation Plan'));
children.push(H2('4.1 Target beneficiaries'));
const t5w = [2600, 6760];
children.push(TBL(t5w, [
  headRow(t5w, ['Beneficiary group', 'How Luma serves them']),
  bodyRow(t5w, ['Blind and low-vision individuals (primary)', 'Independent outdoor mobility with pre-contact hazard warning, plus independent access to the web without sighted help'], { fill: BAND, boldFirst: true }),
  bodyRow(t5w, ['Elderly users with reduced vision or balance', 'Fall prevention through wet-surface and obstacle alerts; one-press SOS with automatic location'], { boldFirst: true }),
  bodyRow(t5w, ['Families and caregivers', 'Live dashboard, movement history and instant emergency alerting — reassurance without surveillance-style intrusion'], { fill: BAND, boldFirst: true }),
  bodyRow(t5w, ['Schools for the blind and rehabilitation centres', 'A low-cost, open, teachable platform for orientation-and-mobility training and digital-literacy classes'], { boldFirst: true }),
  bodyRow(t5w, ['NGOs and government disability programmes', 'A deployable, auditable, free solution that fits assistive-technology distribution schemes'], { fill: BAND, boldFirst: true }),
  bodyRow(t5w, ['Temporarily impaired and situationally limited users', 'Post-surgery patients, users with motor limitations, and anyone whose hands are occupied benefit from full voice control'], { boldFirst: true })
]));

children.push(H2('4.2 Phased deployment roadmap'));
const t6w = [1500, 1500, 3200, 3160];
children.push(TBL(t6w, [
  headRow(t6w, ['Phase', 'Timeline', 'Activities', 'Milestone / exit criterion']),
  bodyRow(t6w, ['Phase 1 — Prototype hardening', 'Months 1–2', 'Consolidate cane firmware; calibrate sensor thresholds; finalise enclosure and ergonomics; battery-life optimisation; publish the extension on the Chrome Web Store', 'Field-ready cane unit; extension installable in one click'], { fill: BAND }),
  bodyRow(t6w, ['Phase 2 — Supervised pilot', 'Months 3–4', 'Deploy 25–50 units at a partner school for the blind and a rehabilitation centre; supervised route trials; caregiver dashboard onboarding', 'Structured usability, latency and false-alert data from real users']),
  bodyRow(t6w, ['Phase 3 — Iterate on evidence', 'Months 5–6', 'Retune thresholds and prompt wording from pilot data; add regional-language voice output; accessibility audit with blind testers as co-designers', 'Version 2 release incorporating user-authored changes'], { fill: BAND }),
  bodyRow(t6w, ['Phase 4 — Institutional rollout', 'Months 7–9', 'Partner with NGOs, CSR programmes and state disability departments; train-the-trainer kits; local assembly and repair documentation', '500+ units in the field with a local support chain']),
  bodyRow(t6w, ['Phase 5 — Open scale', 'Months 10–12', 'Full open-source hardware BOM and firmware release; community contribution process; multi-city expansion', 'Self-sustaining community deployment beyond the founding team'], { fill: BAND })
]));

children.push(H2('4.3 Expected outcomes'));
children.push(BUL('A measurable reduction in obstacle-contact incidents and slip/fall events for pilot users, evidenced by pre- and post-deployment incident logs.'));
children.push(BUL('A measurable increase in independent trips taken without a sighted companion — the clearest single indicator of restored autonomy.'));
children.push(BUL('Emergency response time cut from "however long until someone notices" to seconds, via one-press SOS carrying exact coordinates.'));
children.push(BUL('Independent completion of everyday web tasks — reading news, searching and playing a video, filling a form — with no sighted assistance and no paid screen-reader licence.'));
children.push(BUL('A reduction in caregiver hours spent on constant physical supervision, replaced by lightweight remote reassurance.'));
children.push(BUL('An open, documented, reproducible assistive platform that any institution can build, audit and adapt without paying licence fees.'));

children.push(H2('4.4 Deliverables'));
const t7w = [3000, 6360];
children.push(TBL(t7w, [
  headRow(t7w, ['Deliverable', 'Description']),
  bodyRow(t7w, ['Luma Smart Cane unit', 'Assembled hardware with calibrated ultrasonic, rain, GPS, servo, buzzer and SOS subsystems'], { fill: BAND, boldFirst: true }),
  bodyRow(t7w, ['Cane firmware', 'Documented microcontroller source publishing all seven telemetry fields, with offline alerting that works without network'], { boldFirst: true }),
  bodyRow(t7w, ['Luma IoT Dashboard', 'Live monitoring web app: real-time cards, seven history charts, SOS banner, map link, camera scan panel'], { fill: BAND, boldFirst: true }),
  bodyRow(t7w, ['Virtual Eyes module', 'Camera capture plus structured vision-AI environmental description with spoken output and graceful failure handling'], { boldFirst: true }),
  bodyRow(t7w, ['Luma Voice Extension', 'Published Chrome/Edge/Brave extension with 50+ voice commands and full command reference'], { fill: BAND, boldFirst: true }),
  bodyRow(t7w, ['Documentation set', 'Assembly guide, bill of materials, installation walkthrough, spoken-command reference, and an audio-first quick-start for blind users'], { boldFirst: true }),
  bodyRow(t7w, ['Training kit', 'Train-the-trainer material for O&M instructors, caregivers and digital-literacy educators'], { fill: BAND, boldFirst: true }),
  bodyRow(t7w, ['Pilot evaluation report', 'Quantitative and qualitative findings from the supervised pilot, published openly'], { boldFirst: true })
]));

children.push(H2('4.5 Risks and mitigation'));
const t8w = [2700, 3200, 3460];
children.push(TBL(t8w, [
  headRow(t8w, ['Risk', 'Impact', 'Mitigation']),
  bodyRow(t8w, ['Network unavailable in the field', 'Cloud scan and dashboard unreachable', 'All life-safety alerting (buzzer, actuator, proximity threshold) runs on-device and never depends on connectivity; cloud features are additive, never load-bearing'], { fill: BAND }),
  bodyRow(t8w, ['AI describes the scene incorrectly', 'Misleading guidance', 'Low-temperature, tightly structured prompt; the AI is explicitly positioned as a supplement to — never a replacement for — the cane’s physical contact sensing; training material states this plainly']),
  bodyRow(t8w, ['Alert fatigue from false positives', 'User disables the device', 'Thresholds tuned on pilot data; graded escalation instead of constant beeping; user-adjustable sensitivity'], { fill: BAND }),
  bodyRow(t8w, ['Battery exhaustion mid-journey', 'Loss of assistance', 'Low-power duty cycling, spoken battery warnings well ahead of depletion, and a cane that remains a fully functional white cane at zero charge']),
  bodyRow(t8w, ['Location privacy concerns', 'Reluctance to adopt', 'User-controlled sharing, no third-party data sale, local key storage, and an open codebase anyone can audit'], { fill: BAND }),
  bodyRow(t8w, ['Speech recognition accuracy across accents', 'Commands missed', 'Fuzzy wake-phrase matching, multi-variant command synonyms, automatic recognizer restart on failure, and a keyboard fallback shortcut'])
]));

/* ---------- 5. SAFETY & EFFICIENCY ---------- */
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1('5. Safety and Efficiency Analysis'));
children.push(H2('5.1 Safety architecture'));
children.push(P('Luma is designed around one principle: safety functions must degrade, never disappear. The system is layered so that the loss of any outer layer still leaves the user protected by the one beneath it.'));
const t9w = [1500, 2600, 5260];
children.push(TBL(t9w, [
  headRow(t9w, ['Layer', 'Mechanism', 'Safety guarantee']),
  bodyRow(t9w, ['Layer 0', 'Physical cane', 'Remains a fully functional white cane with zero power, zero network and zero software'], { fill: BAND }),
  bodyRow(t9w, ['Layer 1', 'On-device sensing and alerting', 'Ultrasonic proximity, wet-surface detection, buzzer and actuator all operate locally; no internet required for life-safety alerts']),
  bodyRow(t9w, ['Layer 2', 'Cloud telemetry and caregiver dashboard', 'Remote awareness, SOS broadcast with coordinates, and auditable history'], { fill: BAND }),
  bodyRow(t9w, ['Layer 3', 'AI Virtual Eyes', 'Rich contextual description on demand — additive intelligence, explicitly not a substitute for physical sensing'])
]));

children.push(H2('5.2 How safety is actively improved'));
children.push(BULR([{ t: 'Pre-contact hazard detection. ', b: true, c: NAVY }, { t: 'The fundamental limitation of a white cane is that it only knows about an obstacle once the user has struck it. Ultrasonic sensing moves detection ahead of contact, and vision-based scanning identifies hazards a cane can never reach: overhanging signage, open shutters, low branches, parked vehicles, and approaching people.' }]));
children.push(BULR([{ t: 'Fall prevention. ', b: true, c: NAVY }, { t: 'Wet and slippery surfaces are announced before the user steps onto them, and the AI report explicitly calls out puddles, steps and kerbs. Falls are the highest-consequence everyday risk for this population, and they are almost entirely preventable with advance warning.' }]));
children.push(BULR([{ t: 'Named, positional information instead of anonymous beeping. ', b: true, c: NAVY }, { t: 'A beep tells a user to freeze. "Parked scooter at 1 o’clock, clear to your left" tells them what to do. Clock-face positioning is the standard orientation-and-mobility convention, so the output is immediately actionable by anyone with basic O&M training.' }]));
children.push(BULR([{ t: 'Emergency escalation without sight or dexterity. ', b: true, c: NAVY }, { t: 'One physical press — no menu, no unlocking a phone, no aiming a camera — raises an alert with exact coordinates and a direct map link on every caregiver dashboard.' }]));
children.push(BULR([{ t: 'Closed-loop actuator verification. ', b: true, c: NAVY }, { t: 'Buzzer and servo states are themselves telemetry fields. The system does not assume an alert fired; it confirms it. A silent failure of an alerting device is the most dangerous fault an assistive system can have, and Luma detects it.' }]));
children.push(BULR([{ t: 'Multi-modal redundancy. ', b: true, c: NAVY }, { t: 'Every alert reaches the user through more than one channel — audible buzzer, physical actuator feedback, and spoken description — so a user in traffic noise, wearing gloves, or with a hearing impairment still receives the warning.' }]));
children.push(BULR([{ t: 'Fail-loud, never fail-silent. ', b: true, c: NAVY }, { t: 'Invalid credentials, exhausted quota and network errors each produce a distinct spoken message. The user is never left interpreting silence as "the path is clear".' }]));
children.push(BULR([{ t: 'Digital safety. ', b: true, c: NAVY }, { t: 'Speech is processed by the browser’s built-in engine; the extension transmits and stores no audio, holds no accounts, and keeps any API key in the user’s own local storage. The full source is open for audit — a meaningful safety property for a device that knows where a vulnerable person is.' }]));

children.push(H2('5.3 Efficiency gains — including material and object handling'));
children.push(P('"Efficiency" for a blind user is measured in a currency sighted design rarely counts: the number of seconds, steps, retries and requests for help needed to complete an ordinary task. Luma attacks all four.'));
const t10w = [2600, 2200, 2200, 2360];
children.push(TBL(t10w, [
  headRow(t10w, ['Task', 'Conventional method', 'With Luma', 'Efficiency effect']),
  bodyRow(t10w, ['Detecting an obstacle ahead', 'Sweep and strike it with the cane', 'Announced before contact, with position', 'Removes stop-strike-reorient cycles; smoother, faster, safer walking'], { fill: BAND }),
  bodyRow(t10w, ['Locating and approaching an object or package', 'Repeated sweeping and groping', 'Distance readout plus a spoken description of what is ahead and where', 'Fewer approach attempts; less collision damage to the object and the user']),
  bodyRow(t10w, ['Judging whether a surface is safe to carry a load across', 'Discovered by slipping', 'Wet-surface sensing and AI surface report before stepping', 'Prevents dropped items and falls while carrying — the highest-risk handling scenario'], { fill: BAND }),
  bodyRow(t10w, ['Finding an accessible entrance', 'Trial and error along a wall', 'AI names entrances and flags stairs without ramps', 'Direct routing instead of perimeter searching']),
  bodyRow(t10w, ['Calling for help', 'Find phone, unlock, find contact, describe location', 'One button press with automatic coordinates', 'Response time reduced from minutes to seconds'], { fill: BAND }),
  bodyRow(t10w, ['Reading a web page', 'Sighted help or keyboard-driven screen reader', '"Hey Luma, read page"', 'One spoken command replaces a long navigation sequence']),
  bodyRow(t10w, ['Finding and playing a video', 'Multi-step keyboard traversal', '"Search for X" then "play first video"', 'Two commands replace dozens of keystrokes'], { fill: BAND }),
  bodyRow(t10w, ['Chaining several actions', 'Re-invoke the assistant each time', '30-second multi-command window, re-armed after each command', 'Removes repeated wake-word overhead from every workflow'])
]));

children.push(H3('Specific relevance to material and object handling'));
children.push(P('Handling objects safely — picking something up, carrying it, setting it down, moving through a space while holding it — is where vision impairment is most punishing, because both hands are committed and the cane’s sweep is compromised. Luma addresses this directly:'));
children.push(BUL('Hands-free operation throughout. Every function is voice- or single-button-driven, so the user never has to put down what they are carrying to interact with the system.'));
children.push(BUL('Advance clearance information. Knowing the clear distance ahead in centimetres lets a user carrying a load judge whether a gap is passable before committing to it, instead of discovering the answer by collision.'));
children.push(BUL('Surface risk before load-bearing. Wet-floor detection matters most precisely when the user is carrying something and cannot break a fall with their hands.'));
children.push(BUL('Silent haptic channel. The servo actuator conveys warnings without sound, which is essential in a workshop, a warehouse aisle or a busy street where audio alerts would be lost or unwelcome.'));
children.push(BUL('Verified alerting. Because actuator and buzzer states are reported back, a caregiver or supervisor can confirm that safety feedback is actually reaching a user who is handling materials, rather than trusting that it is.'));
children.push(BUL('Institutional visibility. The dashboard’s history charts let a training centre or workplace identify recurring hazard points along a routine route and fix the environment, not just warn about it.'));

/* ---------- 6. SCALABILITY ---------- */
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1('6. Scalability and Future Development'));
children.push(H2('6.1 Why the architecture scales'));
children.push(BULR([{ t: 'Decoupled layers. ', b: true, c: NAVY }, { t: 'Hardware, cloud telemetry, AI and dashboard communicate over a documented channel schema. Any layer can be replaced — a different microcontroller, a different broker, a different vision model — without rewriting the others.' }]));
children.push(BULR([{ t: 'Standard, commodity IoT transport. ', b: true, c: NAVY }, { t: 'The telemetry model is a plain multi-field channel. Moving from ThingSpeak to a self-hosted MQTT broker or a national health-service backend is a configuration change, not a redesign.' }]));
children.push(BULR([{ t: 'Model-agnostic AI. ', b: true, c: NAVY }, { t: 'The vision layer already cascades through multiple models and can adopt a newer or cheaper one — including an on-device model — without touching the user experience.' }]));
children.push(BULR([{ t: 'Zero-marginal-cost software. ', b: true, c: NAVY }, { t: 'The browser extension distributes for free to unlimited users through existing web stores; the incremental cost of the ten-thousandth user is nil.' }]));
children.push(BULR([{ t: 'Low-cost, locally sourceable hardware. ', b: true, c: NAVY }, { t: 'Every component is a commonly available module, so units can be assembled and repaired regionally rather than imported.' }]));
children.push(BULR([{ t: 'Open source. ', b: true, c: NAVY }, { t: 'Scaling does not depend on the founding team’s capacity. Any institution can fork, localise and deploy.' }]));

children.push(H2('6.2 Scaling roadmap'));
const t11w = [2000, 3600, 3760];
children.push(TBL(t11w, [
  headRow(t11w, ['Horizon', 'Scale target', 'Enablers']),
  bodyRow(t11w, ['Short term (0–6 months)', 'Hundreds of users across pilot institutions', 'Chrome Web Store publication, partner schools and rehabilitation centres, train-the-trainer kits'], { fill: BAND }),
  bodyRow(t11w, ['Medium term (6–18 months)', 'Thousands of users across multiple states or countries', 'Regional-language voice packs, NGO and CSR distribution, local assembly partners, self-hosted backend option']),
  bodyRow(t11w, ['Long term (18 months +)', 'A general assistive platform beyond the cane form factor', 'Wearable and smart-glasses variants, offline on-device AI, integration with public-transport and indoor-navigation data'], { fill: BAND })
]));

children.push(H2('6.3 Planned enhancements'));
const t12w = [2800, 6560];
children.push(TBL(t12w, [
  headRow(t12w, ['Enhancement', 'Value added']),
  bodyRow(t12w, ['Multilingual and regional-language voice', 'Removes the English barrier that excludes the majority of potential users in India and comparable markets'], { fill: BAND, boldFirst: true }),
  bodyRow(t12w, ['Offline / on-device vision model', 'Environmental description without connectivity or per-query cost — critical for rural deployment'], { boldFirst: true }),
  bodyRow(t12w, ['Automatic fall detection (IMU)', 'Triggers the SOS path without requiring the user to press anything, covering the case where they cannot'], { fill: BAND, boldFirst: true }),
  bodyRow(t12w, ['Turn-by-turn voice navigation', 'Extends Luma from hazard avoidance to full destination routing'], { boldFirst: true }),
  bodyRow(t12w, ['Indoor navigation via BLE beacons', 'Hospitals, campuses, metro stations and malls, where GPS is unusable'], { fill: BAND, boldFirst: true }),
  bodyRow(t12w, ['Currency, text and face recognition', 'Reading notes, signage and labels; recognising known people — the highest-requested features in user research on vision assistance'], { boldFirst: true }),
  bodyRow(t12w, ['Voice-driven form filling and e-commerce', 'Completes digital independence: banking, government portals and shopping without sighted help'], { fill: BAND, boldFirst: true }),
  bodyRow(t12w, ['Cross-browser and mobile companion app', 'Firefox and Safari support; a phone app so the suite works away from a desktop'], { boldFirst: true }),
  bodyRow(t12w, ['Wearable form factors', 'Clip-on, wrist and smart-glasses variants for users who do not use a cane at all'], { fill: BAND, boldFirst: true }),
  bodyRow(t12w, ['Anonymised hazard mapping', 'Aggregated, consented telemetry identifies recurring accessibility black spots and gives municipalities hard evidence to fix them'], { boldFirst: true })
]));

children.push(H2('6.4 Adaptation to new environments and use cases'));
children.push(BUL('Industrial and warehouse safety: the same proximity, wet-floor and verified-haptic stack protects any worker moving materials in a low-visibility or high-noise environment.'));
children.push(BUL('Elderly care and assisted living: fall risk, wandering alerts and one-press SOS transfer directly to geriatric care with no architectural change.'));
children.push(BUL('Post-operative and temporary impairment: short-term rental units for patients recovering from eye surgery.'));
children.push(BUL('Education: the dashboard and open firmware make Luma a ready-made teaching platform for IoT, embedded systems and inclusive design.'));
children.push(BUL('Disaster and low-visibility response: proximity sensing plus location broadcast is useful to any responder operating in smoke or darkness.'));

/* ---------- 7. INNOVATION ---------- */
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1('7. Innovation Advantage'));
children.push(H2('7.1 What makes Luma genuinely new'));
children.push(BULR([{ t: '1. It unifies physical and digital accessibility in one identity. ', b: true, c: NAVY }, { t: 'Every competitor solves one half. Smart canes handle the street; screen readers handle the screen. Luma is, to our knowledge, the first free system where the same wake word, the same voice and the same mental model carry the user from the pavement to the browser tab. The cognitive cost of switching tools — which is what actually defeats adoption — is eliminated.' }]));
children.push(BULR([{ t: '2. It replaces alerting with understanding. ', b: true, c: NAVY }, { t: 'Conventional electronic travel aids output an anonymous beep or buzz proportional to proximity: the user learns that something is there, never what or where. Luma’s Virtual Eyes returns a named obstacle at a clock-face position, a surface assessment and a barrier report. That is the difference between an alarm and a guide.' }]));
children.push(BULR([{ t: '3. The AI is prompt-engineered for safety, not for eloquence. ', b: true, c: NAVY }, { t: 'The vision model is constrained to a fixed three-part structure, capped at fifty words, held at low temperature for factual determinism, and explicitly instructed to produce complete sentences — because a description truncated mid-phrase is worse than none at all for someone about to take a step. This is a deliberate safety-engineering decision, not a generic "ask the AI about the photo" integration.' }]));
children.push(BULR([{ t: '4. Actuator states are telemetry. ', b: true, c: NAVY }, { t: 'Buzzer and servo states are published and charted alongside sensor readings, so the system verifies that its own warnings were delivered. Most assistive devices assume their output worked; Luma proves it.' }]));
children.push(BULR([{ t: '5. Fail-loud design. ', b: true, c: NAVY }, { t: 'Every failure mode — bad key, exhausted quota, no network, recognizer crash — is surfaced audibly and recovered from automatically. Silence is treated as a defect, because a blind user cannot see an error message.' }]));
children.push(BULR([{ t: '6. Radically forgiving voice interaction. ', b: true, c: NAVY }, { t: 'Fuzzy regular-expression wake matching accepts a wide range of pronunciations and mishearings; commands accept multiple natural synonyms; the 30-second re-arming window removes wake-word repetition; and a keyboard shortcut covers users who cannot or prefer not to speak.' }]));
children.push(BULR([{ t: '7. It cooperates with existing assistive technology. ', b: true, c: NAVY }, { t: 'ARIA live regions and role="status" mean Luma layers on top of a user’s existing screen reader instead of competing with it — a rare and deliberate choice.' }]));
children.push(BULR([{ t: '8. Free, open source and privacy-preserving by construction. ', b: true, c: NAVY }, { t: 'No accounts, no subscriptions, no audio collection, keys held locally, and a codebase anyone can audit. For a device that knows a vulnerable person’s location, auditability is not a feature — it is a precondition for trust.' }]));
children.push(BULR([{ t: '9. Caregiver visibility without surveillance framing. ', b: true, c: NAVY }, { t: 'The dashboard exists to answer "is everything alright?", not to track. Sharing is user-controlled and emergency-centred.' }]));

children.push(H2('7.2 Competitive differentiation'));
const t13w = [2200, 1720, 1720, 1720, 2000];
children.push(TBL(t13w, [
  headRow(t13w, ['Capability', 'White cane', 'Smart canes', 'Screen readers', 'LUMA']),
  bodyRow(t13w, ['Pre-contact obstacle detection', 'No', 'Yes', 'N/A', 'Yes'], { fill: BAND }),
  bodyRow(t13w, ['Names the obstacle and its direction', 'No', 'No', 'N/A', 'Yes']),
  bodyRow(t13w, ['Wet / slippery surface warning', 'No', 'Rare', 'N/A', 'Yes'], { fill: BAND }),
  bodyRow(t13w, ['One-press SOS with live location', 'No', 'Rare', 'N/A', 'Yes']),
  bodyRow(t13w, ['Caregiver dashboard and history', 'No', 'Rare', 'No', 'Yes'], { fill: BAND }),
  bodyRow(t13w, ['Verified alert delivery', 'N/A', 'No', 'No', 'Yes']),
  bodyRow(t13w, ['Full voice control of any website', 'No', 'No', 'Partial', 'Yes'], { fill: BAND }),
  bodyRow(t13w, ['Hands-free media and YouTube control', 'No', 'No', 'Partial', 'Yes']),
  bodyRow(t13w, ['Single tool across physical and digital life', 'No', 'No', 'No', 'Yes'], { fill: BAND }),
  bodyRow(t13w, ['Cost to the user', 'Low', 'High', 'High or steep learning curve', 'Free'], { fill: LIGHT }),
  bodyRow(t13w, ['Open source and auditable', 'N/A', 'No', 'Partly', 'Yes'], { fill: LIGHT })
]));

children.push(calloutRow('The differentiator in one sentence',
  'Competing products tell a blind user that something is in the way. Luma tells them what it is, where it is, whether the ground is safe, who to call if it goes wrong — and then, when they get home, reads them the news.'));

/* ---------- 8. VISUAL IMPACT ---------- */
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1('8. Visual Impact and User Journey'));
children.push(P('The following before-and-after journey map follows one representative beneficiary — Priya, 34, blind since childhood, living independently in a mid-sized Indian city — through a single ordinary day. Each stage contrasts the current reality with the same stage using Luma.'));

children.push(H2('8.1 Before and after — a day in the life'));
const t14w = [1500, 3930, 3930];
children.push(TBL(t14w, [
  headRow(t14w, ['Stage', 'BEFORE — without Luma', 'AFTER — with Luma']),
  bodyRow(t14w, [['08:00','Leaving home'], 'Waits for a family member to be free to escort her. If nobody is available, the errand is postponed. Dependence begins before she reaches the door.', 'Picks up the Luma cane. Hears "Luma ready. Path clear ahead." Leaves alone, on her own schedule.'], { fill: BAND }),
  bodyRow(t14w, [['08:15','The pavement'], 'Sweeps the cane and strikes a parked two-wheeler at knee height. Startled, bruised, disoriented; several seconds lost re-establishing bearings.', 'At 20 cm clear distance the buzzer sounds and the grip actuator pulses. She asks for a scan: "Direct path obstructed: parked scooter at 1 o\'clock. Clear to your left." She steps around it without breaking stride.']),
  bodyRow(t14w, [['08:25','After rain'], 'Steps onto a wet tiled forecourt she cannot detect. Slips. Even without injury, confidence is damaged for weeks.', 'The rain sensor and the AI surface report warn her: "Surface: wet tiles, slippery." She slows and adjusts her grip before stepping.'], { fill: BAND }),
  bodyRow(t14w, [['08:40','The building entrance'], 'Feels along the wall for the door, then discovers three steps and no ramp only after arriving at them.', 'The scan reports "Entrance ahead, three steps, no ramp, handrail on the right." She approaches prepared and takes the correct line.']),
  bodyRow(t14w, [['11:00','An incident'], 'Trips and falls. No way to signal for help. Waits for a passer-by to notice, describes her location by guesswork, and hopes.', 'Presses the SOS button once. Her family\'s dashboard lights up with an emergency banner and a live map pin. Help is en route within seconds.'], { fill: BAND }),
  bodyRow(t14w, [['16:00','Reading the news online'], 'Either asks a sighted relative to read it aloud, or fights an expensive, keyboard-heavy screen reader that stumbles on dynamic pages.', '"Hey Luma, read headings." "Next heading." "Read page." She reads independently, at her own pace, with nobody in the room.']),
  bodyRow(t14w, [['20:00','Entertainment'], 'Asks someone else to search for a video, open it and set the volume. Simple leisure requires another person\'s time.', '"Hey Luma, search for classical music." "Play first video." "Volume up." Entirely hands-free, entirely her own.'], { fill: BAND }),
  bodyRow(t14w, [['End of day','How it felt'], 'Two people\'s day consumed. Two avoidable injuries. A quiet, corrosive sense of being a burden.', 'One independent day. Zero collisions. Zero requests for help. Autonomy, restored — and a caregiver reassured without ever intruding.'], { fill: LIGHT })
]));

children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H2('8.2 Storyboard — the six-frame hazard sequence'));
children.push(P('The core interaction loop, frame by frame, as it would be illustrated in the accompanying visual panel:'));
const t15w = [900, 2400, 3030, 3030];
children.push(TBL(t15w, [
  headRow(t15w, ['Frame', 'What happens', 'What the user experiences', 'What the system does']),
  bodyRow(t15w, ['1', 'Walking normally', 'Silence. No interruption. The system stays out of the way.', 'Sensors sample continuously; telemetry publishes to the cloud'], { fill: BAND }),
  bodyRow(t15w, ['2', 'Obstacle enters range', 'Buzzer tone plus a pulse in the grip', 'Ultrasonic distance drops below the 20 cm threshold; actuator and buzzer fire']),
  bodyRow(t15w, ['3', 'Scan invited', '"Object detected nearby. Consider scanning the environment."', 'The system prompts rather than deciding for the user — agency is preserved'], { fill: BAND }),
  bodyRow(t15w, ['4', 'Scan performed', 'A camera frame is captured hands-free', 'The frame goes to the vision model with the structured situational-awareness prompt']),
  bodyRow(t15w, ['5', 'Guidance spoken', '"Direct path obstructed: parked scooter at 1 o\'clock. Surface: dry paving. Clear to your left."', 'The reply is parsed and spoken within seconds'], { fill: BAND }),
  bodyRow(t15w, ['6', 'User proceeds', 'Steps around the obstacle confidently and walks on', 'Telemetry and the event are logged for caregiver review and hazard mapping'])
]));

children.push(H2('8.3 The transformation, summarised'));
const t16w = [3120, 3120, 3120];
children.push(TBL(t16w, [
  headRow(t16w, ['Dimension', 'Before Luma', 'After Luma']),
  bodyRow(t16w, ['Mobility', 'Contact-based; reactive; injury-prone', 'Anticipatory; informed; confident'], { fill: BAND, boldFirst: true }),
  bodyRow(t16w, ['Information available', 'Only what the cane touches', 'Obstacle identity, direction, surface, barriers'], { boldFirst: true }),
  bodyRow(t16w, ['Emergency response', 'Depends on a stranger noticing', 'One press; instant alert with coordinates'], { fill: BAND, boldFirst: true }),
  bodyRow(t16w, ['Digital access', 'Sighted help or costly screen reader', 'Free voice control of any website'], { boldFirst: true }),
  bodyRow(t16w, ['Caregiver burden', 'Constant physical accompaniment', 'Lightweight remote reassurance'], { fill: BAND, boldFirst: true }),
  bodyRow(t16w, ['Cost to the user', 'High, recurring, or simply unaffordable', 'Free and open source'], { boldFirst: true }),
  bodyRow(t16w, ['Felt experience', 'Dependence, caution, withdrawal', 'Independence, confidence, participation'], { fill: LIGHT, boldFirst: true })
]));

children.push(H2('8.4 Note on the visual submission'));
children.push(P('This section is written so that it can be laid out directly as the required visual: Section 8.1 is the before/after user-journey map (two parallel tracks across one day), Section 8.2 is the six-frame storyboard of the core hazard interaction, and Section 8.3 is the summary impact panel. In the final PDF submission these are rendered as an illustrated split-screen journey map — the "before" track in muted grey with collision and fall icons, the "after" track in Luma teal with the spoken-guidance callouts reproduced as speech bubbles.'));

/* ---------- CLOSING ---------- */
children.push(H1('Conclusion'));
children.push(P('Luma does not ask a blind person to adapt to technology. It asks technology to notice them. By combining a low-cost sensor-equipped cane, a genuinely descriptive AI vision layer, an emergency channel that works with a single press, and a browser that finally answers to a human voice, Luma removes the two dependencies that most define the daily experience of vision impairment.'));
children.push(P('Both halves are already built and working. What this proposal seeks is the opportunity to harden, pilot and distribute them — free, open, and at scale — to the people who have waited longest for them.'));
children.push(new Paragraph({
  spacing: { before: 400 }, alignment: AlignmentType.CENTER,
  border: { top: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 10 } },
  children: [new TextRun({ text: 'LUMA — Smart Accessibility Suite  ·  Open Source  ·  Free Forever  ·  Built to make the world accessible for everyone.', size: 20, italics: true, bold: true, color: NAVY, font: 'Calibri' })]
}));

/* ============================ DOCUMENT ============================ */
const doc = new Document({
  creator: 'Luma Team',
  title: 'Luma — Smart Accessibility Suite: Project Proposal',
  description: 'Stage 1 Open Applications project proposal for Luma Smart Accessibility Suite',
  styles: {
    default: {
      document: { run: { font: 'Calibri', color: GREY } },
      heading1: { run: { color: NAVY, bold: true } },
      heading2: { run: { color: NAVY, bold: true } },
      heading3: { run: { color: NAVY, bold: true } },
      heading4: { run: { color: NAVY, bold: true } },
      heading5: { run: { color: NAVY, bold: true } },
      heading6: { run: { color: NAVY, bold: true } },
      title: { run: { color: NAVY, bold: true } }
    },
    characterStyles: [
      { id: 'Hyperlink', name: 'Hyperlink', basedOn: 'DefaultParagraphFont',
        run: { color: NAVY, underline: {} } }
    ]
  },
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 460, hanging: 240 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '◦', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 880, hanging: 240 } } } }
        ]
      },
      {
        reference: 'nums',
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 460, hanging: 260 } } } }
        ]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1100, bottom: 1100, left: 1080, right: 1080 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/luma-new/docs/Luma_Project_Proposal.docx', b);
  console.log('written', b.length, 'bytes');
});
