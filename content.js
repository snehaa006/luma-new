// Luma Voice Assistant — Content Script
// Accessibility-first voice assistant for blind and visually impaired users

(function () {
  if (window._lumaInitialized) return;
  window._lumaInitialized = true;

  // --- State ---
  let isAwake = false;
  let isListening = false;
  let hasAnnouncedReady = false;
  let consecutiveFailures = 0;
  let awakeTimeout = null;
  const AWAKE_DURATION = 30000; // Stay awake for 30s after each command
  let headingIndex = -1;
  let linkIndex = -1;
  let isSpeaking = false;

  // --- Wake Phrases ---
  const wakePhrases = [
    "hey luma", "hey looma", "hey luna", "hey luma ai",
    "luma", "looma", "luna", "luma ai",
    "hey assistant", "hey voice",
    "assistant", "voice", "activate", "start", "wake up",
    "listen", "hello", "hi", "computer"
  ];

  const wakeFuzzyPatterns = [
    /hey\s*lu+ma/, /\blu+ma\b/, /\blu+na\b/, /assist/, /voice/,
    /wake/, /activate/, /listen/, /hello/, /\bhi\b/, /computer/
  ];

  let wakeRecognizer, commandRecognizer;

  // --- Status Bar ---
  function createStatusBar() {
    const bar = document.createElement("div");
    bar.id = "luma-status-bar";
    bar.setAttribute("role", "status");
    bar.setAttribute("aria-live", "polite");
    bar.setAttribute("aria-label", "Luma Voice Assistant status");
    bar.style.cssText = `
      position: fixed;
      bottom: 20px;
      left: 20px;
      z-index: 2147483647;
      padding: 12px 16px;
      background: linear-gradient(135deg, #1a1a1a 0%, #2d2d2d 100%);
      color: #00ff88;
      border-radius: 12px;
      font-family: 'Segoe UI', Arial, sans-serif;
      font-size: 14px;
      font-weight: 500;
      box-shadow: 0 8px 32px rgba(0, 255, 136, 0.3);
      border: 1px solid rgba(0, 255, 136, 0.2);
      backdrop-filter: blur(10px);
      transition: all 0.3s ease;
      min-width: 220px;
      max-width: 400px;
    `;
    bar.innerHTML = `
      <div style="display: flex; align-items: center; gap: 8px;">
        <div id="luma-icon" style="font-size: 16px;" aria-hidden="true">🤖</div>
        <div id="luma-text">Initializing...</div>
      </div>
    `;
    document.body.appendChild(bar);
    return bar;
  }

  let statusBar;

  function updateStatus(text, icon) {
    if (!statusBar) return;
    const iconEl = statusBar.querySelector("#luma-icon");
    const textEl = statusBar.querySelector("#luma-text");
    if (iconEl && icon) iconEl.textContent = icon;
    if (textEl) textEl.textContent = text;
  }

  // --- Speech Output ---
  function speak(text, callback) {
    isSpeaking = true;
    updateStatus(text.substring(0, 60) + (text.length > 60 ? "..." : ""), "🔊");
    speechSynthesis.cancel();

    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = "en-US";
    utterance.rate = 0.95;
    utterance.pitch = 1.0;
    utterance.volume = 1.0;

    utterance.onend = () => {
      isSpeaking = false;
      updateStatus(
        isAwake ? "Listening for command..." : "Say 'Hey Luma' to wake me",
        isAwake ? "🎤" : "😴"
      );
      if (callback) callback();
    };

    utterance.onerror = () => {
      isSpeaking = false;
      if (callback) callback();
    };

    speechSynthesis.speak(utterance);
  }

  function stopSpeaking() {
    speechSynthesis.cancel();
    isSpeaking = false;
  }

  // --- Wake Word Detection ---
  function checkWakePhrase(text) {
    for (const phrase of wakePhrases) {
      if (text.includes(phrase)) return true;
    }
    return wakeFuzzyPatterns.some((regex) => regex.test(text));
  }

  function createWakeRecognizer() {
    const recog = new webkitSpeechRecognition();
    recog.continuous = true;
    recog.interimResults = true;
    recog.lang = "en-US";
    recog.maxAlternatives = 5;

    recog.onstart = () => {
      isListening = true;
      consecutiveFailures = 0;
      updateStatus("Say 'Hey Luma' to wake me", "😴");
    };

    recog.onerror = (e) => {
      if (e.error === "aborted") return;
      consecutiveFailures++;
      if (consecutiveFailures < 5) {
        setTimeout(() => {
          if (!isAwake && !isListening) listenForWakeWord();
        }, 1000 * consecutiveFailures);
      }
    };

    recog.onresult = (event) => {
      for (let i = event.resultIndex; i < event.results.length; i++) {
        for (let j = 0; j < event.results[i].length; j++) {
          const text = event.results[i][j].transcript.trim().toLowerCase();
          if (checkWakePhrase(text)) {
            recog.stop();
            activateAssistant();
            return;
          }
        }
      }
    };

    recog.onend = () => {
      isListening = false;
      if (!isAwake) {
        setTimeout(() => {
          if (!isAwake && !isListening) listenForWakeWord();
        }, 300);
      }
    };

    return recog;
  }

  function activateAssistant() {
    isAwake = true;
    resetAwakeTimer();
    speak("Ready! What can I help you with?", () => {
      listenForCommand();
    });
  }

  // --- Multi-Command Mode ---
  // Stays awake for 30 seconds so blind users can chain commands without repeating the wake word
  function resetAwakeTimer() {
    clearTimeout(awakeTimeout);
    awakeTimeout = setTimeout(() => {
      if (isAwake) {
        isAwake = false;
        headingIndex = -1;
        linkIndex = -1;
        stopSpeaking();
        updateStatus("Going to sleep. Say 'Hey Luma' to wake me.", "😴");
        speak("Going to sleep.", () => {
          listenForWakeWord();
        });
      }
    }, AWAKE_DURATION);
  }

  // --- Command Recognition ---
  function createCommandRecognizer() {
    const recog = new webkitSpeechRecognition();
    recog.continuous = false;
    recog.interimResults = false;
    recog.lang = "en-US";
    recog.maxAlternatives = 3;

    recog.onstart = () => {
      isListening = true;
      consecutiveFailures = 0;
      updateStatus("Listening for command...", "🎤");
    };

    recog.onerror = (e) => {
      if (e.error === "aborted") return;
      consecutiveFailures++;
      if (consecutiveFailures < 5) {
        setTimeout(() => {
          if (!isListening) {
            if (isAwake) listenForCommand();
            else listenForWakeWord();
          }
        }, 800);
      }
    };

    recog.onresult = (event) => {
      let bestText = "";
      for (let i = 0; i < event.results[0].length; i++) {
        const alt = event.results[0][i].transcript.trim().toLowerCase();
        if (alt.length > bestText.length) bestText = alt;
      }
      updateStatus("Processing: " + bestText, "🧠");
      handleCommand(bestText);
    };

    recog.onend = () => {
      isListening = false;
      // Stay awake and keep listening if still in awake mode
      if (isAwake) {
        setTimeout(() => {
          if (isAwake && !isListening) listenForCommand();
        }, 500);
      }
    };

    return recog;
  }

  // --- Start Listeners ---
  function listenForWakeWord() {
    if (isListening || isAwake) return;
    try {
      wakeRecognizer = createWakeRecognizer();
      wakeRecognizer.start();
    } catch (e) {
      setTimeout(() => listenForWakeWord(), 2000);
    }
  }

  function listenForCommand() {
    if (isListening) return;
    try {
      commandRecognizer = createCommandRecognizer();
      commandRecognizer.start();
    } catch (e) {
      setTimeout(() => {
        if (isAwake) listenForCommand();
        else listenForWakeWord();
      }, 1000);
    }
  }

  // --- Command Handler ---
  function handleCommand(cmd) {
    cmd = cmd.toLowerCase().trim();
    resetAwakeTimer();

    // Stop / Cancel speech
    if (cmd.includes("stop") || cmd.includes("shut up") || cmd.includes("quiet") || cmd.includes("cancel")) {
      stopSpeaking();
      speak("Stopped.");
      return;
    }

    // Sleep
    if (cmd.includes("go to sleep") || cmd.includes("sleep") || cmd === "bye" || cmd === "goodbye") {
      isAwake = false;
      clearTimeout(awakeTimeout);
      speak("Going to sleep. Say 'Hey Luma' to wake me.", () => {
        listenForWakeWord();
      });
      return;
    }

    // Help
    if (cmd.includes("help") || cmd.includes("commands") || cmd.includes("what can you do")) {
      speak(
        "I can help you with: " +
        "Page reading: say 'read page', 'read headings', or 'read links'. " +
        "Navigation: say 'scroll down', 'scroll up', 'go to top', 'go to bottom', or 'go back'. " +
        "Heading navigation: say 'next heading' or 'previous heading'. " +
        "Link navigation: say 'next link' or 'previous link'. " +
        "YouTube: say 'search for' something, 'play first video', 'play', 'pause', 'forward', or 'backward'. " +
        "Media: say 'volume up', 'volume down', 'mute', or 'fullscreen'. " +
        "Utilities: say 'what time is it', 'what is the date', 'where am I', or 'zoom in', 'zoom out'. " +
        "Websites: say 'open' followed by a website name. " +
        "Say 'stop' to interrupt me, or 'go to sleep' when done."
      );
      return;
    }

    // --- Page Reading Commands ---

    // Read page content
    if (cmd.includes("read page") || cmd.includes("read this page") || cmd.includes("read content") || cmd.includes("read everything")) {
      readPageContent();
      return;
    }

    // Read headings
    if (cmd.includes("read headings") || cmd.includes("list headings") || cmd.includes("headings")) {
      readHeadings();
      return;
    }

    // Read links
    if (cmd.includes("read links") || cmd.includes("list links") || cmd.includes("show links")) {
      readLinks();
      return;
    }

    // Read selected text
    if (cmd.includes("read selection") || cmd.includes("read selected")) {
      const selection = window.getSelection().toString().trim();
      if (selection) {
        speak(selection);
      } else {
        speak("No text is selected.");
      }
      return;
    }

    // --- Heading & Link Navigation ---

    if (cmd.includes("next heading")) {
      navigateHeadings(1);
      return;
    }
    if (cmd.includes("previous heading") || cmd.includes("last heading")) {
      navigateHeadings(-1);
      return;
    }
    if (cmd.includes("next link")) {
      navigateLinks(1);
      return;
    }
    if (cmd.includes("previous link") || cmd.includes("last link")) {
      navigateLinks(-1);
      return;
    }

    // --- Click by text ---
    if (cmd.startsWith("click ")) {
      const target = cmd.replace("click ", "").trim();
      clickElementByText(target);
      return;
    }

    // --- YouTube Search ---
    if (cmd.includes("search for") || cmd.includes("search youtube")) {
      const query = cmd.replace(/search for|search youtube|search/g, "").trim();
      if (query) {
        performYouTubeSearch(query);
      } else {
        speak("What should I search for?");
      }
      return;
    }

    // Play first video
    if (cmd.includes("play first video") || cmd.includes("play first")) {
      playFirstVideo();
      return;
    }

    // Video views
    if (cmd.includes("views") && cmd.includes("first")) {
      getFirstVideoViews();
      return;
    }

    // --- Scroll & Navigation ---

    if (cmd.includes("scroll down") || cmd.includes("go down") || cmd.includes("page down")) {
      window.scrollBy({ top: window.innerHeight * 0.75, behavior: "smooth" });
      speak("Scrolled down.");
      return;
    }
    if (cmd.includes("scroll up") || cmd.includes("go up") || cmd.includes("page up")) {
      window.scrollBy({ top: -window.innerHeight * 0.75, behavior: "smooth" });
      speak("Scrolled up.");
      return;
    }
    if (cmd.includes("go to top") || cmd.includes("top of page") || cmd.includes("scroll to top")) {
      window.scrollTo({ top: 0, behavior: "smooth" });
      speak("At the top of the page.");
      return;
    }
    if (cmd.includes("go to bottom") || cmd.includes("bottom of page") || cmd.includes("scroll to bottom")) {
      window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" });
      speak("At the bottom of the page.");
      return;
    }
    if (cmd.includes("go back") || cmd === "back") {
      speak("Going back.");
      setTimeout(() => window.history.back(), 500);
      return;
    }
    if (cmd.includes("go forward") || cmd === "forward page") {
      speak("Going forward.");
      setTimeout(() => window.history.forward(), 500);
      return;
    }
    if (cmd.includes("refresh") || cmd.includes("reload")) {
      speak("Refreshing the page.");
      setTimeout(() => window.location.reload(), 500);
      return;
    }

    // --- Video Controls ---

    if ((cmd.includes("play") || cmd.includes("resume")) && !cmd.includes("pause") && !cmd.includes("first") && !cmd.includes("search")) {
      playVideo();
      return;
    }
    if (cmd.includes("pause") || cmd.includes("stop video")) {
      pauseVideo();
      return;
    }
    if (cmd.includes("forward") || cmd.includes("skip")) {
      const seconds = extractNumber(cmd, 10);
      skipVideo(seconds);
      return;
    }
    if (cmd.includes("backward") || cmd.includes("rewind")) {
      const seconds = extractNumber(cmd, 10);
      skipVideo(-seconds);
      return;
    }

    // Volume controls
    if (cmd.includes("volume up") || cmd.includes("louder")) {
      adjustVolume(0.2);
      return;
    }
    if (cmd.includes("volume down") || cmd.includes("quieter") || cmd.includes("softer")) {
      adjustVolume(-0.2);
      return;
    }
    if (cmd.includes("mute") || cmd.includes("unmute")) {
      toggleMute();
      return;
    }
    if (cmd.includes("fullscreen") || cmd.includes("full screen")) {
      toggleFullscreen();
      return;
    }

    // --- Page Summary ---
    if (cmd.includes("summarize") || cmd.includes("summary") || cmd.includes("describe page")) {
      summarizePage();
      return;
    }

    // --- Where Am I ---
    if (cmd.includes("where am i") || cmd.includes("what page") || cmd.includes("current page") || cmd.includes("what website")) {
      const title = document.title || "Untitled page";
      const host = window.location.hostname || "unknown site";
      speak(`You are on ${title}, on ${host}.`);
      return;
    }

    // --- Time & Date ---
    if (cmd.includes("time") || cmd.includes("what time")) {
      speak("The time is " + new Date().toLocaleTimeString());
      return;
    }
    if (cmd.includes("date") || cmd.includes("what date") || cmd.includes("today")) {
      speak("Today is " + new Date().toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" }));
      return;
    }

    // --- Zoom ---
    if (cmd.includes("zoom in") || cmd.includes("make bigger") || cmd.includes("increase size")) {
      document.body.style.zoom = (parseFloat(document.body.style.zoom || 1) + 0.1).toString();
      speak("Zoomed in.");
      return;
    }
    if (cmd.includes("zoom out") || cmd.includes("make smaller") || cmd.includes("decrease size")) {
      document.body.style.zoom = (parseFloat(document.body.style.zoom || 1) - 0.1).toString();
      speak("Zoomed out.");
      return;
    }
    if (cmd.includes("reset zoom") || cmd.includes("normal size")) {
      document.body.style.zoom = "1";
      speak("Zoom reset to normal.");
      return;
    }

    // --- Open Website ---
    if (cmd.includes("open ")) {
      const site = cmd.replace("open ", "").trim();
      openWebsite(site);
      return;
    }

    // Greeting
    if (cmd === "hello" || cmd === "hi" || cmd.includes("how are you")) {
      speak("Hello! I'm here to help. Say 'help' for a list of commands.");
      return;
    }

    // Fallback
    speak("I didn't catch that. Say 'help' for available commands.");
  }

  // --- Utility: Extract number from command (e.g., "skip 30 seconds" → 30) ---
  function extractNumber(text, defaultVal) {
    const match = text.match(/(\d+)/);
    return match ? parseInt(match[1], 10) : defaultVal;
  }

  // --- Page Reading ---
  function readPageContent() {
    // Try to find main content area
    const mainContent =
      document.querySelector("main") ||
      document.querySelector("article") ||
      document.querySelector('[role="main"]') ||
      document.querySelector("#content") ||
      document.querySelector(".content") ||
      document.body;

    // Get visible text, skipping nav/header/footer/aside
    const skipTags = new Set(["NAV", "HEADER", "FOOTER", "ASIDE", "SCRIPT", "STYLE", "NOSCRIPT", "SVG"]);
    let text = "";
    const walker = document.createTreeWalker(mainContent, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        if (skipTags.has(node.parentElement?.tagName)) return NodeFilter.FILTER_REJECT;
        if (node.parentElement?.closest("nav, header, footer, aside, script, style")) return NodeFilter.FILTER_REJECT;
        const trimmed = node.textContent.trim();
        if (!trimmed) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      },
    });

    let node;
    while ((node = walker.nextNode())) {
      text += node.textContent.trim() + " ";
      if (text.length > 3000) break; // Cap length for speech
    }

    text = text.replace(/\s+/g, " ").trim();
    if (text.length > 0) {
      speak("Reading page content. " + text + (text.length >= 3000 ? "... Content truncated. Say 'scroll down' and 'read page' for more." : ""));
    } else {
      speak("I couldn't find readable content on this page.");
    }
  }

  function readHeadings() {
    const headings = document.querySelectorAll("h1, h2, h3, h4, h5, h6");
    if (headings.length === 0) {
      speak("No headings found on this page.");
      return;
    }
    let text = `Found ${headings.length} headings. `;
    const limit = Math.min(headings.length, 15);
    for (let i = 0; i < limit; i++) {
      const h = headings[i];
      const level = h.tagName.replace("H", "");
      text += `Level ${level}: ${h.textContent.trim()}. `;
    }
    if (headings.length > limit) {
      text += `And ${headings.length - limit} more. Say 'next heading' to navigate.`;
    }
    speak(text);
  }

  function readLinks() {
    const links = Array.from(document.querySelectorAll("a[href]")).filter((a) => {
      const rect = a.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0 && a.textContent.trim();
    });
    if (links.length === 0) {
      speak("No links found on this page.");
      return;
    }
    let text = `Found ${links.length} links. `;
    const limit = Math.min(links.length, 10);
    for (let i = 0; i < limit; i++) {
      text += `${i + 1}: ${links[i].textContent.trim()}. `;
    }
    if (links.length > limit) {
      text += `And ${links.length - limit} more. Say 'next link' to navigate through them.`;
    }
    speak(text);
  }

  // --- Heading Navigation ---
  function navigateHeadings(direction) {
    const headings = Array.from(document.querySelectorAll("h1, h2, h3, h4, h5, h6"));
    if (headings.length === 0) {
      speak("No headings on this page.");
      return;
    }
    headingIndex += direction;
    if (headingIndex < 0) headingIndex = headings.length - 1;
    if (headingIndex >= headings.length) headingIndex = 0;

    const h = headings[headingIndex];
    h.scrollIntoView({ behavior: "smooth", block: "center" });
    h.focus();
    speak(`Heading ${headingIndex + 1} of ${headings.length}. Level ${h.tagName.replace("H", "")}: ${h.textContent.trim()}`);
  }

  // --- Link Navigation ---
  function navigateLinks(direction) {
    const links = Array.from(document.querySelectorAll("a[href]")).filter((a) => {
      const rect = a.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0 && a.textContent.trim();
    });
    if (links.length === 0) {
      speak("No links on this page.");
      return;
    }
    linkIndex += direction;
    if (linkIndex < 0) linkIndex = links.length - 1;
    if (linkIndex >= links.length) linkIndex = 0;

    const link = links[linkIndex];
    link.scrollIntoView({ behavior: "smooth", block: "center" });
    link.focus();
    // Highlight
    link.style.outline = "3px solid #00ff88";
    link.style.outlineOffset = "2px";
    setTimeout(() => {
      link.style.outline = "";
      link.style.outlineOffset = "";
    }, 3000);
    speak(`Link ${linkIndex + 1} of ${links.length}: ${link.textContent.trim()}. Say 'click' to open it.`);
  }

  // --- Click Element by Text ---
  function clickElementByText(target) {
    // If user says "click" with no target and we have a focused link
    if (!target || target === "it" || target === "this") {
      const links = Array.from(document.querySelectorAll("a[href]")).filter(
        (a) => a.getBoundingClientRect().width > 0 && a.textContent.trim()
      );
      if (linkIndex >= 0 && linkIndex < links.length) {
        links[linkIndex].click();
        speak("Clicked " + links[linkIndex].textContent.trim());
        return;
      }
      speak("No element to click. Say 'next link' first, or say 'click' followed by the text you want to click.");
      return;
    }

    // Search for clickable elements matching the text
    const clickables = document.querySelectorAll("a, button, [role='button'], [role='link'], input[type='submit'], input[type='button']");
    let bestMatch = null;
    let bestScore = 0;
    for (const el of clickables) {
      const elText = (el.textContent || el.value || el.getAttribute("aria-label") || "").trim().toLowerCase();
      if (elText === target) {
        bestMatch = el;
        break;
      }
      if (elText.includes(target) && elText.length < (bestMatch ? bestMatch.textContent.length : Infinity)) {
        bestMatch = el;
        bestScore = 1;
      }
    }

    if (bestMatch) {
      bestMatch.scrollIntoView({ behavior: "smooth", block: "center" });
      bestMatch.click();
      speak("Clicked " + (bestMatch.textContent || bestMatch.value || target).trim().substring(0, 50));
    } else {
      speak(`Couldn't find a clickable element matching "${target}".`);
    }
  }

  // --- YouTube Functions ---
  function performYouTubeSearch(query) {
    const searchInput =
      document.querySelector("input#search") ||
      document.querySelector("input[name='search_query']") ||
      document.querySelector("input[placeholder*='Search']");

    if (searchInput) {
      searchInput.value = query;
      searchInput.focus();
      const form = searchInput.closest("form");
      if (form) {
        form.submit();
      } else {
        searchInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", keyCode: 13, bubbles: true }));
      }
      speak("Searching for " + query);
    } else {
      window.location.href = "https://www.youtube.com/results?search_query=" + encodeURIComponent(query);
      speak("Searching YouTube for " + query);
    }
  }

  function playFirstVideo() {
    const firstVideo =
      document.querySelector("ytd-video-renderer a#video-title") ||
      document.querySelector("ytd-rich-item-renderer a#video-title") ||
      document.querySelector("a#video-title");
    if (firstVideo) {
      firstVideo.click();
      speak("Playing first video.");
    } else {
      speak("No videos found on this page.");
    }
  }

  function getFirstVideoViews() {
    const firstVideo = document.querySelector("ytd-video-renderer") || document.querySelector("ytd-rich-item-renderer");
    if (firstVideo) {
      const views = firstVideo.querySelector("#metadata-line span") || firstVideo.querySelector(".style-scope.ytd-video-meta-block");
      if (views) {
        speak("The first video has " + views.textContent.trim());
      } else {
        speak("View count not available.");
      }
    } else {
      speak("No videos found.");
    }
  }

  // --- Media Controls ---
  function playVideo() {
    const video = document.querySelector("video");
    if (video) {
      video.play();
      speak("Playing.");
    } else {
      const btn = document.querySelector("button[aria-label*='Play']") || document.querySelector(".ytp-play-button");
      if (btn) { btn.click(); speak("Playing."); }
      else speak("No video found on this page.");
    }
  }

  function pauseVideo() {
    const video = document.querySelector("video");
    if (video) {
      video.pause();
      speak("Paused.");
    } else {
      const btn = document.querySelector("button[aria-label*='Pause']") || document.querySelector(".ytp-play-button");
      if (btn) { btn.click(); speak("Paused."); }
      else speak("No video found.");
    }
  }

  function skipVideo(seconds) {
    const video = document.querySelector("video");
    if (video) {
      video.currentTime += seconds;
      speak((seconds > 0 ? "Forward " : "Backward ") + Math.abs(seconds) + " seconds.");
    } else {
      speak("No video found.");
    }
  }

  function adjustVolume(delta) {
    const video = document.querySelector("video") || document.querySelector("audio");
    if (video) {
      video.volume = Math.min(1, Math.max(0, video.volume + delta));
      speak("Volume " + Math.round(video.volume * 100) + " percent.");
    } else {
      speak("No media found to adjust volume.");
    }
  }

  function toggleMute() {
    const video = document.querySelector("video") || document.querySelector("audio");
    if (video) {
      video.muted = !video.muted;
      speak(video.muted ? "Muted." : "Unmuted.");
    } else {
      speak("No media found.");
    }
  }

  function toggleFullscreen() {
    const video = document.querySelector("video");
    if (video) {
      if (document.fullscreenElement) {
        document.exitFullscreen();
        speak("Exited fullscreen.");
      } else {
        video.requestFullscreen().then(() => speak("Fullscreen.")).catch(() => speak("Fullscreen not allowed on this page."));
      }
    } else {
      speak("No video found.");
    }
  }

  // --- Page Summary ---
  function summarizePage() {
    const url = window.location.href;
    if (url.includes("youtube.com")) {
      if (url.includes("watch")) {
        const title = document.querySelector("h1.ytd-video-primary-info-renderer, h1.title, #title h1, h1.style-scope.ytd-watch-metadata");
        const channel = document.querySelector("#channel-name a, .ytd-channel-name a, #owner-name a");
        let msg = "You are watching a YouTube video. ";
        if (title) msg += "Title: " + title.textContent.trim() + ". ";
        if (channel) msg += "By " + channel.textContent.trim() + ". ";
        const video = document.querySelector("video");
        if (video && video.duration) {
          const mins = Math.floor(video.duration / 60);
          const secs = Math.floor(video.duration % 60);
          msg += `Duration: ${mins} minutes and ${secs} seconds. `;
          const current = Math.floor(video.currentTime / 60);
          const curSecs = Math.floor(video.currentTime % 60);
          msg += `You are at ${current} minutes and ${curSecs} seconds.`;
        }
        speak(msg);
      } else {
        const videos = document.querySelectorAll("ytd-video-renderer, ytd-rich-item-renderer");
        let summary = `YouTube page with ${videos.length} videos. `;
        const limit = Math.min(videos.length, 5);
        for (let i = 0; i < limit; i++) {
          const title = videos[i].querySelector("#video-title")?.textContent?.trim();
          if (title) summary += `${i + 1}: ${title}. `;
        }
        speak(summary);
      }
    } else {
      const title = document.title || "Untitled";
      const headings = document.querySelectorAll("h1, h2");
      let summary = `Page: ${title}. `;
      if (headings.length > 0) {
        summary += `Main sections: `;
        const limit = Math.min(headings.length, 5);
        for (let i = 0; i < limit; i++) {
          summary += headings[i].textContent.trim() + ". ";
        }
      }
      const links = document.querySelectorAll("a[href]");
      summary += `There are ${links.length} links on this page.`;
      speak(summary);
    }
  }

  // --- Open Website ---
  function openWebsite(site) {
    const siteMap = {
      youtube: "https://www.youtube.com",
      google: "https://www.google.com",
      wikipedia: "https://www.wikipedia.org",
      gmail: "https://mail.google.com",
      twitter: "https://www.twitter.com",
      reddit: "https://www.reddit.com",
      facebook: "https://www.facebook.com",
      amazon: "https://www.amazon.com",
      netflix: "https://www.netflix.com",
      github: "https://www.github.com",
    };

    const key = Object.keys(siteMap).find((k) => site.includes(k));
    if (key) {
      speak("Opening " + key + ".");
      setTimeout(() => { window.location.href = siteMap[key]; }, 500);
    } else if (site.includes(".")) {
      // Looks like a domain
      const url = site.startsWith("http") ? site : "https://" + site;
      speak("Opening " + site + ".");
      setTimeout(() => { window.location.href = url; }, 500);
    } else {
      // Try as a Google search
      speak("Searching Google for " + site + ".");
      setTimeout(() => {
        window.location.href = "https://www.google.com/search?q=" + encodeURIComponent(site);
      }, 500);
    }
  }

  // --- Keyboard Shortcut: Ctrl+Shift+S to toggle wake ---
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey && e.shiftKey && e.key === "S") {
      e.preventDefault();
      if (isAwake) {
        isAwake = false;
        clearTimeout(awakeTimeout);
        stopSpeaking();
        speak("Going to sleep.", () => listenForWakeWord());
      } else {
        if (wakeRecognizer) {
          try { wakeRecognizer.stop(); } catch (_) {}
        }
        activateAssistant();
      }
    }
  });

  // --- Auto-Restart & Visibility ---
  setInterval(() => {
    if (!isListening && !isAwake && !isSpeaking && consecutiveFailures < 5) {
      listenForWakeWord();
    }
  }, 20000);

  document.addEventListener("visibilitychange", () => {
    if (document.hidden) {
      if (wakeRecognizer) try { wakeRecognizer.stop(); } catch (_) {}
      if (commandRecognizer) try { commandRecognizer.stop(); } catch (_) {}
      stopSpeaking();
    } else {
      setTimeout(() => {
        if (!isListening && !isAwake) listenForWakeWord();
      }, 1000);
    }
  });

  // --- Initialize ---
  function initialize() {
    statusBar = createStatusBar();
    updateStatus("Requesting microphone access...", "🔄");

    navigator.mediaDevices
      .getUserMedia({ audio: true })
      .then(() => {
        if (!hasAnnouncedReady) {
          hasAnnouncedReady = true;
          updateStatus("Voice Assistant Ready", "✅");
          speak("Luma voice assistant ready. Say 'Hey Luma' to activate, or press Control Shift S.", () => {
            listenForWakeWord();
          });
        }
      })
      .catch(() => {
        updateStatus("Microphone access denied", "❌");
        speak("Microphone access is required. Please allow microphone access and refresh the page.");
      });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize);
  } else {
    initialize();
  }
})();
