(function () {
  let chatbotReady = false;
  let chatbotInitPromise = null;

  const CHATBOT_MOUNT_ID = "sia-chatbot-root";
  const DEFAULT_STYLESHEET_URL = "https://cdn.jsdelivr.net/npm/@n8n/chat/dist/style.css";
  const DEFAULT_SCRIPT_URL = "https://cdn.jsdelivr.net/npm/@n8n/chat/dist/chat.bundle.es.js";
  const DEFAULT_ALLOWED_FILES = "image/*,application/pdf";

  const DEFAULT_I18N = {
    en: {
      title: "Asistente SIA",
      subtitle: "Soporte de SIA",
      getStarted: "Nueva conversacion",
      inputPlaceholder: "Escribe tu consulta...",
      footer: "SIA"
    }
  };

  const DEFAULT_INITIAL_MESSAGES = [
    "Hola, soy el Asistente de SIA.",
    "Estoy listo para ayudarte con tus dudas del sistema."
  ];

  function getConfig() {
    const config = window.SIA_CHATBOT_CONFIG || {};
    return {
      enabled: config.enabled ?? true,
      webhookUrl: String(config.webhookUrl || "").trim(),
      metadata: config.metadata || {},
      defaultLanguage: config.defaultLanguage || "en",
      showWelcomeScreen: config.showWelcomeScreen ?? true,
      loadPreviousSession: config.loadPreviousSession ?? true,
      allowFileUploads: config.allowFileUploads ?? true,
      allowedFilesMimeTypes: config.allowedFilesMimeTypes || DEFAULT_ALLOWED_FILES,
      i18n: config.i18n || DEFAULT_I18N,
      initialMessages:
        Array.isArray(config.initialMessages) && config.initialMessages.length
          ? config.initialMessages
          : DEFAULT_INITIAL_MESSAGES,
      stylesheetUrl: config.stylesheetUrl || DEFAULT_STYLESHEET_URL,
      scriptUrl: config.scriptUrl || DEFAULT_SCRIPT_URL
    };
  }

  function ensureStylesheet(stylesheetUrl) {
    const existing = document.querySelector('link[data-chatbot-style="n8n"]');
    if (existing) {
      return;
    }

    const stylesheet = document.createElement("link");
    stylesheet.rel = "stylesheet";
    stylesheet.href = stylesheetUrl;
    stylesheet.dataset.chatbotStyle = "n8n";
    document.head.appendChild(stylesheet);
  }

  function ensureMountNode() {
    let mountNode = document.getElementById(CHATBOT_MOUNT_ID);
    if (mountNode) {
      return mountNode;
    }

    mountNode = document.createElement("div");
    mountNode.id = CHATBOT_MOUNT_ID;
    mountNode.dataset.chatbotMount = "true";
    mountNode.style.display = "none";
    document.body.appendChild(mountNode);
    return mountNode;
  }

  function setChatbotSessionActive(isActive) {
    const active = Boolean(isActive);
    const mountNode = ensureMountNode();

    mountNode.style.display = active ? "" : "none";
    document.body.dataset.siaChatbotSession = active ? "on" : "off";

    const chatRoot = document.querySelector("n8n-chat");
    if (chatRoot) {
      chatRoot.style.display = active ? "" : "none";
    }
  }

  async function initChatbot() {
    const config = getConfig();
    if (!config.enabled) {
      return;
    }

    if (!config.webhookUrl) {
      console.warn("Chatbot deshabilitado: falta webhookUrl en SIA_CHATBOT_CONFIG.");
      return;
    }

    setChatbotSessionActive(true);

    if (chatbotReady) {
      return;
    }

    if (chatbotInitPromise) {
      return chatbotInitPromise;
    }

    ensureStylesheet(config.stylesheetUrl);
    const mountNode = ensureMountNode();

    chatbotInitPromise = import(config.scriptUrl)
      .then((moduleRef) => {
        const createChat = moduleRef && moduleRef.createChat;
        if (typeof createChat !== "function") {
          throw new Error("No se pudo cargar createChat del proveedor del chatbot.");
        }

        createChat({
          webhookUrl: config.webhookUrl,
          metadata: config.metadata,
          defaultLanguage: config.defaultLanguage,
          showWelcomeScreen: config.showWelcomeScreen,
          loadPreviousSession: config.loadPreviousSession,
          allowFileUploads: config.allowFileUploads,
          allowedFilesMimeTypes: config.allowedFilesMimeTypes,
          i18n: config.i18n,
          initialMessages: config.initialMessages,
          target: `#${mountNode.id}`
        });

        chatbotReady = true;
        setChatbotSessionActive(true);
      })
      .catch((error) => {
        console.warn("No se pudo inicializar el chatbot", error);
      })
      .finally(() => {
        chatbotInitPromise = null;
      });

    return chatbotInitPromise;
  }

  window.addEventListener("sia:user-authenticated", () => {
    void initChatbot();
  });

  window.addEventListener("sia:user-logged-out", () => {
    setChatbotSessionActive(false);
  });

  if (document.readyState === "loading") {
    document.addEventListener(
      "DOMContentLoaded",
      () => {
        setChatbotSessionActive(false);
      },
      { once: true }
    );
  } else {
    setChatbotSessionActive(false);
  }

  window.initSiaChatbot = initChatbot;
  window.setSiaChatbotSessionActive = setChatbotSessionActive;
})();
