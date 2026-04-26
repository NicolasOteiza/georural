(function () {
  let chatbotReady = false;
  let chatbotInitPromise = null;
  let authenticatedUser = null;

  const CHATBOT_MOUNT_ID = "sia-chatbot-root";
  const CHATBOT_RUNTIME_STYLE_ID = "sia-chatbot-runtime-style";
  const DEFAULT_STYLESHEET_URL = "https://cdn.jsdelivr.net/npm/@n8n/chat/dist/style.css";
  const DEFAULT_SCRIPT_URL = "https://cdn.jsdelivr.net/npm/@n8n/chat/dist/chat.bundle.es.js";
  const DEFAULT_ALLOWED_FILES = "image/*,application/pdf";
  const DEFAULT_WINDOW_WIDTH = "330px";
  const DEFAULT_WINDOW_HEIGHT = "500px";
  const DEFAULT_WINDOW_WIDTH_MOBILE = "calc(100vw - 20px)";
  const DEFAULT_WINDOW_HEIGHT_MOBILE = "68vh";

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

  function normalizeCssSize(value, fallback) {
    const normalized = String(value || "").trim();
    return normalized || fallback;
  }

  function resolveMetadata(rawMetadata) {
    const metadata =
      rawMetadata && typeof rawMetadata === "object" ? { ...rawMetadata } : {};
    const rawUserName = String(metadata.user_name || "").trim();
    const phpTemplateInValue = rawUserName.includes("<?php");

    if (!rawUserName || phpTemplateInValue) {
      const fallbackUserName = String(
        (authenticatedUser &&
          (authenticatedUser.nombre ||
            authenticatedUser.username ||
            authenticatedUser.user_name)) ||
          ""
      ).trim();

      if (fallbackUserName) {
        metadata.user_name = fallbackUserName;
      } else if (phpTemplateInValue) {
        delete metadata.user_name;
      }
    }

    return metadata;
  }

  function getConfig() {
    const config = window.SIA_CHATBOT_CONFIG || {};
    return {
      enabled: config.enabled ?? true,
      webhookUrl: String(config.webhookUrl || "").trim(),
      metadata: resolveMetadata(config.metadata),
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
      scriptUrl: config.scriptUrl || DEFAULT_SCRIPT_URL,
      windowWidth: normalizeCssSize(config.windowWidth, DEFAULT_WINDOW_WIDTH),
      windowHeight: normalizeCssSize(config.windowHeight, DEFAULT_WINDOW_HEIGHT),
      windowWidthMobile: normalizeCssSize(
        config.windowWidthMobile,
        DEFAULT_WINDOW_WIDTH_MOBILE
      ),
      windowHeightMobile: normalizeCssSize(
        config.windowHeightMobile,
        DEFAULT_WINDOW_HEIGHT_MOBILE
      )
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

  function ensureRuntimeStyles(config) {
    let styleTag = document.getElementById(CHATBOT_RUNTIME_STYLE_ID);
    if (!styleTag) {
      styleTag = document.createElement("style");
      styleTag.id = CHATBOT_RUNTIME_STYLE_ID;
      document.head.appendChild(styleTag);
    }

    styleTag.textContent = `
:root {
  --chat--window--width: ${config.windowWidth};
  --chat--window--height: ${config.windowHeight};
}

@media (max-width: 768px) {
  :root {
    --chat--window--width: ${config.windowWidthMobile};
    --chat--window--height: ${config.windowHeightMobile};
  }
}
`;
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

    ensureRuntimeStyles(config);
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

  window.addEventListener("sia:user-authenticated", (event) => {
    authenticatedUser =
      event && event.detail && event.detail.user ? event.detail.user : null;
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
