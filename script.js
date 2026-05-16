const CONFIG = {
  title: "Ontslag formulier",
  botName: "Google Forms Bot",
  avatarImage: "",
  shortDescription: "",
  colour: "#FCD32F",
  mention: ""
};

function getWebhooks() {
  const raw = PropertiesService.getScriptProperties().getProperty("WEBHOOK_URL");
  return raw ? raw.split(",").map(u => u.trim()).filter(Boolean) : [];
}

function setWebhook() {
  const url = "YOUR_WEBHOOK_URL_NEEDS_TO_COME_HERE";

  if (url.startsWith("VERVANG")) {
    Logger.log("⚠️ Vul eerst je webhook URL in bij de variabele 'url'.");
    return;
  }

  PropertiesService.getScriptProperties().setProperty("WEBHOOK_URL", url);
  Logger.log("✅ Webhook opgeslagen.");
}


function embedText(e) {
  try {
    if (!e || !e.response) {
      Logger.log("⚠️ Geen trigger-event ontvangen. Gebruik testWebhook() om handmatig te testen.");
      return;
    }

    const responses = e.response.getItemResponses();

    const items = [];
    for (const itemResponse of responses) {
      const question = itemResponse.getItem().getTitle();
      const answer = itemResponse.getResponse();

      if (answer !== null && answer !== undefined && answer.toString().trim() !== "") {
        items.push({
          name: question,
          value: Array.isArray(answer) ? answer.join(", ") : answer.toString()
        });
      }
    }

    if (items.length === 0) {
      Logger.log("Geen ingevulde antwoorden.");
      return;
    }

    const description = items
      .map(item => `**${item.name}**\n${item.value}`)
      .join("\n\n")
      .substring(0, 4096);

    const embed = {
      title: CONFIG.title || "Formulierinzending",
      description: CONFIG.shortDescription
        ? `${CONFIG.shortDescription}\n\n${description}`
        : description,
      timestamp: new Date().toISOString()
    };

    if (CONFIG.colour) {
      const parsed = parseInt(CONFIG.colour.replace("#", ""), 16);
      if (!isNaN(parsed)) embed.color = parsed;
    }

    const payload = {
      username: CONFIG.botName || "Google Forms Bot",
      content: CONFIG.mention || "",
      embeds: [embed]
    };

    if (CONFIG.avatarImage) {
      payload.avatar_url = CONFIG.avatarImage;
    }

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const WEBHOOKS = getWebhooks();
    if (WEBHOOKS.length === 0) {
      Logger.log("⚠️ Geen webhook ingesteld. Voer setWebhook() uit.");
      return;
    }

    for (const webhook of WEBHOOKS) {
      const fetchResponse = UrlFetchApp.fetch(webhook, options);
      const status = fetchResponse.getResponseCode();

      if (status === 204) {
        Logger.log("✅ Webhook succesvol verzonden.");
      } else {
        Logger.log(`⚠️ Webhook status ${status}: ${fetchResponse.getContentText()}`);
      }
    }

  } catch (error) {
    Logger.log("❌ FOUT:");
    Logger.log(error.toString());
  }
}

function testWebhook() {
  const WEBHOOKS = getWebhooks();

  if (WEBHOOKS.length === 0) {
    Logger.log("⚠️ Geen webhook gevonden. Voer eerst setWebhook() uit.");
    return;
  }

  const payload = JSON.stringify({
    username: CONFIG.botName || "Google Forms Bot",
    avatar_url: CONFIG.avatarImage || undefined,
    content: "✅ Test bericht vanuit Google Apps Script",
    embeds: [
      {
        title: `📋 TEST — ${CONFIG.title}`,
        description: "**Naam**\nJan Jansen\n\n**E-mail**\njan@voorbeeld.nl\n\n**Bericht**\nDit is een testbericht!",
        color: parseInt((CONFIG.colour || "#5865F2").replace("#", ""), 16),
        timestamp: new Date().toISOString()
      }
    ]
  });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload,
    muteHttpExceptions: true
  };

  for (const webhook of WEBHOOKS) {
    const response = UrlFetchApp.fetch(webhook, options);
    const status = response.getResponseCode();
    Logger.log(`Status: ${status}`);
    if (status !== 204) Logger.log(response.getContentText());
    else Logger.log("✅ Succesvol verzonden (204 No Content).");
  }
}

function createFormTrigger() {
  const form = FormApp.getActiveForm();
  if (!form) {
    Logger.log("⚠️ Geen actieve form gevonden.");
    return;
  }

  const existing = ScriptApp.getUserTriggers(form).filter(
    t => t.getHandlerFunction() === "embedText"
  );
  if (existing.length > 0) {
    Logger.log("ℹ️ Trigger bestaat al.");
    return;
  }

  ScriptApp.newTrigger("embedText")
    .forForm(form)
    .onFormSubmit()
    .create();

  Logger.log("✅ onFormSubmit trigger aangemaakt.");
}

function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log("🗑️ Alle triggers verwijderd.");
}
