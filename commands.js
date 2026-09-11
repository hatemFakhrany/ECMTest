// ==== CONFIGURE THIS ====
const API_ENDPOINT = "https://ileana-nonsculptural-ludie.ngrok-free.dev/api/Email"; // NOTE: must be HTTPS in production, see notes
// =========================

Office.onReady(() => {
  // Required so Outlook knows this file's functions are ready to be called
});

function saveToEcm(event) {
  const item = Office.context.mailbox.item;

  if (!item) {
    showNotification("error", "No email selected.");
    event.completed();
    return;
  }

  collectEmailData(item, (data) => {
    fetch(API_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data)
    })
      .then((response) => {
        if (!response.ok) {
          throw new Error("Server responded with " + response.status);
        }
        return response.text();
      })
      .then(() => {
        showNotification("success", "Saved to ECM successfully.");
        event.completed();
      })
      .catch((err) => {
        showNotification("error", "Save failed: " + err.message);
        event.completed();
      });
  });
}

// Collects subject, from, to, body, and attachment contents (base64)
function collectEmailData(item, callback) {
  const subject = item.subject || "";
  const from = item.from ? item.from.emailAddress : "";
  const to = (item.to || []).map(r => r.emailAddress).join(";");

  item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
    const body = bodyResult.status === Office.AsyncResultStatus.Succeeded
      ? bodyResult.value
      : "";

    const attachmentDefs = (item.attachments || []).filter(a => !a.isInline);

    if (attachmentDefs.length === 0) {
      callback({ subject, from, to, body, attachments: [] });
      return;
    }

    let remaining = attachmentDefs.length;
    const attachments = [];

    attachmentDefs.forEach((att) => {
      item.getAttachmentContentAsync(att.id, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded &&
            result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          attachments.push({
            name: att.name,
            contentType: att.contentType || "application/octet-stream",
            sizeBytes: att.size,
            contentBase64: result.value.content
          });
        }
        remaining--;
        if (remaining === 0) {
          callback({ subject, from, to, body, attachments });
        }
      });
    });
  });
}

// Shows a small banner notification inside the open email in Outlook
function showNotification(type, message) {
  const item = Office.context.mailbox.item;
  if (!item || !item.notificationMessages) return;

  item.notificationMessages.replaceAsync("ecmSaveStatus", {
    type: type === "success"
      ? Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage
      : Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message: message,
    icon: "icon-16",
    persistent: false
  });
}

function createNewDocument(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
    showNotification("error", "No email selected.");
    event.completed();
    return;
  }

  // TODO: replace this with your actual "create document" API call
  collectEmailData(item, (data) => {
    fetch("https://ileana-nonsculptural-ludie.ngrok-free.dev/api/Email/CreateDocument", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data)
    })
      .then((response) => {
        if (!response.ok) throw new Error("Server responded with " + response.status);
        return response.text();
      })
      .then(() => {
        showNotification("success", "Document created in ECM.");
        event.completed();
      })
      .catch((err) => {
        showNotification("error", "Create document failed: " + err.message);
        event.completed();
      });
  });
}

function convertToCorrespondence(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
    showNotification("error", "No email selected.");
    event.completed();
    return;
  }

  // TODO: replace this with your actual "convert to correspondence" API call
  collectEmailData(item, (data) => {
    fetch("https://ileana-nonsculptural-ludie.ngrok-free.dev/api/Email/ConvertToCorrespondence", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data)
    })
      .then((response) => {
        if (!response.ok) throw new Error("Server responded with " + response.status);
        return response.text();
      })
      .then(() => {
        showNotification("success", "Converted to correspondence.");
        event.completed();
      })
      .catch((err) => {
        showNotification("error", "Convert failed: " + err.message);
        event.completed();
      });
  });
}

// Register the functions so the manifest's <FunctionName> tags can find them
Office.actions.associate("createNewDocument", createNewDocument);
Office.actions.associate("convertToCorrespondence", convertToCorrespondence);