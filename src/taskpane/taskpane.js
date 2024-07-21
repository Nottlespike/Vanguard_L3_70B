/* global Office, axios */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("analyze-email").onclick = analyzeEmail;
    document.getElementById("theme-toggle").onclick = toggleTheme;
  }
});

async function analyzeEmail() {
  try {
    const item = Office.context.mailbox.item;
    if (!item) throw new Error("No email item found");

    const subject = item.subject || "";
    const sender = item.sender ? item.sender.emailAddress : "";
    const body = await getBodyAsPlainText(item);

    const result = await sendToOpenAICompatibleEndpoint(subject, sender, body);

    const resultElement = document.getElementById("result");
    if (resultElement) {
      let resultHtml = result.isMalicious
        ? `<p style="color: red;">Warning: This email may be malicious!</p>`
        : `<p style="color: green;">This email appears to be safe.</p>`;

      resultHtml += `<p><strong>Explanation:</strong> ${result.explanation}</p>`;
      resultElement.innerHTML = resultHtml;
    }
  } catch (error) {
    console.error("Error:", error);
    const resultElement = document.getElementById("result");
    if (resultElement) {
      resultElement.innerHTML = "An error occurred while analyzing the email.";
    }
  }
}

function getBodyAsPlainText(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, {}, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to get email body"));
      }
    });
  });
}

async function sendToOpenAICompatibleEndpoint(subject, sender, body) {
  const endpoint = "https://aphrodite.ngrok.io/v1/chat/completions";
  const apiKey = "295a1091a126606dfe47ca8b85539ff2";

  try {
    const response = await axios.post(
      endpoint,
      {
        model: "L3Vanguard",
        messages: [
          {
            role: "system",
            content:
              "You are an AI assistant that analyzes emails for potential security threats.  Respond with a JSON object containing a boolean 'isMalicious' field and a string 'explanation' field that provides a brief explanation of your analysis.",
          },
          {
            role: "user",
            content: `Please analyze this email for potential security threats:\n\nSubject: ${subject}\nFrom: ${sender}\n\nBody:\n${body}`,
          },
        ],
        temperature: 1.25, // Added temperature parameter
      },
      {
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
        },
      },
    );

    const aiResponse = JSON.parse(response.data.choices[0].message.content);
    return {
      isMalicious: aiResponse.isMalicious,
      explanation: aiResponse.explanation,
    };
  } catch (error) {
    console.error("Error calling OpenAI compatible endpoint:", error);
    throw error;
  }
}

function toggleTheme() {
  document.body.classList.toggle("dark-mode");
}

function sanitizeString(str) {
  const map = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#x27;",
    "/": "&#x2F;",
    "`": "&grave;",
    "=": "&#x3D;",
  };
  const reg = /[&<>"'`=\/]/gi;
  return str.replace(reg, (match) => map[match]);
}
