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

      // Sanitize the explanation to prevent XSS attacks
      const sanitizedExplanation = sanitizeString(result.explanation);
      resultHtml += `<p><strong>Explanation:</strong> ${sanitizedExplanation}</p>`;
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
  // TODO: Move these to environment variables or secure configuration
  const endpoint = process.env.API_ENDPOINT || "";
  const apiKey = process.env.API_KEY || "";

  if (!endpoint || !apiKey) {
    throw new Error("API endpoint and key must be configured");
  }

  try {
    const response = await axios.post(
      endpoint,
      {
        model: "L3Vanguard",
        messages: [
          {
            role: "system",
            content:
              "You are an expert cybersecurity AI assistant that analyzes emails for potential security threats. Mull over all information available to you before giving your verdict. Respond with a JSON object containing a boolean 'isMalicious' field and a string 'explanation' field that provides a brief explanation of your analysis.",
          },
          {
            role: "user",
            content: `Please analyze this email for potential security threats:\n\nSubject: ${subject}\nFrom: ${sender}\n\nBody:\n${body}`,
          },
        ],
        temperature: 0.4, // Added temperature parameter
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
