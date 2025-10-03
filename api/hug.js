export default async function handler(req, res) {
  const prompt = req.body.prompt;

  try {
    const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${process.env.HERMES_API_KEY}`,//`Bearer ${process.env.MISTRAL_API_KEY}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "nousresearch/deephermes-3-llama-3-8b-preview:free",//"mistralai/mistral-7b-instruct:free",
        messages: [
          {
            role: "user",
            content: prompt
          }
        ]
      })
    });

    const data = await response.json();
    const message = data.choices?.[0]?.message?.content || "No message returned";

    console.log("OpenRouter response:", message);

    // ✅ Send a proper JSON response
    res.status(200).json({ message });
  } catch (error) {
    console.error("Fetch failed:", error);
    res.status(500).json({ error: "Hermes AI request failed", details: error.message });
  }
}

