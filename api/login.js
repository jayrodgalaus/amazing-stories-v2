import crypto from 'crypto';

export default async function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { email, password } = req.body;

    // Basic validation
    if (
        typeof email !== 'string' ||
        typeof password !== 'string' ||
        !email.trim() ||
        !password.trim() ||
        !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
    ) {
        return res.status(400).json({ error: 'Invalid email or password format' });
    }

    // Sanitize input
    const sanitizedEmail = email.trim().toLowerCase();
    const sanitizedPassword = password.trim();

    // Hash password using SHA-256
    const hashedPassword = crypto
        .createHash('sha256')
        .update(sanitizedPassword)
        .digest('hex');

    // Prepare payload for Power Automate
    const payload = {
        email: sanitizedEmail,
        password: hashedPassword
    };

    try {
        const response = await fetch('https://default93f33571550f43cfb09fcd331338d0.86.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/58c47031d9fa4d11b5bd6b73a7bbb25f/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EZLMgkmiuAzlFhpiDBaPPi2ZC36BDNXGTvZ07lYXG1o', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        const result = await response.json();

        return res.status(response.status).json(result);
    } catch (error) {
        console.error('Error sending to Power Automate:', error);
        return res.status(500).json({ error: 'Internal server error' });
    }
}
