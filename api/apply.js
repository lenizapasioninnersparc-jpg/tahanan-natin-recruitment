// Vercel server-side proxy: browser -> /api/apply -> Google Apps Script.
// GAS_WEB_APP_URL and GAS_API_SECRET must be configured in Vercel Environment Variables.

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed.' });
  }

  const gasUrl = process.env.GAS_WEB_APP_URL;
  const apiSecret = process.env.GAS_API_SECRET;

  if (!gasUrl || !apiSecret) {
    return res.status(500).json({
      success: false,
      message: 'The recruitment form service is not configured yet.'
    });
  }

  // Accept only the standard Apps Script deployed web-app URL.
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(gasUrl)) {
    return res.status(500).json({
      success: false,
      message: 'GAS_WEB_APP_URL must be the deployed Apps Script /exec URL.'
    });
  }

  if (!req.body || typeof req.body !== 'object' || Array.isArray(req.body)) {
    return res.status(400).json({ success: false, message: 'Invalid application data.' });
  }

  try {
    // Node fetch follows Google's ContentService redirect automatically.
    const upstream = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        apiSecret: apiSecret,
        data: req.body
      }),
      redirect: 'follow'
    });

    const raw = await upstream.text();
    let result;
    try {
      result = JSON.parse(raw);
    } catch (_) {
      throw new Error('Google Apps Script returned an unreadable response. Check its deployment and execution logs.');
    }

    if (!upstream.ok || !result.success) {
      return res.status(502).json({
        success: false,
        message: result.message || 'The application could not be saved.'
      });
    }

    return res.status(200).json(result);
  } catch (err) {
    console.error('Recruitment backend proxy error:', err);
    return res.status(502).json({
      success: false,
      message: err.message || 'Could not reach the recruitment backend.'
    });
  }
};
