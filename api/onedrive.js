// api/onedrive.js
// Vercel Serverless Function to proxy Microsoft OneDrive Excel (.xlsx) downloads bypassing CORS

module.exports = async (req, res) => {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const url = (req.query && req.query.url) || (req.body && req.body.url);
  if (!url) {
    return res.status(400).json({ error: 'Thiếu tham số url của file OneDrive.' });
  }

  try {
    let currentUrl = String(url).trim();
    if (!currentUrl.includes('download=1')) {
      currentUrl += (currentUrl.includes('?') ? '&' : '?') + 'download=1';
    }

    const cookies = {};
    let redirectCount = 0;
    const maxRedirects = 10;

    while (redirectCount < maxRedirects) {
      redirectCount++;
      const headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'vi,en-US;q=0.9,en;q=0.8'
      };

      const cookieEntries = Object.entries(cookies);
      if (cookieEntries.length > 0) {
        headers['Cookie'] = cookieEntries.map(([k, v]) => `${k}=${v}`).join('; ');
      }

      const response = await fetch(currentUrl, {
        method: 'GET',
        headers,
        redirect: 'manual'
      });

      // Collect Set-Cookie
      let setCookieList = [];
      if (typeof response.headers.getSetCookie === 'function') {
        setCookieList = response.headers.getSetCookie();
      } else if (response.headers.raw && typeof response.headers.raw === 'function') {
        setCookieList = response.headers.raw()['set-cookie'] || [];
      } else {
        const raw = response.headers.get('set-cookie');
        if (raw) setCookieList = [raw];
      }

      for (const sc of setCookieList) {
        if (!sc) continue;
        const part = sc.split(';')[0].trim();
        const eqIdx = part.indexOf('=');
        if (eqIdx !== -1) {
          cookies[part.slice(0, eqIdx).trim()] = part.slice(eqIdx + 1).trim();
        }
      }

      const status = response.status;
      if (status >= 300 && status < 400) {
        const loc = response.headers.get('location');
        if (!loc) throw new Error('Redirect without Location header from OneDrive');
        currentUrl = new URL(loc, currentUrl).toString();
        continue;
      }

      if (status === 200) {
        const arrayBuffer = await response.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="checklist.xlsx"');
        res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
        return res.status(200).send(buffer);
      }

      throw new Error(`OneDrive phản hồi mã lỗi HTTP ${status}`);
    }

    throw new Error('Quá nhiều chuyển hướng từ liên kết OneDrive');
  } catch (err) {
    console.error('OneDrive proxy error:', err);
    return res.status(500).json({ error: err.message || 'Lỗi khi tải file từ OneDrive.' });
  }
};

module.exports.default = module.exports;
