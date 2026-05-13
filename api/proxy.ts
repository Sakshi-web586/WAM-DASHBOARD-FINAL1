import type { VercelRequest, VercelResponse } from '@vercel/node';

export default async function handler(req: VercelRequest, res: VercelResponse) {
  const targetUrl = req.query.url as string;

  if (!targetUrl) {
    return res.status(400).json({ error: 'Missing url parameter' });
  }

  try {
    const method = req.method || 'GET';
    const headers: Record<string, string> = {
      'User-Agent': 'Vercel Proxy',
    };

    if (req.headers['content-type']) {
      headers['Content-Type'] = req.headers['content-type'] as string;
    }

    const options: RequestInit = {
      method,
      headers,
      redirect: 'follow',
    };

    if (method !== 'GET' && method !== 'HEAD' && req.body) {
      options.body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
      if (!headers['Content-Type']) {
        headers['Content-Type'] = 'application/json';
      }
    }

    const response = await fetch(targetUrl, options);
    const data = await response.text();

    res.status(response.status);
    const contentType = response.headers.get('content-type');
    if (contentType) {
      res.setHeader('Content-Type', contentType);
    }
    
    return res.send(data);
  } catch (error: any) {
    console.error('Proxy error:', error);
    return res.status(500).json({
      error: 'Proxy failed to reach Google Apps Script',
      message: error.message,
    });
  }
}
