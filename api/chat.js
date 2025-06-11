export default async function handler(req, res) {
  const allowedOrigins = [
    'https://personalai.sharepoint.com',

  ];
  
  const origin = req.headers.origin;
  const referer = req.headers.referer;
  
  const isValidOrigin = allowedOrigins.some(allowed => 
    origin?.includes(allowed) || referer?.includes(allowed)
  );
  
  if (!isValidOrigin) {
    console.log(`Blocked request from: ${origin || 'unknown'}, referer: ${referer || 'unknown'}`);
    return res.status(403).json({ 
      error: 'Access denied',
      message: 'This chatbot is only available on authorized domains' 
    });
  }

  const userAgent = req.headers['user-agent'];
  const spHeaders = [
    req.headers['x-sharepoint-health-score'],
    req.headers['sprequestguid'],
    req.headers['x-forms_based_auth_accepted']
  ];
  
  const isLikelySharePoint = spHeaders.some(header => header !== undefined) ||
    userAgent?.includes('SharePoint') ||
    referer?.includes('sharepoint.com');

  const clientDomain = new URL(referer || origin || 'unknown').hostname;
  const rateLimitKey = `${clientDomain}-${Math.floor(Date.now() / 3600000)}`;
  
  const requestCount = global.requestCounts?.get(rateLimitKey) || 0;
  if (requestCount > 500) { 
    return res.status(429).json({ 
      error: 'Rate limit exceeded',
      message: 'Too many requests from this domain' 
    });
  }
  
  if (!global.requestCounts) global.requestCounts = new Map();
  global.requestCounts.set(rateLimitKey, requestCount + 1);

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { Text, UserName, SourceName, SessionId, DomainName } = req.body;

  if (!Text && Text !== '') {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  if (Text.length > 2000) {
    return res.status(400).json({ error: 'Message too long' });
  }

  console.log(`Chat request from: ${clientDomain}, User: ${UserName}, Text length: ${Text.length}`);

  try {
    const response = await fetch(process.env.ACTUAL_API_ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ACTUAL_API_KEY,
      },
      body: JSON.stringify({
        Text,
        UserName: UserName || 'SharePoint User',
        SourceName: SourceName || 'SharePoint Chatbot',
        SessionId,
        DomainName: process.env.DOMAIN_NAME,
        is_draft: false,
      }),
    });

    if (!response.ok) {
      throw new Error(`API responded with status: ${response.status}`);
    }

    // Handle streaming response
    res.setHeader('Content-Type', 'text/plain');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      
      const chunk = decoder.decode(value);
      res.write(chunk);
    }
    
    res.end();

  } catch (error) {
    console.error('API Error:', error);
    res.status(500).json({ 
      error: 'Internal server error',
      message: 'Failed to process chat request'
    });
  }
}
