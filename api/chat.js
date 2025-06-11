export default async function handler(req, res) {
  // Handle CORS preflight requests
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(200).end();
  }

  // Only allow POST requests for actual chat
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const allowedOrigins = [
    'https://personalai.sharepoint.com',
    'https://chatbot-config-delta.vercel.app',
  ];
  
  const origin = req.headers.origin;
  const referer = req.headers.referer;
  
  // Check if request is from allowed domain
  const isValidOrigin = allowedOrigins.some(allowed => 
    origin?.includes(allowed.replace('https://', '')) || 
    referer?.includes(allowed.replace('https://', ''))
  );
  
  // Set CORS headers for valid origins
  if (isValidOrigin || origin?.includes('vercel.app')) {
    res.setHeader('Access-Control-Allow-Origin', origin || '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  
  // Block unauthorized domains
  if (!isValidOrigin && !origin?.includes('vercel.app')) {
    return res.status(403).json({ 
      error: 'Access denied',
      message: 'This chatbot is only available on authorized domains' 
    });
  }

try {
  console.log('Making request to Personal.ai API...');
  
  // Make the request to Personal.ai API with correct format
  const response = await fetch(process.env.ACTUAL_API_ENDPOINT, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': process.env.ACTUAL_API_KEY,
    },
    body: JSON.stringify({
      Text: req.body.Text || '',
      UserName: req.body.UserName || 'Visitor',
      SourceName: req.body.SourceName || 'Chatbot',
      SessionId: req.body.SessionId || `session_${Date.now()}`,
      DomainName: req.body.DomainName || 'default',
      is_draft: req.body.is_draft || false,
    }),
  });

  console.log('Personal.ai API response status:', response.status);

  if (!response.ok) {
    const errorText = await response.text();
    console.error('Personal.ai API error:', errorText);
    throw new Error(`Personal.ai API responded with status: ${response.status}`);
  }

  // Handle streaming response from Personal.ai
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
    message: 'Failed to process chat request',
    details: error.message
  });
}
}
