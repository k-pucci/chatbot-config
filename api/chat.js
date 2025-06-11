export default async function handler(req, res) {
  // CORS and Security Headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Auth-Token');

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // Only allow POST requests
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // Security Validation - Check Referrer/Origin
  const referer = req.headers.referer || req.headers.referrer || '';
  const origin = req.headers.origin || '';
  const userAgent = req.headers['user-agent'] || '';
  
  const allowedDomains = [
    'sharepoint.com',
    'personalai.sharepoint.com'
  ];
  
  const isAuthorizedReferrer = allowedDomains.some(domain => 
    referer.includes(domain) || origin.includes(domain)
  );
  
  // Optional: Check for SharePoint-specific patterns
  const isFromSharePoint = 
    referer.includes('sharepoint.com') ||
    origin.includes('sharepoint.com') ||
    userAgent.includes('SharePoint') ||
    req.headers['x-sharepoint-context'];
  
  // Block unauthorized access
  if (!isAuthorizedReferrer && !isFromSharePoint) {
    console.log('Blocked access from:', { referer, origin, userAgent });
    return res.status(403).json({ 
      error: 'Access denied: This service is only available from authorized SharePoint domains' 
    });
  }

  // Optional: Token-based authentication (uncomment to enable)
  /*
  const authToken = req.headers['x-auth-token'];
  const expectedToken = process.env.SHAREPOINT_AUTH_TOKEN;
  
  if (!authToken || authToken !== expectedToken) {
    return res.status(403).json({ 
      error: 'Access denied: Invalid or missing authentication token' 
    });
  }
  */

  // Validate request body
  const { Text, UserName, SourceName, SessionId, DomainName, is_draft } = req.body;
  
  if (!Text && Text !== "") {
    return res.status(400).json({ error: 'Text is required' });
  }
  
  if (!DomainName) {
    return res.status(400).json({ error: 'DomainName is required' });
  }

  // Personal.ai API configuration from environment variables
  const PERSONAL_AI_API_KEY = process.env.PERSONAL_AI_API_KEY;
  const PERSONAL_AI_API_URL = process.env.PERSONAL_AI_API_URL || 'https://api.personal.ai/v1/message';
  
  if (!PERSONAL_AI_API_KEY) {
    console.error('Personal.ai API key not configured');
    return res.status(500).json({ error: 'Service configuration error' });
  }

  try {
    // Make request to Personal.ai API
    const personalAiResponse = await fetch(PERSONAL_AI_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': PERSONAL_AI_API_KEY,
      },
      body: JSON.stringify({
        Text: Text,
        UserName: UserName || 'Visitor',
        SourceName: SourceName || 'Chatbot',
        SessionId: SessionId || `session_${Date.now()}`,
        DomainName: DomainName,
        is_draft: is_draft || false,
      }),
    });

    if (!personalAiResponse.ok) {
      console.error('Personal.ai API error:', personalAiResponse.status, personalAiResponse.statusText);
      return res.status(personalAiResponse.status).json({ 
        error: 'Failed to get response from AI service' 
      });
    }

    // Check if response is streaming
    const contentType = personalAiResponse.headers.get('content-type');
    const isStreaming = contentType && (
      contentType.includes('text/event-stream') || 
      contentType.includes('application/stream+json') ||
      contentType.includes('text/plain')
    );

    if (isStreaming) {
      // Handle streaming response
      res.setHeader('Content-Type', 'text/plain');
      res.setHeader('Cache-Control', 'no-cache');
      res.setHeader('Connection', 'keep-alive');

      const reader = personalAiResponse.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';

      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          
          // Keep the last incomplete line in buffer
          buffer = lines.pop() || '';
          
          for (const line of lines) {
            if (line.trim() === '') continue;
            
            try {
              let messageText = '';
              let hasFollowup = false;
              
              // Handle different response formats from Personal.ai
              if (line.startsWith('data: ')) {
                const personalAiData = line.slice(6).trim();
                
                if (personalAiData === '[DONE]' || personalAiData === '') {
                  // Send final message and break
                  const finalFormat = {
                    ai_message: '',
                    session_id: SessionId,
                    has_followup: false
                  };
                  res.write(`data: ${JSON.stringify(finalFormat)}\n\n`);
                  break;
                }
                
                try {
                  const personalAiJson = JSON.parse(personalAiData);
                  messageText = personalAiJson.text || personalAiJson.content || personalAiJson.message || personalAiJson.ai_message || '';
                  hasFollowup = personalAiJson.has_followup || false;
                } catch (parseError) {
                  // If not JSON, treat as plain text
                  messageText = personalAiData;
                }
              } else {
                // Try parsing as direct JSON
                try {
                  const personalAiJson = JSON.parse(line);
                  messageText = personalAiJson.text || personalAiJson.content || personalAiJson.message || personalAiJson.ai_message || '';
                  hasFollowup = personalAiJson.has_followup || false;
                } catch (parseError) {
                  // If not JSON, treat as plain text
                  messageText = line.trim();
                }
              }
              
              // Transform to frontend's expected format
              const frontendFormat = {
                ai_message: messageText,
                session_id: SessionId,
                has_followup: hasFollowup
              };
              
              // Send in Server-Sent Events format that frontend expects
              res.write(`data: ${JSON.stringify(frontendFormat)}\n\n`);
              
            } catch (error) {
              console.error('Error processing streaming line:', error, 'Line:', line);
              // Continue processing other lines
            }
          }
        }
        
        // Process any remaining buffer
        if (buffer.trim()) {
          try {
            let messageText = buffer.trim();
            let hasFollowup = false;
            
            if (buffer.startsWith('data: ')) {
              const personalAiData = buffer.slice(6).trim();
              try {
                const personalAiJson = JSON.parse(personalAiData);
                messageText = personalAiJson.text || personalAiJson.content || personalAiJson.message || personalAiJson.ai_message || '';
                hasFollowup = personalAiJson.has_followup || false;
              } catch (parseError) {
                messageText = personalAiData;
              }
            } else {
              try {
                const personalAiJson = JSON.parse(buffer);
                messageText = personalAiJson.text || personalAiJson.content || personalAiJson.message || personalAiJson.ai_message || '';
                hasFollowup = personalAiJson.has_followup || false;
              } catch (parseError) {
                messageText = buffer.trim();
              }
            }
            
            const frontendFormat = {
              ai_message: messageText,
              session_id: SessionId,
              has_followup: hasFollowup
            };
            
            res.write(`data: ${JSON.stringify(frontendFormat)}\n\n`);
          } catch (error) {
            console.error('Error processing remaining buffer:', error);
          }
        }
        
      } catch (streamError) {
        console.error('Streaming error:', streamError);
        // Send error message to frontend
        const errorFormat = {
          ai_message: 'I apologize, but I encountered an error while processing your request. Please try again.',
          session_id: SessionId,
          has_followup: false
        };
        res.write(`data: ${JSON.stringify(errorFormat)}\n\n`);
      }
      
      res.end();
      
    } else {
      // Handle non-streaming response
      const personalAiData = await personalAiResponse.json();
      
      // Transform response to match frontend expectations
      const responseData = {
        ai_message: personalAiData.text || personalAiData.content || personalAiData.message || personalAiData.ai_message || 'No response available',
        session_id: SessionId,
        has_followup: personalAiData.has_followup || false
      };
      
      // Send as Server-Sent Events format for consistency
      res.setHeader('Content-Type', 'text/plain');
      res.write(`data: ${JSON.stringify(responseData)}\n\n`);
      res.end();
    }

  } catch (error) {
    console.error('API Handler Error:', error);
    
    // Send error response in expected format
    res.setHeader('Content-Type', 'text/plain');
    const errorResponse = {
      ai_message: 'I apologize, but I\'m having trouble responding right now. Please try again.',
      session_id: SessionId,
      has_followup: false
    };
    res.write(`data: ${JSON.stringify(errorResponse)}\n\n`);
    res.end();
  }
}
