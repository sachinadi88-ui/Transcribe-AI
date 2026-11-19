import { GoogleGenAI, ThinkingLevel } from '@google/genai';
import type { Request, Response } from 'express';

export function removeTimestamps(text: string): string {
  return text
    // 1. Remove bracketed or parenthesized timestamps or ranges: [00:00], [00:00 - 00:15], (0:12), (01:23:45)
    .replace(/[\[\(]\s*\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:\s*[-–—~to]\s*\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?)?\s*[\]\)]/gi, '')
    // 2. Remove line-leading timestamps like "00:00: ", "01:23 - ", "0:15 "
    .replace(/^\s*\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?(?:\s*[-–—~to]\s*\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?)?\s*[:\-–—]?\s*/gm, '')
    // 3. Remove inline timestamps like " 00:15:23 " or " 02:45 "
    .replace(/(?<=\s|^)\d{1,2}:[0-5]\d(?::[0-5]\d)?(?:\.\d+)?(?=\s|[:\-–—]|$)/gm, '')
    // 4. Clean up dangling colons or dashes from stripped timestamps
    .replace(/\s+([,.:;?!])/g, '$1')
    // 5. Clean up multiple spaces
    .replace(/[ \t]+/g, ' ')
    // 6. Clean up empty lines
    .split('\n')
    .map((line) => line.trim())
    .filter((line, i, arr) => line.length > 0 || (i > 0 && arr[i - 1].length > 0))
    .join('\n')
    .trim();
}

const MODEL_ATTEMPTS: Array<{
  model: string;
  apiVersion?: string;
  supportsSystemInstruction?: boolean;
  useLowThinking?: boolean;
}> = [
  { model: 'gemini-3.5-transcribe', apiVersion: 'v1beta', supportsSystemInstruction: false },
  { model: 'gemini-4-hydrogen', supportsSystemInstruction: true, useLowThinking: true },
];

function isTransientOrNotFoundError(err: any): boolean {
  const msg = String(err?.message || '');
  const status = err?.status;
  return (
    status === 503 ||
    status === 404 ||
    msg.includes('503') ||
    msg.includes('404') ||
    msg.includes('high demand') ||
    msg.includes('UNAVAILABLE') ||
    msg.includes('NOT_FOUND')
  );
}

export async function handleTranscription(req: Request, res: Response) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed. Use POST.' });
  }

  try {
    const { base64Data, mimeType, fileName } = req.body || {};

    if (!base64Data) {
      return res.status(400).json({ error: 'Missing audio data for transcription.' });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({
        error: 'Gemini API key is not configured on the server. Please configure GEMINI_API_KEY.',
      });
    }

    let finalMimeType = mimeType;
    if (!finalMimeType) {
      if (fileName?.endsWith('.mp3')) finalMimeType = 'audio/mpeg';
      else if (fileName?.endsWith('.wav')) finalMimeType = 'audio/wav';
      else if (fileName?.endsWith('.m4a')) finalMimeType = 'audio/mp4';
      else finalMimeType = 'audio/mpeg';
    }

    let lastError: any = null;
    let transcriptionText = '';

    for (const attempt of MODEL_ATTEMPTS) {
      try {
        const ai = new GoogleGenAI({
          apiKey,
          ...(attempt.apiVersion ? { apiVersion: attempt.apiVersion } : {}),
          httpOptions: {
            headers: {
              'User-Agent': 'aistudio-build',
            },
          },
        });

        const config: Record<string, any> = {};
        if (attempt.supportsSystemInstruction) {
          config.systemInstruction =
            'You are an expert audio transcriptionist. Transcribe the audio verbatim into continuous, clear text paragraphs. You must NEVER output timestamps, time codes, or duration markers (such as [00:00], 0:15, or time intervals) anywhere in your response. Output only the plain transcribed spoken words.';
        }
        if (attempt.useLowThinking) {
          config.thinkingConfig = { thinkingLevel: ThinkingLevel.LOW };
        }

        const response = await ai.models.generateContent({
          model: attempt.model,
          ...(Object.keys(config).length > 0 ? { config } : {}),
          contents: {
            parts: [
              {
                inlineData: {
                  mimeType: finalMimeType,
                  data: base64Data,
                },
              },
              {
                text: 'Transcribe the audio verbatim into clean text. Strictly DO NOT include any timestamps, time markers, or time ranges.',
              },
            ],
          },
        });

        if (response.text) {
          transcriptionText = removeTimestamps(response.text);
          break;
        }
      } catch (err: any) {
        lastError = err;
        if (isTransientOrNotFoundError(err)) {
          continue;
        }
        throw err;
      }
    }

    if (!transcriptionText) {
      if (lastError) throw lastError;
      return res.status(422).json({
        error: 'The AI was unable to generate a transcript for this file. It might be silent or containing unsupported content.',
      });
    }

    return res.status(200).json({ text: transcriptionText });
  } catch (error: any) {
    console.error('Server transcription error:', error);
    let message = error?.message || 'An error occurred during transcription.';
    let statusCode = 500;

    if (message.includes('API_KEY_INVALID') || message.includes('API key not valid') || message.includes('PERMISSION_DENIED')) {
      statusCode = 401;
      message = 'Invalid or unauthorized API Key. Please verify GEMINI_API_KEY in Settings > Secrets.';
    } else if (
      message.includes('503') ||
      message.includes('high demand') ||
      message.includes('UNAVAILABLE')
    ) {
      statusCode = 503;
      message =
        "Google's AI model is temporarily experiencing high server demand (503). Please wait a few moments and try again.";
    } else if (message.includes('Too Many Requests') || message.includes('429') || message.includes('RESOURCE_EXHAUSTED')) {
      statusCode = 429;
      message = 'Rate limit exceeded. Please wait a moment and try again.';
    } else if (message.includes('NOT_FOUND') || message.includes('404')) {
      statusCode = 404;
      message = 'The configured transcription model was not found. Please try again.';
    }

    return res.status(statusCode).json({ error: message });
  }
}

export default handleTranscription;
