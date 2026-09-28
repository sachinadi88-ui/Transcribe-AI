/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useRef, useCallback, DragEvent, ChangeEvent } from 'react';
import { GoogleGenAI } from "@google/genai";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import { 
  Upload, 
  FileAudio, 
  Loader2, 
  CheckCircle2, 
  Copy, 
  Download,
  Trash2, 
  AlertCircle,
  Clock,
  Mic
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Recommended model for audio transcription - Using the latest stable flash alias
const MODEL_NAME = "gemini-flash-latest";

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isTranscribing, setIsTranscribing] = useState(false);
  const [transcription, setTranscription] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragActive, setDragActive] = useState(false);
  const [copied, setCopied] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDrag = useCallback((e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  }, []);

  const handleDrop = useCallback((e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const droppedFile = e.dataTransfer.files[0];
      validateAndSetFile(droppedFile);
    }
  }, []);

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      validateAndSetFile(e.target.files[0]);
    }
  };

  const validateAndSetFile = (selectedFile: File) => {
    setError(null);
    setTranscription(null);

    // Check file type
    if (!selectedFile.type.startsWith('audio/') && !selectedFile.name.endsWith('.mp3') && !selectedFile.name.endsWith('.wav')) {
      setError("Please upload a valid audio file (MP3, WAV, etc.).");
      return;
    }

    // Check file size (Gemini API limit for inline data is around 20MB, but let's be safe at 15MB)
    if (selectedFile.size > 15 * 1024 * 1024) {
      setError("File is too large. Please upload an audio file smaller than 15MB for better results.");
      return;
    }

    setFile(selectedFile);
  };

  const clearFile = () => {
    setFile(null);
    setTranscription(null);
    setError(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const copyToClipboard = async () => {
    if (transcription) {
      await navigator.clipboard.writeText(transcription);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  const downloadWordDoc = async () => {
    if (!transcription) return;

    setIsDownloading(true);
    try {
      // Split transcription by lines to preserve paragraphs
      const lines = transcription.split('\n');
      const paragraphs = lines.map((line) => {
        return new Paragraph({
          children: [
            new TextRun({
              text: line || " ",
              font: "Calibri",
              size: 24, // 12pt (docx uses half-points: 24 = 12pt)
            }),
          ],
          spacing: {
            after: line.trim() === "" ? 120 : 160,
            line: 360, // 1.5 line spacing
          },
        });
      });

      const doc = new Document({
        sections: [
          {
            properties: {},
            children: [
              new Paragraph({
                text: file?.name ? `Audio Transcription: ${file.name}` : "Audio Transcription",
                heading: HeadingLevel.HEADING_1,
                spacing: {
                  after: 240,
                },
              }),
              ...paragraphs,
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const baseName = file?.name ? file.name.replace(/\.[^/.]+$/, "") : "audio";
      link.href = url;
      link.download = `${baseName}_transcription.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Failed to generate Word document:", err);
      setError("Failed to generate Word document. Please try again.");
    } finally {
      setIsDownloading(false);
    }
  };

  const transcribeAudio = async () => {
    if (!file) return;

    setIsTranscribing(true);
    setError(null);

    try {
      // 1. Convert file to base64
      const reader = new FileReader();
      const base64Promise = new Promise<string>((resolve, reject) => {
        reader.onload = () => {
          const base64 = (reader.result as string).split(',')[1];
          resolve(base64);
        };
        reader.onerror = () => reject(new Error("Failed to read the audio file."));
        reader.readAsDataURL(file);
      });

      const base64Data = await base64Promise;

      // 2. Initialize Gemini API
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      
      // Map common mime types if browser reports them oddly
      let mimeType = file.type;
      if (!mimeType) {
        if (file.name.endsWith('.mp3')) mimeType = 'audio/mpeg';
        if (file.name.endsWith('.wav')) mimeType = 'audio/wav';
        if (file.name.endsWith('.m4a')) mimeType = 'audio/mp4';
      }

      // 3. Request transcription
      const response = await ai.models.generateContent({
        model: MODEL_NAME,
        contents: [
          {
            parts: [
              {
                inlineData: {
                  mimeType: mimeType || 'audio/mpeg',
                  data: base64Data,
                },
              },
              {
                text: "Provide a verbatim transcription of this audio. Do not include timestamps.",
              },
            ],
          },
        ],
      });

      if (!response.text) {
        throw new Error("The AI was unable to generate a transcript for this file. It might be silent or containing unsupported content.");
      }

      setTranscription(response.text);
    } catch (err: any) {
      console.error("Transcription error:", err);
      let errorMsg = "An error occurred during transcription.";
      
      if (err.message?.includes("API_KEY_INVALID")) {
        errorMsg = "Invalid API Key. Please check your project settings.";
      } else if (err.message?.includes("Model")) {
        errorMsg = `The selected model (${MODEL_NAME}) is currently unavailable or unsupported for this task.`;
      } else if (err.message?.includes("Too Many Requests")) {
        errorMsg = "Rate limit exceeded. Please wait a moment and try again.";
      } else {
        errorMsg = err.message || errorMsg;
      }
      
      setError(errorMsg);
    } finally {
      setIsTranscribing(false);
    }
  };

  const formatFileSize = (bytes: number) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-4 md:p-8 flex flex-col items-center">
      <header className="max-w-4xl w-full mb-12 text-center md:text-left flex flex-col md:flex-row md:items-end justify-between gap-6">
        <div>
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="flex items-center gap-3 mb-2 justify-center md:justify-start"
          >
            <div className="bg-indigo-600 p-2 rounded-xl text-white">
              <Mic size={24} />
            </div>
            <h1 className="text-3xl font-bold tracking-tight bg-gradient-to-r from-indigo-600 to-violet-600 bg-clip-text text-transparent">
              AudioTranscribe AI
            </h1>
          </motion.div>
          <p className="text-slate-500 font-medium">High-precision transcription powered by Gemini AI</p>
        </div>
        
        <div className="flex items-center gap-4 bg-white px-4 py-2 rounded-full border border-slate-200 shadow-sm self-center md:self-auto">
          <div className="flex items-center gap-2 text-xs font-mono text-slate-400">
            <span className="w-2 h-2 rounded-full bg-green-500 animate-pulse"></span>
            TRANSCRIPTION ENGINE: {MODEL_NAME === 'gemini-flash-latest' ? 'Gemini Flash' : MODEL_NAME}
          </div>
        </div>
      </header>

      <main className="max-w-4xl w-full grid grid-cols-1 gap-8">
        {!transcription ? (
          <motion.div
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="flex flex-col gap-6"
          >
            <div 
              onDragEnter={handleDrag}
              onDragLeave={handleDrag}
              onDragOver={handleDrag}
              onDrop={handleDrop}
              className={`
                relative border-2 border-dashed rounded-3xl p-12 text-center transition-all duration-300
                ${dragActive ? 'border-indigo-500 bg-indigo-50/50 scale-[1.01]' : 'border-slate-200 bg-white hover:border-indigo-300 hover:bg-slate-50/50'}
                ${file ? 'border-indigo-200 bg-indigo-50/20' : ''}
              `}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept="audio/*"
                onChange={handleFileChange}
                className="hidden"
              />
              
              <AnimatePresence mode="wait">
                {!file ? (
                  <motion.div 
                    key="upload-prompt"
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    exit={{ opacity: 0 }}
                    className="flex flex-col items-center"
                  >
                    <div className="w-16 h-16 bg-indigo-100 rounded-full flex items-center justify-center text-indigo-600 mb-6 group-hover:scale-110 transition-transform">
                      <Upload size={32} />
                    </div>
                    <h2 className="text-xl font-semibold mb-2">Upload your audio</h2>
                    <p className="text-slate-500 mb-8 max-w-xs mx-auto">
                      Drag and drop your audio file here, or click to browse. MP3, WAV, M4A supported.
                    </p>
                    <button 
                      onClick={() => fileInputRef.current?.click()}
                      className="px-8 py-3 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold rounded-xl shadow-lg shadow-indigo-200 transition-all hover:-translate-y-0.5 active:translate-y-0"
                    >
                      Select Audio File
                    </button>
                  </motion.div>
                ) : (
                  <motion.div 
                    key="selected-file"
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    exit={{ opacity: 0 }}
                    className="flex flex-col items-center"
                  >
                    <div className="w-16 h-16 bg-indigo-600 rounded-full flex items-center justify-center text-white mb-6">
                      <FileAudio size={32} />
                    </div>
                    <h2 className="text-xl font-semibold mb-2 truncate max-w-md">{file.name}</h2>
                    <div className="flex items-center gap-4 text-slate-500 mb-8 font-mono text-sm">
                      <span className="flex items-center gap-1.5"><Clock size={14}/> {formatFileSize(file.size)}</span>
                      <span className="w-1 h-1 rounded-full bg-slate-300"></span>
                      <span>{file.type.split('/')[1].toUpperCase()}</span>
                    </div>
                    
                    <div className="flex gap-4">
                      <button 
                        disabled={isTranscribing}
                        onClick={transcribeAudio}
                        className={`
                          px-8 py-3 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold rounded-xl shadow-lg shadow-indigo-200 transition-all flex items-center gap-2
                          ${isTranscribing ? 'opacity-50 cursor-not-allowed' : 'hover:-translate-y-0.5 active:translate-y-0'}
                        `}
                      >
                        {isTranscribing ? (
                          <>
                            <Loader2 size={20} className="animate-spin" />
                            Transcribing...
                          </>
                        ) : (
                          <>
                            Transcribe Now
                          </>
                        )}
                      </button>
                      <button 
                        disabled={isTranscribing}
                        onClick={clearFile}
                        className="px-4 py-3 bg-white border border-slate-200 text-slate-600 font-semibold rounded-xl hover:bg-slate-50 transition-all"
                      >
                        <Trash2 size={20} />
                      </button>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>

            {error && (
              <motion.div 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="bg-red-50 border border-red-100 p-4 rounded-2xl flex items-start gap-3 text-red-700"
              >
                <AlertCircle size={20} className="shrink-0 mt-0.5" />
                <div>
                  <p className="font-semibold text-sm">Action failed</p>
                  <p className="text-sm opacity-90">{error}</p>
                </div>
              </motion.div>
            )}
          </motion.div>
        ) : (
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white border border-slate-200 rounded-3xl overflow-hidden shadow-sm flex flex-col h-[600px]"
          >
            <div className="px-6 py-4 border-bottom border-slate-100 flex items-center justify-between bg-slate-50/50 sticky top-0 z-10 backdrop-blur-md">
              <div className="flex items-center gap-3">
                <CheckCircle2 size={18} className="text-green-500" />
                <h2 className="font-semibold text-slate-700">Transcription Result</h2>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button 
                  onClick={copyToClipboard}
                  className={`
                    p-2 rounded-lg transition-all flex items-center gap-2 text-sm font-medium
                    ${copied ? 'bg-green-100 text-green-700' : 'hover:bg-slate-200 text-slate-600'}
                  `}
                >
                  {copied ? <CheckCircle2 size={16} /> : <Copy size={16} />}
                  {copied ? 'Copied' : 'Copy Text'}
                </button>
                <button 
                  onClick={downloadWordDoc}
                  disabled={isDownloading}
                  className={`
                    p-2 rounded-lg transition-all flex items-center gap-2 text-sm font-medium hover:bg-slate-200 text-slate-600 disabled:opacity-50
                  `}
                  title="Download transcription as Word document (.docx)"
                >
                  {isDownloading ? <Loader2 size={16} className="animate-spin text-indigo-600" /> : <Download size={16} />}
                  {isDownloading ? 'Downloading...' : 'Download Word'}
                </button>
                <button 
                  onClick={() => setTranscription(null)}
                  className="p-2 hover:bg-slate-200 text-slate-600 rounded-lg transition-all flex items-center gap-2 text-sm font-medium"
                >
                  <Upload size={16} />
                  New Upload
                </button>
              </div>
            </div>
            
            <div className="p-8 flex-1 overflow-y-auto prose prose-slate max-w-none">
              <div className="whitespace-pre-wrap font-sans text-slate-700 leading-relaxed text-lg">
                {transcription}
              </div>
            </div>

            <div className="px-6 py-4 border-t border-slate-100 bg-slate-50/50 flex flex-wrap gap-4 items-center justify-between">
              <div className="flex items-center gap-3 text-sm text-slate-500 font-mono">
                <span className="flex items-center gap-1.5"><FileAudio size={14}/> {file?.name}</span>
                <span className="w-1 h-1 rounded-full bg-slate-300"></span>
                <span>{formatFileSize(file?.size || 0)}</span>
              </div>
              <p className="text-[10px] uppercase tracking-widest font-bold text-slate-400">Processed by Gemini Flash</p>
            </div>
          </motion.div>
        )}
      </main>

      <footer className="mt-auto py-8 text-slate-400 text-sm flex items-center gap-4">
        <span>AudioTranscribe AI &copy; 2026</span>
        <span className="w-1 h-1 rounded-full bg-slate-300"></span>
        <span>Secure & Private</span>
        <span className="w-1 h-1 rounded-full bg-slate-300"></span>
        <a href="#" className="hover:text-indigo-600 transition-colors">Privacy Policy</a>
      </footer>
    </div>
  );
}
