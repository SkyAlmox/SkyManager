import React, { useState, useRef, useEffect } from 'react';
import { ChatMessage } from '../types';
import { sendMessageToGemini } from '../services/geminiService';

const Assistant: React.FC = () => {
  const [isOpen, setIsOpen] = useState(false);
  const [messages, setMessages] = useState<ChatMessage[]>([
    { role: 'model', text: 'Bem-vindo à SKYMÍDIA. Como posso ajudar com seu projeto de comunicação visual hoje?', timestamp: Date.now() }
  ]);
  const [inputValue, setInputValue] = useState('');
  const [isThinking, setIsThinking] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [messages, isOpen]);

  const handleSend = async () => {
    if (!inputValue.trim()) return;

    const userMsg: ChatMessage = { role: 'user', text: inputValue, timestamp: Date.now() };
    setMessages(prev => [...prev, userMsg]);
    setInputValue('');
    setIsThinking(true);

    try {
      const history = messages.map(m => ({ role: m.role, text: m.text }));
      const responseText = await sendMessageToGemini(history, userMsg.text);

      const aiMsg: ChatMessage = { role: 'model', text: responseText, timestamp: Date.now() };
      setMessages(prev => [...prev, aiMsg]);
    } catch (error) {
        // Error handled in service
    } finally {
      setIsThinking(false);
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  return (
    <div className="fixed bottom-8 right-8 z-50 flex flex-col items-end font-sans">
      {isOpen && (
        <div className="bg-brand-bg rounded-none shadow-2xl shadow-black/20 w-[90vw] sm:w-[380px] h-[550px] mb-6 flex flex-col overflow-hidden border border-brand-hover/20 animate-slide-up-fade">
          {/* Header */}
          <div className="bg-brand-bg p-5 border-b border-brand-hover/20 flex justify-between items-center">
            <div className="flex items-center gap-3">
                <div className="w-2 h-2 bg-brand-hover rounded-full animate-pulse"></div>
                <span className="font-serif italic text-brand-text text-lg">Suporte</span>
            </div>
            <button onClick={() => setIsOpen(false)} className="text-brand-text/40 hover:text-brand-hover transition-colors">
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1} stroke="currentColor" className="w-6 h-6">
                <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
              </svg>
            </button>
          </div>

          {/* Chat Area */}
          <div className="flex-1 overflow-y-auto p-6 space-y-8 bg-brand-bg" ref={scrollRef}>
            {messages.map((msg, idx) => (
              <div key={idx} className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'}`}>
                <div
                  className={`max-w-[85%] p-5 text-sm leading-relaxed ${
                    msg.role === 'user'
                      ? 'bg-brand-hover text-brand-bg'
                      : 'bg-white/5 border border-brand-hover/10 text-brand-text shadow-sm'
                  }`}
                >
                  {msg.text}
                </div>
              </div>
            ))}
            {isThinking && (
               <div className="flex justify-start">
                 <div className="bg-white/5 border border-brand-hover/10 p-5 flex gap-1 items-center shadow-sm">
                   <div className="w-1.5 h-1.5 bg-brand-hover rounded-full animate-bounce"></div>
                   <div className="w-1.5 h-1.5 bg-brand-hover rounded-full animate-bounce delay-75"></div>
                   <div className="w-1.5 h-1.5 bg-brand-hover rounded-full animate-bounce delay-150"></div>
                 </div>
               </div>
            )}
          </div>

          {/* Input Area */}
          <div className="p-5 bg-brand-bg border-t border-brand-hover/20">
            <div className="flex gap-2 relative">
              <input
                type="text"
                value={inputValue}
                onChange={(e) => setInputValue(e.target.value)}
                onKeyDown={handleKeyPress}
                placeholder="Como podemos ajudar?"
                className="flex-1 bg-white/5 border border-brand-hover/20 focus:border-brand-hover px-4 py-3 text-sm outline-none transition-colors placeholder-brand-text/30 text-brand-text"
              />
              <button
                onClick={handleSend}
                disabled={!inputValue.trim() || isThinking}
                className="bg-brand-hover text-brand-bg px-4 hover:bg-brand-hover/80 transition-colors disabled:opacity-50"
              >
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-5 h-5">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M13.5 4.5L21 12m0 0l-7.5 7.5M21 12H3" />
                </svg>
              </button>
            </div>
          </div>
        </div>
      )}

      <button
        onClick={() => setIsOpen(!isOpen)}
        className="bg-brand-hover text-brand-bg w-14 h-14 flex items-center justify-center rounded-full shadow-xl hover:scale-105 transition-all duration-500 z-50"
      >
        {isOpen ? (
             <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1} stroke="currentColor" className="w-6 h-6">
                <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 8.25l-7.5 7.5-7.5-7.5" />
             </svg>
        ) : (
            <span className="font-serif italic text-lg">IA</span>
        )}
      </button>
    </div>
  );
};

export default Assistant;