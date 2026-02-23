import React from 'react';
import { JournalArticle } from '../types';

interface JournalDetailProps {
  article: JournalArticle;
  onBack: () => void;
  onNavigateToContact?: () => void;
}

const JournalDetail: React.FC<JournalDetailProps> = ({ article, onBack, onNavigateToContact }) => {
  const isContactPage = article.id === 7;
  const bgClass = isContactPage ? 'bg-[#E6E7E9]' : 'bg-brand-bg';
  const cardBgClass = 'bg-[#727377]';
  const titleColorClass = 'text-[#039AAD]';
  // const textColorClass = 'text-brand-dark';
  const textColorClass = '!text-[#E6E7E9]';
  const proseClass = 'prose-slate';

  return (
    <div className={`min-h-screen ${bgClass} animate-fade-in-up`}>
       {/* Hero Image for Article - Full bleed to top so navbar sits on it */}
       <div className="w-full h-[50vh] md:h-[60vh] relative overflow-hidden">
          <img
             src={article.image}
             alt={article.title}
             className="w-full h-full object-cover"
          />
          <div className="absolute inset-0 bg-brand-bg/20"></div>
       </div>

       <div className="max-w-6xl mx-auto px-6 md:px-12 -mt-32 relative z-10 pb-32">
          <div className={`${cardBgClass} p-8 md:p-16 shadow-xl shadow-black/20 border border-brand-hover/10`}>
             <div className="flex justify-between items-center mb-12 border-b border-brand-hover/20 pb-8">
                <button
                  onClick={onBack}
                  className={`group flex items-center gap-2 text-xs font-medium uppercase tracking-widest ${isContactPage ? 'text-brand-dark/60' : 'text-brand-text/60'} hover:text-brand-hover transition-colors`}
                >
                  <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-4 h-4 group-hover:-translate-x-1 transition-transform">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 19.5L8.25 12l7.5-7.5" />
                  </svg>
                  Voltar
                </button>
                <span className={`text-xs font-medium uppercase tracking-widest ${isContactPage ? 'text-brand-dark/60' : 'text-brand-text/60'}`}>{article.date}</span>
             </div>

             {/* <h1 className={`text-4xl md:text-6xl font-serif ${titleColorClass} mb-12 leading-tight text-center`}>
               {article.title}
             </h1> */}
             <h1 className={`text-4xl md:text-6xl font-serif ${titleColorClass} mb-12 leading-tight text-center`}>
                {article.titleImageUrl ? (
                  <img
                    src={article.titleImageUrl}
                    alt={article.title}
                    className="h-64 md:h-80 object-contain mx-auto"
                  />
                ) : (
                  article.title
                )}
              </h1>

             <div className="prose prose-lg mx-auto font-light leading-loose !text-[#E6E7E9]">
                {article.content}
              </div>

             {!isContactPage && (
               <div className="mt-12 flex justify-center">
                 <button
                   onClick={() => {
                     if (onNavigateToContact) {
                       onNavigateToContact();
                     }
                   }}
                   className="bg-[#039AAD] text-white px-12 py-4 rounded-sm uppercase tracking-widest font-bold hover:bg-[#039AAD]/80 transition-all shadow-lg shadow-[#039AAD]/20"
                 >
                   Solicitar Orçamento
                 </button>
               </div>
             )}

          </div>
       </div>
    </div>
  );
};

export default JournalDetail;
