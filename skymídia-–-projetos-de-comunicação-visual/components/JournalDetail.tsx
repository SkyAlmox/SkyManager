import React from 'react';
import { JournalArticle } from '../types';

interface JournalDetailProps {
  article: JournalArticle;
  onBack: () => void;
  onNavigateToContact?: () => void;
}

const JournalDetail: React.FC<JournalDetailProps> = ({ article, onBack, onNavigateToContact }) => {
  const isContactPage = article.id === 7;
  const bgClass = '';
  const cardBgClass = 'bg-brand-dark/5';
  const titleColorClass = 'text-brand-hover';
  const textColorClass = 'text-brand-text';
  const proseClass = 'prose-invert';

  return (
    <div className={`min-h-screen ${bgClass} animate-fade-in-up`}>
       <div className="w-full h-[50vh] md:h-[60vh] relative overflow-hidden">
          <img 
             src={article.image} 
             alt={article.title} 
             className="w-full h-full object-cover"
          />
          <div className="absolute inset-0 bg-brand-bg/40"></div>
       </div>

       <div className="max-w-6xl mx-auto px-6 md:px-12 -mt-32 relative z-10 pb-32">
          <div className={`${cardBgClass} p-8 md:p-16 backdrop-blur-md border border-brand-hover/10 rounded-sm`}>
             <div className="flex justify-between items-center mb-12 border-b border-brand-hover/20 pb-8">
                <button 
                  onClick={onBack}
                  className={`group flex items-center gap-2 text-xs font-medium uppercase tracking-widest text-brand-text hover:text-brand-hover transition-colors`}
                >
                  <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-4 h-4 group-hover:-translate-x-1 transition-transform">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 19.5L8.25 12l7.5-7.5" />
                  </svg>
                  Voltar
                </button>
                <span className={`text-xs font-medium uppercase tracking-widest text-brand-text`}>{article.date}</span>
             </div>

             <div className="flex justify-center mb-12">
               {article.icon ? (
                 <div className="h-24 md:h-32">
                   <img 
                     src={article.icon} 
                     alt={article.title} 
                     className="h-full object-contain"
                     referrerPolicy="no-referrer"
                   />
                 </div>
               ) : (
                 <h1 className={`text-4xl md:text-6xl font-serif ${titleColorClass} leading-tight text-center`}>
                   {article.title}
                 </h1>
               )}
             </div>

             <div className={`mx-auto font-light leading-loose text-brand-text`}>
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
                   className="bg-brand-hover text-brand-bg px-12 py-4 rounded-sm uppercase tracking-widest font-bold hover:bg-brand-hover/80 transition-all shadow-lg shadow-brand-hover/20"
                 >
                   FALE CONOSCO !
                 </button>
               </div>
             )}

          </div>
       </div>
    </div>
  );
};

export default JournalDetail;
