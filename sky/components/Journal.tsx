import React from 'react';
import { JOURNAL_ARTICLES } from '../constants';
import { JournalArticle } from '../types';

interface JournalProps {
  onArticleClick: (article: JournalArticle) => void;
}

const Journal: React.FC<JournalProps> = ({ onArticleClick }) => {
  return (
    <section id="journal" className="bg-brand-text py-32 px-6 md:px-12">
      <div className="max-w-[1800px] mx-auto">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-end mb-20 pb-8 border-b border-brand-dark/10">
            <div>
                <h2 className="text-4xl md:text-6xl font-serif text-brand-dark">Produtos</h2>
            </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-12">
            {JOURNAL_ARTICLES.filter(a => a.id !== 7).map((article) => (
                <div key={article.id} className="group cursor-pointer flex flex-col text-left" onClick={() => onArticleClick(article)}>
                    <div className="flex flex-col flex-1 text-left">
                        <span className="text-xs font-medium uppercase tracking-widest text-brand-dark/40 mb-3">{article.date}</span>
                        <h3 className="text-2xl font-serif text-brand-dark mb-4 leading-tight group-hover:underline decoration-brand-dark/30 underline-offset-4 transition-all">{article.title}</h3>
                        <p className="text-brand-dark/70 font-light leading-relaxed mb-6">{article.excerpt}</p>
                    </div>
                    <div className="w-full aspect-[4/3] overflow-hidden mb-8 bg-brand-dark/5">
                        <img
                            src={article.image}
                            alt={article.title}
                            className="w-full h-full object-cover transition-transform duration-1000 group-hover:scale-105 grayscale-[0.2] group-hover:grayscale-0"
                        />
                    </div>
                </div>
            ))}
        </div>
      </div>
    </section>
  );
};

export default Journal;
