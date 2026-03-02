import React from 'react';
import { motion } from 'framer-motion';
import { JOURNAL_ARTICLES } from '../constants';
import { JournalArticle } from '../types';

interface JournalProps {
  onArticleClick: (article: JournalArticle) => void;
}

const Journal: React.FC<JournalProps> = ({ onArticleClick }) => {
  const containerVariants = {
    hidden: { opacity: 0 },
    visible: {
      opacity: 1,
      transition: {
        staggerChildren: 0.15,
      },
    },
  };

  const itemVariants = {
    hidden: { opacity: 0, y: 30 },
    visible: { 
      opacity: 1, 
      y: 0,
      transition: { duration: 0.8, ease: [0.21, 0.47, 0.32, 0.98] }
    },
  };

  return (
    <section id="journal" className="bg-brand-bg py-32 px-6 md:px-12">
      <div className="max-w-[1800px] mx-auto">
        <motion.div 
          initial={{ opacity: 0, x: -20 }}
          whileInView={{ opacity: 1, x: 0 }}
          viewport={{ once: true }}
          transition={{ duration: 0.8 }}
          className="flex flex-col md:flex-row justify-between items-start md:items-end mb-20 pb-8 border-b border-brand-hover/20"
        >
            <div>
                <h2 className="text-4xl md:text-6xl font-serif text-brand-text">Produtos</h2>
            </div>
        </motion.div>

        <motion.div 
          variants={containerVariants}
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-50px" }}
          className="grid grid-cols-1 md:grid-cols-3 gap-12"
        >
            {JOURNAL_ARTICLES.filter(a => a.id !== 7).map((article) => (
                <motion.div 
                  key={article.id} 
                  variants={itemVariants}
                  className="group cursor-pointer flex flex-col text-left bg-brand-dark/5 p-8 rounded-sm border border-brand-hover/5 hover:border-brand-hover/20 transition-all duration-500" 
                  onClick={() => onArticleClick(article)}
                >
                    <div className="w-full aspect-[4/3] overflow-hidden mb-8 bg-brand-dark/5 rounded-sm">
                        <img 
                            src={article.image} 
                            alt={article.title} 
                            className="w-full h-full object-cover transition-transform duration-1000 group-hover:scale-105 grayscale-[0.2] group-hover:grayscale-0"
                        />
                    </div>
                    <div className="flex flex-col flex-1 text-left">
                        <span className="text-xs font-medium uppercase tracking-widest text-brand-text/40 mb-3">{article.date}</span>
                        {article.icon ? (
                            <div className="h-24 mb-6 flex items-center">
                                <img 
                                    src={article.icon} 
                                    alt={article.title} 
                                    className="h-full object-contain transition-transform duration-500 group-hover:scale-110"
                                    referrerPolicy="no-referrer"
                                />
                            </div>
                        ) : (
                            <h3 className="text-2xl font-serif text-brand-text mb-4 leading-tight group-hover:text-brand-hover transition-all">{article.title}</h3>
                        )}
                        <p className="text-brand-text font-light leading-relaxed">{article.excerpt}</p>
                    </div>
                </motion.div>
            ))}
        </motion.div>
      </div>
    </section>
  );
};

export default Journal;
