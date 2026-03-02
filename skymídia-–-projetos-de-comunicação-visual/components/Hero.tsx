import React from 'react';
import { motion } from 'framer-motion';
import CircuitBackground from './CircuitBackground';
import { JOURNAL_ARTICLES } from '../constants';
import { JournalArticle } from '../types';

interface HeroProps {
  onNavClick?: (e: React.MouseEvent<HTMLAnchorElement>, targetId: string) => void;
  onProductClick: (article: JournalArticle) => void;
}

const Hero: React.FC<HeroProps> = ({ onNavClick, onProductClick }) => {
  const products = JOURNAL_ARTICLES.filter(a => a.id !== 7);

  const containerVariants = {
    hidden: { opacity: 0 },
    visible: {
      opacity: 1,
      transition: {
        staggerChildren: 0.1,
        delayChildren: 0.2,
      }
    }
  };

  const itemVariants = {
    hidden: { opacity: 0, y: 40, scale: 0.8, filter: 'blur(10px)' },
    visible: { 
      opacity: 1, 
      y: 0,
      scale: 1,
      filter: 'blur(0px)',
      transition: { 
        type: "spring",
        stiffness: 100,
        damping: 20,
        mass: 1,
        duration: 0.8 
      }
    }
  };

  return (
    <section id="journal" className="relative w-full min-h-screen flex items-center justify-center pt-32 pb-20">
      <div className="w-[85vw] relative z-10 flex items-center">
        <motion.div 
          variants={containerVariants}
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 md:gap-6 w-full"
        >
          {products.map((product) => (
            <motion.div
              key={product.id}
              variants={itemVariants}
              whileHover={{ 
                y: -10, 
                scale: 1.02,
                boxShadow: "0 20px 40px rgba(0,0,0,0.2)",
                borderColor: "rgba(21, 240, 219, 0.4)"
              }}
              className="group relative flex flex-col items-center justify-center bg-brand-dark/5 p-8 md:p-10 rounded-sm border border-brand-hover/5 transition-all duration-500 cursor-pointer overflow-hidden text-center min-h-[320px]"
              onClick={() => onProductClick(product)}
            >
              <motion.div
                animate={{
                  y: [0, -4, 0],
                }}
                transition={{
                  duration: 4 + Math.random() * 2,
                  repeat: Infinity,
                  ease: "easeInOut",
                  delay: Math.random() * 2
                }}
                className="w-full h-full flex flex-col items-center justify-center"
              >
                <div className="h-20 md:h-24 mb-6 flex items-center justify-center">
                  {product.icon ? (
                    <img 
                      src={product.icon} 
                      alt={product.title} 
                      className="h-full object-contain transition-transform duration-700 group-hover:scale-110"
                      referrerPolicy="no-referrer"
                    />
                  ) : null}
                </div>
                
                <p className="text-brand-text/60 font-light leading-relaxed text-xs md:text-sm max-w-[280px]">
                  {product.excerpt}
                </p>
              </motion.div>

              <div className="absolute top-4 right-4 opacity-0 group-hover:opacity-100 transition-opacity duration-300">
                <div className="w-5 h-5 rounded-full border border-brand-hover flex items-center justify-center text-brand-hover">
                  <svg className="w-2.5 h-2.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                  </svg>
                </div>
              </div>
            </motion.div>
          ))}
        </motion.div>
      </div>

      <div className="absolute top-0 left-0 w-full h-full pointer-events-none z-0 opacity-10">
        <motion.div 
          animate={{ 
            scale: [1, 1.2, 1],
            rotate: [0, 90, 0],
            x: [0, 100, 0],
            y: [0, -50, 0]
          }}
          transition={{ duration: 20, repeat: Infinity, ease: "linear" }}
          className="absolute -top-1/4 -left-1/4 w-1/2 h-1/2 bg-brand-hover blur-[150px] rounded-full"
        />
        <motion.div 
          animate={{ 
            scale: [1.2, 1, 1.2],
            rotate: [0, -90, 0],
            x: [0, -100, 0],
            y: [0, 50, 0]
          }}
          transition={{ duration: 25, repeat: Infinity, ease: "linear" }}
          className="absolute -bottom-1/4 -right-1/4 w-1/2 h-1/2 bg-brand-hover blur-[150px] rounded-full"
        />
      </div>
    </section>
  );
};

export default Hero;
