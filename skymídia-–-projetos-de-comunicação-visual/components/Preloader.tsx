import React, { useEffect, useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';

interface PreloaderProps {
  onComplete: () => void;
}

const Preloader: React.FC<PreloaderProps> = ({ onComplete }) => {
  const [isFinished, setIsFinished] = useState(false);
  const [timerId, setTimerId] = useState<NodeJS.Timeout | null>(null);
  
  const fullText = "OLÁ !\nSomos a SKYMÍDIA\nTransformamos matérias-primas comuns em projetos de Comunicação Visual";
  const lines = fullText.split("\n");
  
  useEffect(() => {
    const timer = setTimeout(() => {
      setIsFinished(true);
      setTimeout(onComplete, 1200);
    }, 7000);
    
    setTimerId(timer);
    return () => clearTimeout(timer);
  }, [onComplete]);

  const handleSkip = () => {
    if (timerId) clearTimeout(timerId);
    setIsFinished(true);
    setTimeout(onComplete, 1200);
  };

  const containerVariants = {
    initial: { opacity: 1, scale: 1, filter: "blur(0px)" },
    exit: {
      opacity: 0,
      scale: 3,
      filter: "blur(12px)",
      transition: { 
        duration: 1.2, 
        ease: [0.4, 0, 0.2, 1]
      }
    }
  };

  const lineVariants = {
    hidden: { opacity: 0, y: 20, filter: 'blur(10px)' },
    visible: { 
      opacity: 1, 
      y: 0, 
      filter: 'blur(0px)',
      transition: { duration: 1, ease: [0.215, 0.61, 0.355, 1] }
    }
  };

  return (
    <AnimatePresence>
      {!isFinished && (
        <motion.div
          variants={containerVariants}
          initial="initial"
          exit="exit"
          className="fixed inset-0 z-[999] flex items-center justify-center overflow-hidden bg-black"
        >
          <motion.button
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 1, duration: 1 }}
            onClick={handleSkip}
            className="absolute top-8 right-8 z-[1000] text-[10px] md:text-xs font-sans font-bold uppercase tracking-[0.3em] text-brand-text/40 hover:text-brand-hover transition-colors duration-300 flex items-center gap-2 group"
          >
            PULAR INTRODUÇÃO
            <svg 
              className="w-4 h-4 group-hover:translate-x-1 transition-transform" 
              fill="none" 
              viewBox="0 0 24 24" 
              stroke="currentColor"
            >
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 5l7 7-7 7M5 5l7 7-7 7" />
            </svg>
          </motion.button>

          <div className="relative z-20 w-full md:w-2/3 text-center px-6">
            <motion.div 
              initial="hidden"
              animate="visible"
              transition={{ staggerChildren: 0.8, delayChildren: 0.5 }}
            >
              {lines.map((line, lineIdx) => (
                <div key={lineIdx} className="overflow-hidden mb-4">
                  <motion.p 
                    variants={lineVariants}
                    className={`text-2xl md:text-4xl lg:text-5xl font-light leading-tight tracking-tight ${
                      line.includes('SKYMÍDIA') ? 'text-brand-text' : 'text-brand-text/80'
                    }`}
                  >
                    {line.split('SKYMÍDIA').map((part, i, arr) => (
                      <React.Fragment key={i}>
                        {part}
                        {i < arr.length - 1 && (
                          <span className="font-bold text-brand-hover">SKYMÍDIA</span>
                        )}
                      </React.Fragment>
                    ))}
                  </motion.p>
                </div>
              ))}
            </motion.div>
            
            <motion.div 
              initial={{ scaleX: 0 }}
              animate={{ scaleX: 1 }}
              transition={{ duration: 6, ease: "linear" }}
              className="h-px bg-brand-hover/30 mt-12 origin-left w-full max-w-xs mx-auto"
            />
          </div>
        </motion.div>
      )}
    </AnimatePresence>
  );
};

export default Preloader;
