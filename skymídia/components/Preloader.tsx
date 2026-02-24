import React, { useEffect, useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';

interface PreloaderProps {
  onComplete: () => void;
}

const Preloader: React.FC<PreloaderProps> = ({ onComplete }) => {
  const [isFinished, setIsFinished] = useState(false);

  const line1 = "OLÁ !";
  const line2Prefix = "Somos a ";
  const brand = "SKYMÍDIA";
  const line3 = "Transformamos matérias-primas comuns em projetos de Comunicação Visual.";

  const splitChars = (text: string) => text.split("");

  useEffect(() => {
    const timer = setTimeout(() => {
      setIsFinished(true);
      setTimeout(onComplete, 1000);
    }, 7500);

    return () => clearTimeout(timer);
  }, [onComplete]);

  const containerVariants = {
    hidden: { opacity: 0, scale: 0.8 },
    visible: { opacity: 1, scale: 1, transition: { duration: 0.5 } },
    exit: {
      opacity: 0,
      scale: 3,
      filter: "blur(8px)",
      transition: { duration: 1, ease: "easeInOut" }
    }
  };

  const charVariants = {
    hidden: { opacity: 0 },
    visible: { opacity: 1 },
    exit: {
      opacity: 0,
      y: () => Math.random() * 200 - 100,
      x: () => Math.random() * 200 - 100,
      transition: { duration: 1 }
    }
  };

  return (
    <AnimatePresence>
      {!isFinished && (
        <motion.div
          variants={containerVariants}
          initial="hidden"
          animate="visible"
          exit="exit"
          className="fixed inset-0 z-[999] flex items-center justify-center bg-brand-dark px-6"
        >
          <div className="w-2/3 text-center">

            {/* Linha 1 */}
            <motion.p
              className="text-6xl md:text-9xl lg:text-[10rem] font-light"
              initial="hidden"
              animate="visible"
              exit="exit"
              variants={{
                visible: {
                  transition: { staggerChildren: 0.08, delayChildren: 0.5 }
                }
              }}
            >
              {splitChars(line1).map((char, i) => (
                <motion.span key={i} variants={charVariants} style={{ color: '#E6E7E9' }}>
                  {char}
                </motion.span>
              ))}
            </motion.p>

            {/* Linha em branco */}
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>

            {/* Linha 2 */}
            <motion.p
              className="text-3xl md:text-5xl lg:text-6xl font-light"
              initial="hidden"
              animate="visible"
              exit="exit"
              variants={{
                visible: {
                  transition: { staggerChildren: 0.08, delayChildren: line1.length * 0.08 + 1 }
                }
              }}
            >
              {splitChars(line2Prefix).map((char, i) => (
                <motion.span key={i} variants={charVariants} style={{ color: '#E6E7E9' }}>
                  {char}
                </motion.span>
              ))}
              <span className="font-bold" style={{ color: '#15F0DB' }}>
                {brand}
              </span>
            </motion.p>

            {/* Linha em branco */}
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>

            {/* Linha 3 */}
            <motion.p
              className="text-xl md:text-3xl lg:text-4xl font-light"
              initial="hidden"
              animate="visible"
              exit="exit"
              variants={{
                visible: {
                  transition: { staggerChildren: 0.03, delayChildren: (line1.length * 0.08) + (line2Prefix.length * 0.08) + 2 }
                }
              }}
            >
              {splitChars(line3).map((char, i) => (
                <motion.span key={i} variants={charVariants} style={{ color: '#E6E7E9' }}>
                  {char}
                </motion.span>
              ))}
            </motion.p>

          </div>
        </motion.div>
      )}
    </AnimatePresence>
  );
};

export default Preloader;