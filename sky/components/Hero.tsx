import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';

const images = [
  "/index_01.jpg",
  "/index_02.jpg",
  "/index_03.jpg",
  "/index_04.jpg",
  "/index_05.jpg",
  "/index_06.jpg",
  "/index_07.jpg",
  "/index_08.jpg",
  "/index_09.jpg",
  "/index_10.jpg",
  "/index_11.jpg"
];

interface HeroProps {
  onNavClick?: (e: React.MouseEvent<HTMLAnchorElement>, targetId: string) => void;
}

const Hero: React.FC<HeroProps> = ({ onNavClick }) => {
  const [currentIndex, setCurrentIndex] = useState(0);

  useEffect(() => {
    const timer = setInterval(() => {
      setCurrentIndex((prevIndex) => (prevIndex + 1) % images.length);
    }, 7000);

    return () => clearInterval(timer);
  }, []);

  const handleProductsClick = (e: React.MouseEvent<HTMLAnchorElement>) => {
    if (onNavClick) {
      onNavClick(e, 'journal');
    } else {
      e.preventDefault();
      const element = document.getElementById('journal');
      if (element) {
        element.scrollIntoView({ behavior: 'smooth' });
      }
    }
  };

  return (
    <section className="relative w-full h-screen min-h-[800px] overflow-hidden bg-brand-bg">

      {/* Carousel Background */}
      <div className="absolute inset-0 w-full h-full">
        <AnimatePresence>
          <motion.div
            key={currentIndex}
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 2.5, ease: "easeInOut" }}
            className="absolute inset-0 w-full h-full"
          >
            <motion.img
                initial={{ scale: 1.1 }}
                animate={{ scale: 1 }}
                transition={{ duration: 8, ease: "linear" }}
                src={images[currentIndex]}
                alt={`Slide ${currentIndex + 1}`}
                className="w-full h-full object-cover grayscale contrast-[0.7] brightness-[0.95]"
            />
          </motion.div>
        </AnimatePresence>

        {/* Dark Overlay for Richness */}
        <div className="absolute inset-0 bg-brand-bg/40 mix-blend-multiply z-[1]"></div>
        {/* Subtle Depth Overlay */}
        <div className="absolute inset-0 bg-black/20 z-[2]"></div>
      </div>

      {/* Content */}
      <div className="relative z-10 h-full flex flex-col justify-center items-start text-left md:items-center md:text-center px-6">
        <div className="animate-fade-in-up w-full md:w-auto">
          <h1 className="text-6xl md:text-8xl lg:text-9xl font-serif font-normal text-brand-text tracking-tight mb-8 drop-shadow-sm">
            Somos a <span className="italic text-brand-hover">SKYMÍDIA</span>
          </h1>
          <div className="max-w-3xl mx-0 md:mx-auto text-lg md:text-xl text-brand-text/90 font-light leading-relaxed mb-12 text-shadow-sm space-y-6">
            <p>
              Transformamos matérias-primas comuns em projetos de Comunicação Visual, de modo que cada cliente obtenha uma solução sob medida para cada uma das suas necessidades.
            </p>
            <p>
              Hoje temos 6 divisões principais de produção, e vamos te mostrar cada uma delas com mais detalhes nas <a href="#journal" onClick={handleProductsClick} className="text-brand-hover font-medium underline underline-offset-4 hover:opacity-80 transition-opacity">próximas páginas</a>.
            </p>
          </div>
        </div>
      </div>

      {/* Scroll Indicator */}
      <div className="absolute bottom-12 left-1/2 -translate-x-1/2 animate-bounce text-brand-text/50 z-10">
        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1} d="M19 14l-7 7m0 0l-7-7m7 7V3" />
        </svg>
      </div>
    </section>
  );
};

export default Hero;
