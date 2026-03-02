import React, { useState, useEffect, useCallback } from 'react';
import { createPortal } from 'react-dom';
import { motion, AnimatePresence } from 'framer-motion';
import { X, ZoomIn, ZoomOut, Maximize2, ChevronLeft, ChevronRight } from 'lucide-react';

export interface GalleryCategory {
  title: string;
  subtitle: string;
  images: string[];
}

interface ProductGalleryProps {
  categories?: GalleryCategory[];
}

const DEFAULT_CATEGORIES: GalleryCategory[] = [
  {
    title: 'PLACAS E LETREIROS',
    subtitle: '2D e 3D com e sem iluminação',
    images: Array.from({ length: 15 }, (_, i) => `https://picsum.photos/seed/placa-${i}/800/600`)
  },
  {
    title: 'PAINÉIS IMPRESSOS',
    subtitle: 'Adesivos Acoplados | Lonas | Tecidos',
    images: Array.from({ length: 15 }, (_, i) => `https://picsum.photos/seed/painel-${i}/800/600`)
  },
  {
    title: 'TOTENS',
    subtitle: 'Sinalização | Institucionais | Promocionais',
    images: Array.from({ length: 15 }, (_, i) => `https://picsum.photos/seed/totem-${i}/800/600`)
  },
  {
    title: 'PLOTAGENS',
    subtitle: 'Plotter de Recorte | Grandes Formatos',
    images: Array.from({ length: 15 }, (_, i) => `https://picsum.photos/seed/plotagem-${i}/800/600`)
  },
  {
    title: 'NEON',
    subtitle: 'Tradicional | Led',
    images: Array.from({ length: 15 }, (_, i) => `https://picsum.photos/seed/neon-${i}/800/600`)
  },
  {
    title: 'SINALIZAÇÃO',
    subtitle: 'Externa | Interna',
    images: Array.from({ length: 15 }, (_, i) => `https://picsum.photos/seed/sinal-${i}/800/600`)
  },
  {
    title: 'PROJETOS ESPECIAIS',
    subtitle: '100% Personalizados',
    images: Array.from({ length: 15 }, (_, i) => `https://picsum.photos/seed/especial-${i}/800/600`)
  }
];

const ProductGallery: React.FC<ProductGalleryProps> = ({ categories = DEFAULT_CATEGORIES }) => {
  const [selectedIndex, setSelectedIndex] = useState<{ catIdx: number; imgIdx: number } | null>(null);
  const [zoom, setZoom] = useState(1);

  const selectedImage = selectedIndex !== null 
    ? categories[selectedIndex.catIdx].images[selectedIndex.imgIdx] 
    : null;

  const handleImageClick = (catIdx: number, imgIdx: number) => {
    setSelectedIndex({ catIdx, imgIdx });
    setZoom(1);
    document.body.style.overflow = 'hidden';
  };

  const closeModal = useCallback(() => {
    setSelectedIndex(null);
    setZoom(1);
    document.body.style.overflow = 'unset';
  }, []);

  const handleNext = useCallback((e?: React.MouseEvent) => {
    e?.stopPropagation();
    if (selectedIndex === null) return;
    
    const { catIdx, imgIdx } = selectedIndex;
    const currentCat = categories[catIdx];
    
    if (imgIdx < currentCat.images.length - 1) {
      setSelectedIndex({ catIdx, imgIdx: imgIdx + 1 });
    } else if (catIdx < categories.length - 1) {
      setSelectedIndex({ catIdx: catIdx + 1, imgIdx: 0 });
    } else {
      setSelectedIndex({ catIdx: 0, imgIdx: 0 });
    }
    setZoom(1);
  }, [selectedIndex, categories]);

  const handlePrev = useCallback((e?: React.MouseEvent) => {
    e?.stopPropagation();
    if (selectedIndex === null) return;
    
    const { catIdx, imgIdx } = selectedIndex;
    
    if (imgIdx > 0) {
      setSelectedIndex({ catIdx, imgIdx: imgIdx - 1 });
    } else if (catIdx > 0) {
      const prevCatIdx = catIdx - 1;
      setSelectedIndex({ catIdx: prevCatIdx, imgIdx: categories[prevCatIdx].images.length - 1 });
    } else {
      const lastCatIdx = categories.length - 1;
      setSelectedIndex({ catIdx: lastCatIdx, imgIdx: categories[lastCatIdx].images.length - 1 });
    }
    setZoom(1);
  }, [selectedIndex, categories]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (selectedIndex === null) return;
      if (e.key === 'ArrowRight') handleNext();
      if (e.key === 'ArrowLeft') handlePrev();
      if (e.key === 'Escape') closeModal();
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [selectedIndex, handleNext, handlePrev, closeModal]);

  const toggleZoom = () => {
    setZoom(prev => (prev === 1 ? 2 : 1));
  };

  const [touchStart, setTouchStart] = useState<number | null>(null);

  const handleTouchStart = (e: React.TouchEvent) => {
    setTouchStart(e.targetTouches[0].clientX);
  };

  const handleTouchMove = (e: React.TouchEvent) => {
    if (touchStart === null) return;
    const touchEnd = e.targetTouches[0].clientX;
    const diff = touchStart - touchEnd;

    if (diff > 50) {
      handleNext();
      setTouchStart(null);
    } else if (diff < -50) {
      handlePrev();
      setTouchStart(null);
    }
  };

  const handleTouchEnd = () => {
    setTouchStart(null);
  };

  return (
    <div className="mt-20 space-y-24">
      {categories.map((cat, idx) => (
        <div key={idx} className="space-y-8">
          <div className="border-l-4 border-brand-hover pl-6">
            <h3 className="text-2xl font-serif text-brand-text/90 font-bold tracking-tight">{cat.title}</h3>
            <p className="text-sm text-brand-text/60 uppercase tracking-widest mt-1">{cat.subtitle}</p>
          </div>
          
          <div className="grid grid-cols-3 sm:grid-cols-5 gap-2 md:gap-4">
            {cat.images.map((img, imgIdx) => (
              <motion.div 
                key={imgIdx}
                whileHover={{ scale: 1.05 }}
                whileTap={{ scale: 0.95 }}
                className="aspect-square overflow-hidden bg-brand-dark/5 cursor-pointer group relative"
                onClick={() => handleImageClick(idx, imgIdx)}
              >
                <img 
                  src={img} 
                  alt={`${cat.title} ${imgIdx + 1}`}
                  className="w-full h-full object-cover transition-transform duration-500 group-hover:scale-110"
                  referrerPolicy="no-referrer"
                />
                <div className="absolute inset-0 bg-black/20 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                  <Maximize2 className="text-white w-5 h-5" />
                </div>
              </motion.div>
            ))}
          </div>
        </div>
      ))}

      {typeof document !== 'undefined' && createPortal(
        <AnimatePresence>
          {selectedImage && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 z-[100] flex items-center justify-center bg-black/90 p-2 md:p-10"
              onClick={closeModal}
              onTouchStart={handleTouchStart}
              onTouchMove={handleTouchMove}
              onTouchEnd={handleTouchEnd}
            >
              <motion.div 
                initial={{ scale: 0.9, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.9, opacity: 0 }}
                className="relative w-full max-w-[95vw] md:max-w-[80vw] max-h-[90vh] overflow-hidden bg-brand-bg flex items-center justify-center"
                onClick={(e) => e.stopPropagation()}
              >
                <button 
                  className="absolute top-4 right-4 z-[110] bg-black/50 text-white p-2 rounded-full hover:bg-black/70 transition-colors"
                  onClick={closeModal}
                >
                  <X size={24} />
                </button>

                <button 
                  className="absolute left-2 md:left-4 top-1/2 -translate-y-1/2 z-[110] bg-black/50 text-white p-2 md:p-3 rounded-full hover:bg-black/70 transition-colors flex items-center justify-center"
                  onClick={handlePrev}
                  aria-label="Previous image"
                >
                  <ChevronLeft size={24} className="md:w-8 md:h-8" />
                </button>

                <button 
                  className="absolute right-2 md:right-4 top-1/2 -translate-y-1/2 z-[110] bg-black/50 text-white p-2 md:p-3 rounded-full hover:bg-black/70 transition-colors flex items-center justify-center"
                  onClick={handleNext}
                  aria-label="Next image"
                >
                  <ChevronRight size={24} className="md:w-8 md:h-8" />
                </button>

                <div className="absolute bottom-4 left-1/2 -translate-x-1/2 z-[110] flex gap-4 bg-black/50 p-2 rounded-full backdrop-blur-sm">
                  <button 
                    className="text-white p-1 hover:text-brand-hover transition-colors"
                    onClick={() => setZoom(prev => Math.max(1, prev - 0.5))}
                  >
                    <ZoomOut size={20} />
                  </button>
                  <span className="text-white text-xs font-mono flex items-center px-2">
                    {Math.round(zoom * 100)}%
                  </span>
                  <button 
                    className="text-white p-1 hover:text-brand-hover transition-colors"
                    onClick={() => setZoom(prev => Math.min(3, prev + 0.5))}
                  >
                    <ZoomIn size={20} />
                  </button>
                </div>

                <div className="w-full h-full overflow-auto flex items-center justify-center p-4">
                  <motion.img 
                    key={selectedImage}
                    src={selectedImage} 
                    alt="Gallery Preview"
                    initial={{ opacity: 0, x: 20 }}
                    animate={{ opacity: 1, x: 0, scale: zoom }}
                    transition={{ 
                      scale: { type: 'spring', stiffness: 300, damping: 30 },
                      opacity: { duration: 0.2 }
                    }}
                    className="max-w-full max-h-full object-contain cursor-zoom-in"
                    onClick={toggleZoom}
                    referrerPolicy="no-referrer"
                  />
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>,
        document.body
      )}
    </div>
  );
};

export default ProductGallery;
