import React, { useState } from 'react';
import { createPortal } from 'react-dom';
import { motion, AnimatePresence } from 'framer-motion';
import { X, ZoomIn, ZoomOut, Maximize2 } from 'lucide-react';

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
  const [currentCategory, setCurrentCategory] = useState<GalleryCategory | null>(null);
  const [currentIndex, setCurrentIndex] = useState<number | null>(null);
  const [zoom, setZoom] = useState(1);

  const handleImageClick = (cat: GalleryCategory, idx: number) => {
    setCurrentCategory(cat);
    setCurrentIndex(idx);
    setZoom(1);
    document.body.style.overflow = 'hidden';
  };

  const closeModal = () => {
    setCurrentCategory(null);
    setCurrentIndex(null);
    setZoom(1);
    document.body.style.overflow = 'unset';
  };

  const toggleZoom = () => {
    setZoom(prev => (prev === 1 ? 2 : 1));
  };

  const showNext = () => {
    if (currentCategory && currentIndex !== null) {
      setCurrentIndex((currentIndex + 1) % currentCategory.images.length);
    }
  };

  const showPrev = () => {
    if (currentCategory && currentIndex !== null) {
      setCurrentIndex((currentIndex - 1 + currentCategory.images.length) % currentCategory.images.length);
    }
  };

  return (
    <div className="mt-20 space-y-24">
      {categories.map((cat, idx) => (
        <div key={idx} className="space-y-8">
          <div className="border-l-4 border-brand-hover pl-6">
            <h3 className="text-2xl font-serif text-brand-dark/70 font-bold tracking-tight">{cat.title}</h3>
            <p className="text-sm text-brand-dark/70 uppercase tracking-widest mt-1">{cat.subtitle}</p>
          </div>

          <div className="grid grid-cols-3 sm:grid-cols-5 gap-2 md:gap-4">
            {cat.images.map((img, imgIdx) => (
              <motion.div
                key={imgIdx}
                whileHover={{ scale: 1.05 }}
                whileTap={{ scale: 0.95 }}
                className="aspect-square overflow-hidden bg-brand-dark/5 cursor-pointer group relative"
                onClick={() => handleImageClick(cat, imgIdx)}
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

      {/* Modal com setas de navegação */}
      {typeof document !== 'undefined' && createPortal(
        <AnimatePresence>
          {currentCategory && currentIndex !== null && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 z-[100] flex items-center justify-center bg-black/90 p-4 md:p-10"
              onClick={closeModal}
            >
              <motion.div
                initial={{ scale: 0.9, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.9, opacity: 0 }}
                className="relative w-full max-w-[66.66vw] max-h-[90vh] overflow-hidden bg-brand-bg flex items-center justify-center"
                onClick={(e) => e.stopPropagation()}
              >
                {/* Botão fechar */}
                <button
                  className="absolute top-4 right-4 z-[110] bg-black/50 text-white p-2 rounded-full hover:bg-black/70 transition-colors"
                  onClick={closeModal}
                >
                  <X size={24} />
                </button>

                {/* Setas de navegação */}
                <button
                  className="absolute left-4 top-1/2 -translate-y-1/2 text-white p-2 rounded-full bg-black/50 hover:bg-black/70"
                  onClick={showPrev}
                >
                  ←
                </button>
                <button
                  className="absolute right-4 top-1/2 -translate-y-1/2 text-white p-2 rounded-full bg-black/50 hover:bg-black/70"
                  onClick={showNext}
                >
                  →
                </button>

                {/* Controles de zoom */}
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

                {/* Imagem atual */}
                <div className="w-full h-full overflow-auto flex items-center justify-center p-4">
                  <motion.img
                    src={currentCategory.images[currentIndex]}
                    alt="Gallery Preview"
                    animate={{ scale: zoom }}
                    transition={{ type: 'spring', stiffness: 300, damping: 30 }}
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