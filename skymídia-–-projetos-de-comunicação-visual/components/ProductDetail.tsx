import React, { useState } from 'react';
import { Product } from '../types';

interface ProductDetailProps {
  product: Product;
  onBack: () => void;
  onNavigateToContact: () => void;
}

const ProductDetail: React.FC<ProductDetailProps> = ({ product, onBack, onNavigateToContact }) => {
  const [selectedSize, setSelectedSize] = useState<string | null>(null);
  
  const sizes = ['S', 'M', 'L'];
  const showSizes = product.category === 'Wearable';

  return (
    <div className="pt-24 min-h-screen animate-fade-in-up">
      <div className="max-w-[1800px] mx-auto px-6 md:px-12 pb-24">
        
        <button 
          onClick={onBack}
          className="group flex items-center gap-2 text-xs font-medium uppercase tracking-widest text-brand-text/60 hover:text-brand-hover transition-colors mb-8"
        >
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-4 h-4 group-hover:-translate-x-1 transition-transform">
            <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 19.5L8.25 12l7.5-7.5" />
          </svg>
          Voltar para Produtos
        </button>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 lg:gap-24">
          
          <div className="flex flex-col gap-4">
            <div className="w-full aspect-[4/5] bg-brand-text/5 overflow-hidden">
              <img 
                src={product.imageUrl} 
                alt={product.name} 
                className="w-full h-full object-cover animate-fade-in-up"
              />
            </div>
          </div>

          <div className="flex flex-col justify-center max-w-xl bg-brand-dark/5 p-8 md:p-12 rounded-sm border border-brand-hover/10 shadow-sm backdrop-blur-sm">
             <span className="text-sm font-medium text-brand-text/60 uppercase tracking-widest mb-2">{product.category}</span>
             <h1 className="text-4xl md:text-5xl font-serif text-brand-text mb-6">{product.name}</h1>
             
             <p className="text-brand-text/80 leading-relaxed font-light text-lg mb-8">
               {product.longDescription || product.description}
             </p>

             {showSizes && (
                <div className="mb-8">
                  <span className="block text-xs font-bold uppercase tracking-widest text-brand-text mb-4">Select Size</span>
                  <div className="flex gap-4">
                    {sizes.map(size => (
                      <button 
                        key={size}
                        onClick={() => setSelectedSize(size)}
                        className={`w-12 h-12 flex items-center justify-center border transition-all duration-300 ${
                          selectedSize === size 
                            ? 'border-brand-hover bg-brand-hover text-brand-bg' 
                            : 'border-brand-text/20 text-brand-text/60 hover:border-brand-hover'
                        }`}
                      >
                        {size}
                      </button>
                    ))}
                  </div>
                </div>
             )}

             <div className="flex flex-col gap-4">
               <button 
                 onClick={onNavigateToContact}
                 className="w-full py-5 bg-brand-hover text-brand-bg uppercase tracking-widest text-sm font-bold hover:bg-brand-hover/80 transition-colors shadow-lg shadow-brand-hover/10"
               >
                 FALE CONOSCO !
               </button>
               <ul className="mt-8 space-y-2 text-sm text-brand-text/70">
                 {product.features.map((feature, idx) => (
                   <li key={idx} className="flex items-center gap-3">
                     <span className="w-1 h-1 bg-brand-hover rounded-full"></span>
                     {feature}
                   </li>
                 ))}
               </ul>
             </div>
          </div>

        </div>
      </div>
    </div>
  );
};

export default ProductDetail;
