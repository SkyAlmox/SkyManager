import React, { useState, useMemo } from 'react';
import { Product } from '../types';
import ProductCard from './ProductCard';

const PRODUCTS: Product[] = [
  {
    id: 1,
    name: "Carregador Turbo",
    category: "Carregadores de Celular",
    description: "Carregadores rápidos e seguros para diversos modelos de smartphones.",
    imageUrl: "/images/carregador.jpg",
    price: 99.90
  },
  {
    id: 2,
    name: "Painel Publicitário",
    category: "Comunicação Visual",
    description: "Soluções criativas em comunicação visual para destacar sua marca.",
    imageUrl: "/images/comunicacao-visual.jpg",
    price: 1200.00
  },
  {
    id: 3,
    name: "Estrutura Outdoor",
    category: "Estruturas para Mídia OOH",
    description: "Estruturas resistentes e modernas para mídia externa e outdoors.",
    imageUrl: "/images/estrutura-ooh.jpg",
    price: 3500.00
  },
  {
    id: 4,
    name: "Móveis Sob Medida",
    category: "Marcenaria",
    description: "Marcenaria personalizada para ambientes corporativos e residenciais.",
    imageUrl: "/images/marcenaria.jpg",
    price: 2500.00
  },
  {
    id: 5,
    name: "Projeto Arquitetônico Especial",
    category: "Projetos Especiais de Arquitetura",
    description: "Projetos exclusivos que unem inovação e funcionalidade.",
    imageUrl: "/images/arquitetura.jpg",
    price: 8000.00
  },
  {
    id: 6,
    name: "Totem Interativo",
    category: "Totens Digitais e Interativos",
    description: "Totens digitais para interação e experiência diferenciada.",
    imageUrl: "/images/totem.jpg",
    price: 4500.00
  }
];

const categories = [
  "Carregadores de Celular",
  "Comunicação Visual",
  "Estruturas para Mídia OOH",
  "Marcenaria",
  "Projetos Especiais de Arquitetura",
  "Totens Digitais e Interativos"
];

interface ProductGridProps {
  onProductClick: (product: Product) => void;
}

const ProductGrid: React.FC<ProductGridProps> = ({ onProductClick }) => {
  const [activeCategory, setActiveCategory] = useState(categories[0]);

  const filteredProducts = useMemo(() => {
    return PRODUCTS.filter(p => p.category === activeCategory);
  }, [activeCategory]);

  return (
    <section id="products" className="py-32 px-6 md:px-12 bg-[#F5F2EB]">
      <div className="max-w-[1800px] mx-auto">
        <div className="flex flex-col items-center text-center mb-24 space-y-8">
          <h2 className="text-4xl md:text-6xl font-serif text-[#2C2A26]">Produtos</h2>
          <div className="flex flex-wrap justify-center gap-8 pt-4 border-t border-[#D6D1C7]/50 w-full max-w-2xl">
            {categories.map(cat => (
              <button
                key={cat}
                onClick={() => setActiveCategory(cat)}
                className={`text-sm uppercase tracking-widest pb-1 border-b transition-all duration-300 ${
                  activeCategory === cat
                    ? 'border-[#2C2A26] text-[#2C2A26]'
                    : 'border-transparent text-[#A8A29E] hover:text-[#2C2A26]'
                }`}
              >
                {cat}
              </button>
            ))}
          </div>
        </div>
        <div className="flex flex-row flex-wrap justify-center gap-8">
          {filteredProducts.map(product => (
            <ProductCard key={product.id} product={product} onClick={onProductClick} />
          ))}
        </div>
      </div>
    </section>
  );
};

export default ProductGrid;