
/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/


import React from 'react';
import { Product } from '../types';

interface CheckoutProps {
  items: Product[];
  onBack: () => void;
}

const Checkout: React.FC<CheckoutProps> = ({ items, onBack }) => {
  const subtotal = items.reduce((sum, item) => sum + item.price, 0);
  const shipping = 0; // Free shipping
  const total = subtotal + shipping;

  return (
    <div className="min-h-screen pt-24 pb-24 px-6 bg-brand-bg animate-fade-in-up">
      <div className="max-w-6xl mx-auto">
        <button 
          onClick={onBack}
          className="group flex items-center gap-2 text-xs font-medium uppercase tracking-widest text-brand-text/60 hover:text-brand-hover transition-colors mb-12"
        >
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-4 h-4 group-hover:-translate-x-1 transition-transform">
            <path strokeLinecap="round" strokeLinejoin="round" d="M15.75 19.5L8.25 12l7.5-7.5" />
          </svg>
          Voltar
        </button>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-16 lg:gap-24">
          
          {/* Left Column: Form */}
          <div>
            <h1 className="text-3xl font-serif text-brand-text mb-4">Orçamento</h1>
            <p className="text-sm text-brand-text/60 mb-12">Entre em contato para um orçamento personalizado.</p>
            
            <div className="space-y-12">
              {/* Section 1: Contact */}
              <div>
                <h2 className="text-xl font-serif text-brand-text mb-6">Informações de Contato</h2>
                <div className="space-y-4">
                   <input type="email" placeholder="Endereço de e-mail" className="w-full bg-transparent border-b border-brand-hover/20 py-3 text-brand-text placeholder-brand-text/40 outline-none focus:border-brand-hover transition-colors" />
                </div>
              </div>

              {/* Section 2: Details */}
              <div>
                <h2 className="text-xl font-serif text-brand-text mb-6">Detalhes do Projeto</h2>
                <div className="space-y-4">
                   <div className="grid grid-cols-2 gap-4">
                      <input type="text" placeholder="Nome" className="w-full bg-transparent border-b border-brand-hover/20 py-3 text-brand-text placeholder-brand-text/40 outline-none focus:border-brand-hover transition-colors" />
                      <input type="text" placeholder="Empresa" className="w-full bg-transparent border-b border-brand-hover/20 py-3 text-brand-text placeholder-brand-text/40 outline-none focus:border-brand-hover transition-colors" />
                   </div>
                   <textarea placeholder="Descreva seu projeto" className="w-full bg-transparent border-b border-brand-hover/20 py-3 text-brand-text placeholder-brand-text/40 outline-none focus:border-brand-hover transition-colors h-32 resize-none"></textarea>
                </div>
              </div>

              <div>
                <button 
                    className="w-full py-5 bg-brand-hover text-brand-bg uppercase tracking-widest text-sm font-medium hover:bg-brand-hover/80 transition-colors"
                >
                    Enviar Solicitação
                </button>
              </div>
            </div>
          </div>

          {/* Right Column: Summary */}
          <div className="lg:pl-12 lg:border-l border-brand-hover/20">
            <h2 className="text-xl font-serif text-brand-text mb-8">Itens Selecionados</h2>
            
            <div className="space-y-6 mb-8">
               {items.map((item, idx) => (
                 <div key={idx} className="flex gap-4">
                    <div className="w-16 h-16 bg-brand-text/5 relative">
                       <img src={item.imageUrl} alt={item.name} className="w-full h-full object-cover" />
                    </div>
                    <div className="flex-1">
                       <h3 className="font-serif text-brand-text text-base">{item.name}</h3>
                       <p className="text-xs text-brand-text/60">{item.category}</p>
                    </div>
                 </div>
               ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Checkout;