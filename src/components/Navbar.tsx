import React, { useState, useEffect } from 'react';
import { Instagram, MessageCircle } from 'lucide-react';
import { BRAND_NAME, PRIMARY_COLOR, ACCENT_COLOR } from '../constants';

interface NavbarProps {
  onNavClick: (e: React.MouseEvent<HTMLAnchorElement>, targetId: string) => void;
  cartCount: number;
  onOpenCart: () => void;
}

const Navbar: React.FC<NavbarProps> = ({ onNavClick, cartCount, onOpenCart }) => {
  const [scrolled, setScrolled] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [submenuOpen, setSubmenuOpen] = useState(false);

  useEffect(() => {
    const handleScroll = () => {
      setScrolled(window.scrollY > 50);
    };
    window.addEventListener('scroll', handleScroll);
    return () => window.removeEventListener('scroll', handleScroll);
  }, []);

  const handleLinkClick = (e: React.MouseEvent<HTMLAnchorElement>, targetId: string) => {
    setMobileMenuOpen(false);
    onNavClick(e, targetId);
  };

  const textColorClass = (scrolled || mobileMenuOpen) ? 'text-[#2C2A26]' : 'text-[#F5F2EB]';

  return (
    <>
      <nav
        className={`fixed top-0 left-0 right-0 z-50 transition-all duration-700 ease-in-out ${
          scrolled || mobileMenuOpen
            ? `bg-[${PRIMARY_COLOR}]/90 backdrop-blur-md py-4 shadow-sm`
            : 'bg-transparent py-8'
        }`}
      >
        <div className="max-w-[1800px] mx-auto px-8 flex items-center justify-between">
          {/* Logo */}
          <a
            href="#"
            onClick={(e) => {
              e.preventDefault();
              window.scrollTo({ top: 0, behavior: 'smooth' });
              onNavClick(e, '');
            }}
            className="z-50 relative"
          >
            <img
              src="/logo-branca.png"
              alt="Logo da Empresa"
              className="h-8 w-auto sm:h-6 md:h-8 lg:h-10 transition-all duration-500"
            />
          </a>

          {/* Links - Desktop */}
          <div className={`hidden md:flex items-center gap-12 text-sm font-medium tracking-widest uppercase transition-colors duration-500 ${textColorClass}`}>

            {/* Produtos com submenu */}
            <div className="relative group">
              <a
                href="#products"
                onClick={(e) => handleLinkClick(e, 'products')}
                className={`transition-colors duration-300 hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4`}
              >
                Produtos
              </a>
              <div className="absolute left-0 mt-2 w-72 bg-[${PRIMARY_COLOR}]/90 backdrop-blur-md rounded-md shadow-lg opacity-0 group-hover:opacity-100 transition-opacity duration-300 z-50">
                {[
                  "Carregadores de Celular",
                  "Comunicação Visual",
                  "Estruturas para Mídia OOH",
                  "Projetos Especiais de Arquitetura",
                  "Totens Digitais e Interativos"
                ].map((item) => (
                  <a
                    key={item}
                    href="#"
                    className="block px-4 py-2 text-sm text-[#F5F2EB] hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4 transition-colors"
                  >
                    {item}
                  </a>
                ))}
              </div>
            </div>

            <a href="#about" onClick={(e) => handleLinkClick(e, 'about')}
               className={`transition-colors duration-300 hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4`}>
              About
            </a>
            <a href="#journal" onClick={(e) => handleLinkClick(e, 'journal')}
               className={`transition-colors duration-300 hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4`}>
              Journal
            </a>
          </div>

          {/* Ícones + Menu Mobile Toggle */}
          <div className={`flex items-center gap-5 z-50 relative transition-colors duration-500 ${textColorClass}`}>
            <a
              href="https://instagram.com"
              target="_blank"
              rel="noopener noreferrer"
              className={`transition-colors duration-300 hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4`}
              aria-label="Instagram"
            >
              <Instagram size={20} strokeWidth={1.5} />
            </a>
            <a
              href="https://wa.me/yournumber"
              target="_blank"
              rel="noopener noreferrer"
              className={`transition-colors duration-300 hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4`}
              aria-label="WhatsApp"
            >
              <MessageCircle size={20} strokeWidth={1.5} />
            </a>

            {/* Botão do Menu Mobile */}
            <button
              className={`block md:hidden focus:outline-none transition-colors duration-500 ${textColorClass}`}
              onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
            >
              {mobileMenuOpen ? (
                <svg xmlns="http://www.w3.org/2000/svg" fill="none"
                     viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor"
                     className="w-6 h-6">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                </svg>
              ) : (
                <svg xmlns="http://www.w3.org/2000/svg" fill="none"
                     viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor"
                     className="w-6 h-6">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5" />
                </svg>
              )}
            </button>
          </div>
        </div>
      </nav>

      {/* Menu Mobile Overlay */}
      <div className={`fixed inset-0 bg-[${PRIMARY_COLOR}] z-40 flex flex-col justify-center items-center transition-all duration-500 ease-in-out ${
        mobileMenuOpen ? 'opacity-100 translate-y-0 pointer-events-auto' : 'opacity-0 -translate-y-10 pointer-events-none'
      }`}>
        <div className="flex flex-col items-center space-y-8 text-xl font-serif font-medium text-[#F5F2EB]">
          <a href="#products" onClick={(e) => handleLinkClick(e, 'products')}
             className={`transition-colors duration-300 hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4`}>
            Produtos
          </a>
          <a href="#about" onClick={(e) => handleLinkClick(e, 'about')}
             className={`transition-colors duration-300 hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4`}>
            About
          </a>
          <a href="#journal" onClick={(e) => handleLinkClick(e, 'journal')}
             className={`transition-colors duration-300 hover:text-[${ACCENT_COLOR}] hover:underline underline-offset-4`}>
            Journal
          </a>
        </div>
      </div>
    </>
  );
};

export default Navbar;
