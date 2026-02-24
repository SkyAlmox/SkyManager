import React, { useState, useEffect } from 'react';
import { BRAND_NAME, JOURNAL_ARTICLES } from '../constants';
import { Instagram, MessageCircle, ChevronDown } from 'lucide-react';
import { JournalArticle } from '../types';

interface NavbarProps {
  onNavClick: (e: React.MouseEvent<HTMLAnchorElement>, targetId: string) => void;
  onArticleClick: (article: JournalArticle) => void;
}

const Navbar: React.FC<NavbarProps> = ({ onNavClick, onArticleClick }) => {
  const [scrolled, setScrolled] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [dropdownOpen, setDropdownOpen] = useState(false);

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

  // Determine text color based on state
  const textColorClass = (scrolled || mobileMenuOpen) ? 'text-brand-text' : 'text-brand-text';

  return (
    <>
      <nav
        className={`fixed top-0 left-0 right-0 z-50 transition-all duration-700 ease-in-out ${
          scrolled || mobileMenuOpen ? 'bg-brand-bg/90 backdrop-blur-md py-4 shadow-sm' : 'bg-transparent py-8'
        }`}
      >
        <div className="max-w-[1800px] mx-auto px-8 flex items-center justify-between">
          {/* Logo */}
          <a
            href="#"
            onClick={(e) => {
                e.preventDefault();
                window.scrollTo({ top: 0, behavior: 'smooth' });
                onNavClick(e, ''); // Pass empty string to just reset to home
            }}
            className="z-50 relative block"
          >
            <img
              src="/logo.png"
              alt={BRAND_NAME}
              className="h-12 md:h-16 w-auto transition-all duration-500"
            />
          </a>

          {/* Center Links - Desktop */}
          <div className={`hidden md:flex items-center gap-10 text-sm font-medium tracking-widest uppercase transition-colors duration-500 ${textColorClass}`}>
            <a
              href="#"
              onClick={(e) => {
                e.preventDefault();
                window.scrollTo({ top: 0, behavior: 'smooth' });
                onNavClick(e, '');
              }}
              className="hover:text-brand-hover transition-colors whitespace-nowrap"
            >
              Home
            </a>
            {/* <a href="#about" onClick={(e) => handleLinkClick(e, 'about')} className="hover:text-brand-hover transition-colors whitespace-nowrap">Identidade Visual</a> */}

            {/* Dropdown Produtos */}
            <div
              className="relative group py-2"
              onMouseEnter={() => setDropdownOpen(true)}
              onMouseLeave={() => setDropdownOpen(false)}
            >
              <a
                href="#journal"
                onClick={(e) => handleLinkClick(e, 'journal')}
                className="hover:text-brand-hover transition-colors flex items-center gap-1"
              >
                Produtos <ChevronDown size={14} className={`transition-transform duration-300 ${dropdownOpen ? 'rotate-180' : ''}`} />
              </a>

              {/* Dropdown Content */}
              <div className={`absolute top-full left-1/2 -translate-x-1/2 pt-4 transition-all duration-300 ${dropdownOpen ? 'opacity-100 translate-y-0 pointer-events-auto' : 'opacity-0 -translate-y-2 pointer-events-none'}`}>
                <div className="bg-brand-bg border border-brand-hover shadow-xl py-6 min-w-[280px] rounded-sm">
                  {JOURNAL_ARTICLES
                  .filter(a => a.id !== 7)
                  .sort((a, b) => {
                    const order = [3, 2, 5, 4, 6, 1];
                    return order.indexOf(a.id) - order.indexOf(b.id);
                  })
                  .map((article) => (
                    <a
                      key={article.id}
                      href="#journal"
                      onClick={(e) => {
                        e.preventDefault();
                        setDropdownOpen(false);
                        onArticleClick(article);
                      }}
                      className="flex items-center gap-3 px-8 py-3 text-[11px] tracking-[0.15em] text-brand-text/80 hover:text-brand-hover hover:bg-white/5 transition-colors"
                    >
                      {article.icon && (
                        <img
                          src={article.icon}
                          alt=""
                          className="h-6 w-auto object-contain"
                          referrerPolicy="no-referrer"
                        />
                      )}
                      <span>{article.title}</span>
                    </a>
                  ))}
                </div>
              </div>
            </div>

            <a
              href="#footer"
              onClick={(e) => {
                e.preventDefault();
                const contactArticle = JOURNAL_ARTICLES.find(a => a.id === 7);
                if (contactArticle) onArticleClick(contactArticle);
              }}
              className="hover:text-brand-hover transition-colors whitespace-nowrap"
            >
              Vamos Conversar ?
            </a>
          </div>

          {/* Right Actions */}
          <div className={`flex items-center gap-6 z-50 relative transition-colors duration-500 ${textColorClass}`}>
            <div className="flex items-center gap-5">
              <a
                href="https://instagram.com"
                target="_blank"
                rel="noopener noreferrer"
                className="hover:opacity-60 transition-opacity"
                aria-label="Instagram"
              >
                <Instagram size={20} strokeWidth={1.5} />
              </a>
              <a
                href="https://wa.me/#"
                target="_blank"
                rel="noopener noreferrer"
                className="hover:opacity-60 transition-opacity"
                aria-label="WhatsApp"
              >
                <MessageCircle size={20} strokeWidth={1.5} />
              </a>
            </div>

            {/* Mobile Menu Toggle */}
            <button
              className={`block md:hidden focus:outline-none transition-colors duration-500 ${textColorClass}`}
              onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
            >
               {mobileMenuOpen ? (
                 <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-6 h-6">
                   <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                 </svg>
               ) : (
                 <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-6 h-6">
                   <path strokeLinecap="round" strokeLinejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5" />
                 </svg>
               )}
            </button>
          </div>
        </div>
      </nav>

      {/* Mobile Menu Overlay */}
      <div className={`fixed inset-0 bg-brand-bg z-40 flex flex-col justify-center items-center transition-all duration-500 ease-in-out ${
          mobileMenuOpen ? 'opacity-100 translate-y-0 pointer-events-auto' : 'opacity-0 -translate-y-10 pointer-events-none'
      }`}>
          <div className="flex flex-col items-center space-y-8 text-xl font-serif font-medium text-brand-text w-full px-8">
            <a
              href="#"
              onClick={(e) => {
                e.preventDefault();
                setMobileMenuOpen(false);
                window.scrollTo({ top: 0, behavior: 'smooth' });
                onNavClick(e, '');
              }}
              className="hover:text-brand-hover transition-colors"
            >
              Home
            </a>
            <a href="#about" onClick={(e) => handleLinkClick(e, 'about')} className="hover:text-brand-hover transition-colors">Identidade Visual</a>

            <div className="flex flex-col items-center w-full">
              <button
                onClick={() => setDropdownOpen(!dropdownOpen)}
                className="hover:text-brand-hover transition-colors flex items-center gap-2 mb-4"
              >
                Produtos <ChevronDown size={20} className={`transition-transform duration-300 ${dropdownOpen ? 'rotate-180' : ''}`} />
              </button>

              {dropdownOpen && (
                <div className="flex flex-col items-center space-y-4 mb-4 animate-fade-in">
                  {JOURNAL_ARTICLES
                  .filter(a => a.id !== 7)
                  .sort((a, b) => {
                    const order = [3, 2, 5, 4, 6, 1];
                    return order.indexOf(a.id) - order.indexOf(b.id);
                  })
                  .map((article) => (
                    <a
                      key={article.id}
                      href="#journal"
                      onClick={(e) => {
                        e.preventDefault();
                        setMobileMenuOpen(false);
                        setDropdownOpen(false);
                        onArticleClick(article);
                      }}
                      className="flex flex-col items-center gap-2 text-sm font-sans uppercase tracking-widest text-brand-text/70 hover:text-brand-hover"
                    >
                      <span>{article.title}</span>
                    </a>
                  ))}
                </div>
              )}
            </div>

            <a
              href="#footer"
              onClick={(e) => {
                e.preventDefault();
                setMobileMenuOpen(false);
                const contactArticle = JOURNAL_ARTICLES.find(a => a.id === 7);
                if (contactArticle) onArticleClick(contactArticle);
              }}
              className="hover:text-brand-hover transition-colors"
            >
              Vamos Conversar ?
            </a>

            <div className="flex items-center gap-8 mt-8">
              <a
                href="https://instagram.com"
                target="_blank"
                rel="noopener noreferrer"
                className="hover:opacity-60 transition-opacity"
              >
                <Instagram size={28} strokeWidth={1.5} />
              </a>
              <a
                href="https://wa.me/#"
                target="_blank"
                rel="noopener noreferrer"
                className="hover:opacity-60 transition-opacity"
              >
                <MessageCircle size={28} strokeWidth={1.5} />
              </a>
            </div>
          </div>
      </div>
    </>
  );
};

export default Navbar;
