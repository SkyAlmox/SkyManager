import React, { useState, useEffect } from 'react';
import { BRAND_NAME, JOURNAL_ARTICLES } from '../constants';
import { MessageCircle, ChevronDown } from 'lucide-react';
import { JournalArticle } from '../types';
import { motion, AnimatePresence } from 'framer-motion';

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

  const textColorClass = 'text-brand-text';

  const menuVariants = {
    hidden: { opacity: 0, scale: 0.95, filter: 'blur(10px)' },
    visible: { 
      opacity: 1, 
      scale: 1,
      filter: 'blur(0px)',
      transition: { 
        duration: 0.6, 
        ease: [0.16, 1, 0.3, 1] 
      }
    },
    exit: { 
      opacity: 0, 
      scale: 1.05,
      filter: 'blur(10px)',
      transition: { 
        duration: 0.4, 
        ease: [0.7, 0, 0.84, 0] 
      }
    }
  };

  const containerVariants = {
    hidden: { opacity: 0 },
    visible: {
      opacity: 1,
      transition: {
        staggerChildren: 0.12,
        delayChildren: 0.2
      }
    }
  };

  const itemVariants = {
    hidden: { opacity: 0, y: 30, scale: 0.9 },
    visible: { 
      opacity: 1, 
      y: 0,
      scale: 1,
      transition: { 
        duration: 0.7, 
        ease: [0.16, 1, 0.3, 1] 
      }
    }
  };

  return (
    <>
      <nav 
        className={`fixed top-0 left-0 right-0 z-50 transition-all duration-700 ease-in-out ${
          scrolled || mobileMenuOpen ? 'bg-brand-bg/90 backdrop-blur-md py-4 shadow-sm' : 'bg-transparent py-8'
        }`}
      >
        <div className="w-full max-w-[1800px] mx-auto px-4 md:px-8 flex items-center justify-between">
          <a 
            href="#" 
            onClick={(e) => {
                e.preventDefault();
                window.scrollTo({ top: 0, behavior: 'smooth' });
                onNavClick(e, '');
            }}
            className="z-50 relative block"
          >
            <img 
              src="/logo.png" 
              alt={BRAND_NAME} 
              className="h-12 md:h-16 w-auto transition-all duration-500"
            />
          </a>
          
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
              
              <div className={`absolute top-full left-1/2 -translate-x-1/2 pt-4 transition-all duration-300 ${dropdownOpen ? 'opacity-100 translate-y-0 pointer-events-auto' : 'opacity-0 -translate-y-2 pointer-events-none'}`}>
                <div className="bg-brand-bg border border-brand-hover shadow-xl py-6 min-w-[280px] rounded-sm">
                  {JOURNAL_ARTICLES.filter(a => a.id !== 7).map((article) => (
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
                      {article.submenuIcon && (
                        <img src={article.submenuIcon} alt="" className="h-6 w-auto object-contain" referrerPolicy="no-referrer" />
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

          <div className={`flex items-center gap-3 md:gap-6 z-50 relative transition-colors duration-500 ${textColorClass}`}>
            <div className="hidden sm:flex items-center gap-5">
              <a 
                href="https://www.instagram.com/skymidiabh/" 
                target="_blank" 
                rel="noopener noreferrer"
                className="text-brand-text/40 hover:text-brand-hover transition-colors"
                aria-label="Instagram"
              >
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  width="20"
                  height="20"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.5"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <rect width="20" height="20" x="2" y="2" rx="5" ry="5" />
                  <path d="M16 11.37A4 4 0 1 1 12.63 8 4 4 0 0 1 16 11.37z" />
                  <line x1="17.5" x2="17.51" y1="6.5" y2="6.5" />
                </svg>
              </a>
              <a 
                href="https://wa.me/#" 
                target="_blank" 
                rel="noopener noreferrer"
                className="text-brand-text/40 hover:text-brand-hover transition-colors"
                aria-label="WhatsApp"
              >
                <MessageCircle size={20} strokeWidth={1.5} />
              </a>
            </div>
            
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

      <AnimatePresence>
        {mobileMenuOpen && (
          <motion.div 
            variants={menuVariants}
            initial="hidden"
            animate="visible"
            exit="exit"
            className="fixed inset-0 bg-brand-bg z-40 flex flex-col items-center overflow-y-auto"
          >
            <motion.div 
              variants={containerVariants}
              initial="hidden"
              animate="visible"
              className="flex flex-col items-center space-y-8 text-xl font-serif font-medium text-brand-text w-full px-8 pt-32 pb-12"
            >
              <motion.a 
                variants={itemVariants}
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
              </motion.a>
              
              <motion.div variants={itemVariants} className="flex flex-col items-center w-full">
                <button 
                  onClick={() => setDropdownOpen(!dropdownOpen)}
                  className="hover:text-brand-hover transition-colors flex items-center gap-2 mb-4"
                >
                  Produtos <ChevronDown size={20} className={`transition-transform duration-300 ${dropdownOpen ? 'rotate-180' : ''}`} />
                </button>
                
                {dropdownOpen && (
                  <div className="flex flex-col items-center space-y-4 mb-4 animate-fade-in">
                    {JOURNAL_ARTICLES.filter(a => a.id !== 7).map((article) => (
                      <a 
                        key={article.id}
                        href="#journal"
                        onClick={(e) => {
                          e.preventDefault();
                          setMobileMenuOpen(false);
                          setDropdownOpen(false);
                          onArticleClick(article);
                        }}
                        className="flex flex-col items-center gap-2 text-sm font-sans uppercase tracking-widest text-brand-text/70 hover:text-brand-hover text-center w-full max-w-[250px]"
                      >
                        {article.submenuIcon && (
                          <img src={article.submenuIcon} alt="" className="h-10 w-auto object-contain mx-auto" referrerPolicy="no-referrer" />
                        )}
                        <span className="block w-full">{article.title}</span>
                      </a>
                    ))}
                  </div>
                )}
              </motion.div>

              <motion.a 
                variants={itemVariants}
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
              </motion.a>
              
              <motion.div variants={itemVariants} className="flex items-center gap-8 mt-8">
                <a 
                  href="https://www.instagram.com/skymidiabh/" 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="text-brand-text/40 hover:text-brand-hover transition-colors"
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    width="28"
                    height="28"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="1.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <rect width="20" height="20" x="2" y="2" rx="5" ry="5" />
                    <path d="M16 11.37A4 4 0 1 1 12.63 8 4 4 0 0 1 16 11.37z" />
                    <line x1="17.5" x2="17.51" y1="6.5" y2="6.5" />
                  </svg>
                </a>
                <a 
                  href="https://wa.me/#" 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="text-brand-text/40 hover:text-brand-hover transition-colors"
                >
                  <MessageCircle size={28} strokeWidth={1.5} />
                </a>
              </motion.div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </>
  );
};

export default Navbar;
