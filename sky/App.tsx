import React, { useState } from 'react';
import Navbar from './components/Navbar';
import Hero from './components/Hero';
import About from './components/About';
import Journal from './components/Journal';
import Assistant from './components/Assistant';
import Footer from './components/Footer';
import ProductDetail from './components/ProductDetail';
import JournalDetail from './components/JournalDetail';
import Checkout from './components/Checkout';
import { Product, JournalArticle, ViewState } from './types';
import { JOURNAL_ARTICLES } from './constants';

function App() {
  // Test comment
  const [view, setView] = useState<ViewState>({ type: 'home' });
  const [cartItems, setCartItems] = useState<Product[]>([]);
  const [isCartOpen, setIsCartOpen] = useState(false);

  // Handle navigation (clicks on Navbar or Footer links)
  const handleNavClick = (e: React.MouseEvent<HTMLAnchorElement>, targetId: string) => {
    e.preventDefault();

    // If we are not home, go home first
    if (view.type !== 'home') {
      setView({ type: 'home' });
      // Allow state update to render Home before scrolling
      setTimeout(() => scrollToSection(targetId), 0);
    } else {
      scrollToSection(targetId);
    }
  };

  const scrollToSection = (targetId: string) => {
    if (!targetId) {
        window.scrollTo({ top: 0, behavior: 'smooth' });
        return;
    }

    const element = document.getElementById(targetId);
    if (element) {
      // Manual scroll calculation to account for fixed header
      const headerOffset = 85;
      const elementPosition = element.getBoundingClientRect().top;
      const offsetPosition = elementPosition + window.scrollY - headerOffset;

      window.scrollTo({
        top: offsetPosition,
        behavior: "smooth"
      });

      try {
        window.history.pushState(null, '', `#${targetId}`);
      } catch (err) {
        // Ignore SecurityError in restricted environments
      }
    }
  };

  const addToCart = (product: Product) => {
    setCartItems([...cartItems, product]);
    setIsCartOpen(true);
  };

  const removeFromCart = (index: number) => {
    const newItems = [...cartItems];
    newItems.splice(index, 1);
    setCartItems(newItems);
  };

  return (
    <div className="min-h-screen bg-brand-bg font-sans text-brand-text selection:bg-brand-hover selection:text-brand-bg">
      {view.type !== 'checkout' && (
        <Navbar
            onNavClick={handleNavClick}
            onArticleClick={(a) => {
                window.scrollTo({ top: 0, behavior: 'smooth' });
                setView({ type: 'journal', article: a });
            }}
        />
      )}

      <main>
        {view.type === 'home' && (
          <>
            <Hero onNavClick={handleNavClick} />
            <About />
            <Journal onArticleClick={(a) => {
                window.scrollTo({ top: 0, behavior: 'smooth' });
                setView({ type: 'journal', article: a });
            }} />
          </>
        )}

        {view.type === 'product' && (
          <ProductDetail
            product={view.product}
            onBack={() => {
              setView({ type: 'home' });
              setTimeout(() => scrollToSection('products'), 50);
            }}
            onAddToCart={addToCart}
          />
        )}

        {view.type === 'journal' && (
          <JournalDetail
            article={view.article}
            onBack={() => setView({ type: 'home' })}
            onNavigateToContact={() => {
              const contactArticle = JOURNAL_ARTICLES.find(a => a.id === 7);
              if (contactArticle) {
                window.scrollTo({ top: 0, behavior: 'smooth' });
                setView({ type: 'journal', article: contactArticle });
              }
            }}
          />
        )}

        {view.type === 'checkout' && (
            <Checkout
                items={cartItems}
                onBack={() => setView({ type: 'home' })}
            />
        )}
      </main>

      {view.type !== 'checkout' && <Footer onLinkClick={handleNavClick} />}

      <Assistant />
    </div>
  );
}

export default App;
