import React from 'react';
import { PRIMARY_COLOR, ACCENT_COLOR } from '../constants';

const Hero: React.FC = () => {
  const handleNavClick = (
    e: React.MouseEvent<HTMLAnchorElement>,
    targetId: string
  ) => {
    e.preventDefault();
    const element = document.getElementById(targetId);
    if (element) {
      const headerOffset = 85;
      const elementPosition = element.getBoundingClientRect().top;
      const offsetPosition = elementPosition + window.scrollY - headerOffset;

      window.scrollTo({
        top: offsetPosition,
        behavior: 'smooth',
      });

      try {
        window.history.pushState(null, '', `#${targetId}`);
      } catch (err) {}
    }
  };

  return (
    <section
      className="relative w-full h-screen min-h-[800px] overflow-hidden"
      style={{ backgroundColor: PRIMARY_COLOR }}
    >
      {/* Background Image */}
      <div className="absolute inset-0 w-full h-full">
        <img
          src="/hero2.png?auto=format&fit=crop&q=80&w=2000"
          alt="Serene misty landscape"
          className="w-full h-full object-cover grayscale contrast-[0.7] brightness-[0.85] opacity-50"
        />
        {/* Overlay para escurecer o fundo */}
        <div className="absolute inset-0 bg-black/50 mix-blend-multiply"></div>
      </div>

      {/* Conteúdo */}
      <div className="relative z-10 h-full flex flex-col justify-center items-start text-left md:items-center md:text-center px-6">
        <div className="animate-fade-in-up w-full md:w-auto">
          <span className="block text-xs md:text-sm font-medium uppercase tracking-[0.2em] text-white mb-6 backdrop-blur-sm bg-white/10 px-4 py-2 rounded-full mx-0 md:mx-auto w-fit">
            Olá
          </span>
          <h1 className="text-6xl md:text-8xl lg:text-9xl font-serif font-normal text-white tracking-tight mb-8 drop-shadow-[0_4px_6px_rgba(0,0,0,0.7)]">
            Somos a{' '}
            <span className="italic">
              SkyMídia
            </span>
          </h1>
          <p className="max-w-lg mx-0 md:mx-auto text-lg md:text-xl text-white font-light leading-relaxed mb-12 drop-shadow-[0_2px_4px_rgba(0,0,0,0.8)] p-4 rounded-lg">
            Transformamos matérias-primas comuns em projetos de Comunicação
            Visual, de modo que cada cliente obtenha uma solução sob medida para
            cada uma das suas necessidades.
            <br />
            <br />
            Hoje temos 6 divisões principais de produção, e vamos te mostrar
            cada uma delas com mais detalhes nas próximas páginas.
          </p>

          <a
            href="#products"
            onClick={(e) => handleNavClick(e, 'products')}
            className="group relative px-10 py-4 rounded-full text-sm font-semibold uppercase tracking-widest transition-all duration-500 overflow-hidden shadow-lg hover:shadow-xl inline-block"
            style={{
              backgroundColor: ACCENT_COLOR,
              color: '#2C2A26',
            }}
          >
            <span className="relative z-10 group-hover:text-[#2C2A26]">
              Saiba Mais!
            </span>
          </a>
        </div>
      </div>

      {/* Indicador de Scroll */}
      <div className="absolute bottom-12 left-1/2 -translate-x-1/2 animate-bounce text-white/70">
        <svg
          className="w-6 h-6"
          fill="none"
          stroke="currentColor"
          viewBox="0 0 24 24"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            strokeWidth={1}
            d="M19 14l-7 7m0 0l-7-7m7 7V3"
          />
        </svg>
      </div>
    </section>
  );
};

export default Hero;