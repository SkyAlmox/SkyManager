import React from 'react';
import { MapPin } from 'lucide-react';

interface FooterProps {
  onLinkClick: (e: React.MouseEvent<HTMLAnchorElement>, targetId: string) => void;
}

const Footer: React.FC<FooterProps> = ({ onLinkClick }) => {
  return (
    <footer className="bg-brand-bg/80 border-t border-brand-hover/20 py-12 px-6 text-brand-text">
      <div className="max-w-[1800px] mx-auto flex flex-col lg:flex-row items-center gap-12">
        <div className="flex flex-col md:flex-row gap-12 md:gap-24 w-full lg:w-auto">
          <a 
            href="https://www.google.com/maps/place/R.+Cop%C3%A9rnico+Pinto+Coelho,+91+-+Santa+L%C3%BAcia,+Belo+Horizonte+-+MG,+30350-290/data=!4m2!3m1!1s0xa69780cb645323:0x953c95182ebdd6e?sa=X&ved=2ahUKEwjt8MjW7oSBAxVpHbkGHZ7vAQgQ8gF6BAgXEAA&ved=2ahUKEwjt8MjW7oSBAxVpHbkGHZ7vAQgQ8gF6BAgaEAI"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-6 group hover:text-brand-hover transition-colors w-fit"
          >
            <MapPin size={40} strokeWidth={1} className="text-brand-text/40 group-hover:text-brand-hover transition-colors flex-shrink-0" />
            <div className="font-light leading-relaxed text-sm text-left">
              <p className="font-bold mb-1 uppercase tracking-widest text-[10px] text-brand-hover">Unidade Santo Antônio</p>
              <p>Rua Copérnico Pinto Coelho, 91</p>
              <p>Belo Horizonte - MG - CEP 30.350-290</p>
              <p>(31) 3213 1919</p>
            </div>
          </a>

          <a 
            href="https://www.google.com/maps/place/R.+Gurutuba,+336+-+Nova+Esperan%C3%A7a,+Belo+Horizonte+-+MG,+31230-210/@-19.902787,-43.9587055,17z/data=!3m1!4b1!4m5!3m4!1s0xa690b309301537:0x85599a457d6b026a!8m2!3d-19.902787!4d-43.9561306?entry=ttu&g_ep=EgoyMDI2MDIxOC4wIKXMDSoASAFQAw%3D%3D"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-6 group hover:text-brand-hover transition-colors w-fit"
          >
            <MapPin size={40} strokeWidth={1} className="text-brand-text/40 group-hover:text-brand-hover transition-colors flex-shrink-0" />
            <div className="font-light leading-relaxed text-sm text-left">
              <p className="font-bold mb-1 uppercase tracking-widest text-[10px] text-brand-hover">Unidade Nova Esperança</p>
              <p>Rua Gurutuba, 336</p>
              <p>Belo Horizonte - MG - CEP 31.230-210</p>
              <p>(31) 3213 1919</p>
            </div>
          </a>
        </div>

        <div className="flex-1 flex flex-col items-center justify-center lg:pt-4">
          <p className="text-[10px] uppercase tracking-[0.2em] opacity-40 mb-2 text-center">
            © 2026 SKYMÍDIA
          </p>
          <p className="text-[9px] uppercase tracking-[0.2em] opacity-30 text-center">
            Todos os direitos reservados
          </p>
        </div>
      </div>
    </footer>
  );
};

export default Footer;
