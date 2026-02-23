import React from 'react';
import { MapPin } from 'lucide-react';

interface FooterProps {
  onLinkClick: (e: React.MouseEvent<HTMLAnchorElement>, targetId: string) => void;
}

const Footer: React.FC<FooterProps> = ({ onLinkClick }) => {
  return (
    <footer className="bg-brand-bg/80 border-t border-brand-hover/20 py-8 px-6 text-brand-text">
      <div className="max-w-[1800px] mx-auto flex flex-col md:flex-row items-center">

        <div className="flex flex-col items-start flex-shrink-0">
          <a
            href="https://www.google.com/maps/place/R.+Cop%C3%A9rnico+Pinto+Coelho,+91+-+Santa+L%C3%BAcia,+Belo+Horizonte+-+MG,+30350-290/data=!4m2!3m1!1s0xa69780cb645323:0x953c95182ebdd6e?sa=X&ved=2ahUKEwjt8MjW7oSBAxVpHbkGHZ7vAQgQ8gF6BAgXEAA&ved=2ahUKEwjt8MjW7oSBAxVpHbkGHZ7vAQgQ8gF6BAgaEAI"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-6 group hover:text-brand-hover transition-colors w-fit"
          >
            <div className="font-light leading-relaxed text-sm text-left">
              <p>Rua Copérnico Pinto Coelho, 91 - Santo Antônio</p>
              <p>Belo Horizonte - MG - CEP 30.350-290</p>
              <p>(31) 3213 1919</p>
            </div>
            <MapPin size={56} strokeWidth={1} className="text-brand-text/40 group-hover:text-brand-hover transition-colors flex-shrink-0" />
          </a>
        </div>

        <div className="flex-1 flex justify-center mt-8 md:mt-0">
          <p className="text-xs uppercase tracking-[0.2em] opacity-60">
            © 2026 - Todos os direitos reservados
          </p>
        </div>
      </div>
    </footer>
  );
};

export default Footer;
