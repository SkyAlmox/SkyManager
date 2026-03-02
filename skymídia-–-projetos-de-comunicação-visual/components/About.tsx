import React from 'react';
import { motion } from 'framer-motion';
import ImageCarousel from './ImageCarousel';

const About: React.FC = () => {
  const slideInLeft = {
    hidden: { opacity: 0, x: -50 },
    visible: { 
      opacity: 1, 
      x: 0,
      transition: { duration: 1, ease: [0.21, 0.47, 0.32, 0.98] }
    }
  };

  const slideInRight = {
    hidden: { opacity: 0, x: 50 },
    visible: { 
      opacity: 1, 
      x: 0,
      transition: { duration: 1, ease: [0.21, 0.47, 0.32, 0.98] }
    }
  };

  const imageVariants = {
    hidden: { scale: 1.1, opacity: 0, filter: 'blur(10px)' },
    visible: { 
      scale: 1, 
      opacity: 1,
      filter: 'blur(0px)',
      transition: { duration: 1.4, ease: "easeOut" }
    }
  };

  return (
    <section id="about" className="pt-32 overflow-hidden">
      <div className="grid grid-cols-1 lg:grid-cols-2 min-h-[80vh]">
        <motion.div 
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          variants={slideInLeft}
          className="order-2 lg:order-1 relative h-[500px] lg:h-auto overflow-hidden"
        >
           <ImageCarousel 
             images={['/identidade_visual_01.jpg', '/identidade_visual_02.jpg', '/identidade_visual_03.jpg']} 
             alt="Identidade Visual"
             className="absolute inset-0 w-full h-full"
           />
        </motion.div>
        <motion.div 
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          variants={slideInRight}
          className="order-1 lg:order-2 flex flex-col justify-center p-12 lg:p-24"
        >
           <span className="text-xs font-bold uppercase tracking-[0.2em] text-brand-text/60 mb-6">Identidade Visual</span>
           <h3 className="text-4xl md:text-5xl font-serif mb-8 text-brand-text leading-tight">
             Vamos falar de DESIGN?
           </h3>
           <p className="text-lg text-brand-text/80 font-light leading-relaxed mb-12 max-w-md">
             Sim, o DESIGN está presente em todos os projetos que envolvem a comunicação visual e nós respeitamos muito isso, pois ele é o ponto de partida para que a identidade da sua empresa ganhe aplicações corretas e legítimas em todos os modelos de construção.
           </p>
        </motion.div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 min-h-[80vh]">
        <motion.div 
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          variants={slideInLeft}
          className="flex flex-col justify-center p-12 lg:p-24 text-brand-text"
        >
           <span className="text-xs font-bold uppercase tracking-[0.2em] text-brand-text/60 mb-6">Comunicação Visual</span>
           <p className="text-lg text-brand-text/80 font-light leading-relaxed mb-12 max-w-md">
             O que vamos te mostrar a seguir são exemplos da nossa capacidade técnica para transformar chapas de aço, alumínio, acrílico e outros diversos materiais, em peças de Identidade Visual capazes de representar com excelência a sua marca.
             <br/><br/>
             Para entregarmos projetos com resultado estético de alta qualidade contamos com pessoas muito bem treinadas e equipamentos de alta tecnologia e precisão alinhados à processos de produção cada dia mais eficientes para economizar tempo e matéria prima.
           </p>
        </motion.div>
        <motion.div 
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          variants={slideInRight}
          className="relative h-[500px] lg:h-auto overflow-hidden"
        >
           <ImageCarousel 
             images={['/identidade_visual_04.jpg', '/identidade_visual_05.jpg', '/identidade_visual_06.jpg']} 
             alt="Comunicação Visual"
             className="absolute inset-0 w-full h-full brightness-90"
           />
        </motion.div>
      </div>
    </section>
  );
};

export default About;
