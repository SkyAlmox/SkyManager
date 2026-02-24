/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/


import React from 'react';
import { motion } from 'framer-motion';

const About: React.FC = () => {
  const fadeInVariants = {
    hidden: { opacity: 0, y: 30 },
    visible: { 
      opacity: 1, 
      y: 0,
      transition: { duration: 0.8, ease: [0.21, 0.47, 0.32, 0.98] }
    }
  };

  const imageVariants = {
    hidden: { scale: 1.1, opacity: 0 },
    visible: { 
      scale: 1, 
      opacity: 1,
      transition: { duration: 1.2, ease: "easeOut" }
    }
  };

  return (
    <section id="about" className="bg-brand-text pt-32">
      
      {/* Philosophy Blocks (Formerly Features) */}
      <div className="grid grid-cols-1 lg:grid-cols-2 min-h-[80vh]">
        <motion.div 
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          variants={imageVariants}
          className="order-2 lg:order-1 relative h-[500px] lg:h-auto overflow-hidden group"
        >
           <img 
             src="/identidade_visual_01.jpg" 
             alt="Identidade Visual - NEOOH" 
             className="absolute inset-0 w-full h-full object-cover transition-transform duration-[2s] group-hover:scale-105"
           />
        </motion.div>
        <motion.div 
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          variants={fadeInVariants}
          className="order-1 lg:order-2 flex flex-col justify-center p-12 lg:p-24 bg-brand-text"
        >
           <span className="text-xs font-bold uppercase tracking-[0.2em] text-brand-dark/60 mb-6">Identidade Visual</span>
           <h3 className="text-4xl md:text-5xl font-serif mb-8 text-brand-dark leading-tight">
             Vamos falar de DESIGN?
           </h3>
           <p className="text-lg text-brand-dark/80 font-light leading-relaxed mb-12 max-w-md">
             Sim, o DESIGN está presente em todos os projetos que envolvem a comunicação visual e nós respeitamos muito isso, pois ele é o ponto de partida para que a identidade da sua empresa ganhe aplicações corretas e legítimas em todos os modelos de construção.
           </p>
        </motion.div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 min-h-[80vh]">
        <motion.div 
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          variants={fadeInVariants}
          className="flex flex-col justify-center p-12 lg:p-24 bg-brand-text text-brand-dark"
        >
           <span className="text-xs font-bold uppercase tracking-[0.2em] text-brand-dark/60 mb-6">Comunicação Visual</span>
           <p className="text-lg text-brand-dark/80 font-light leading-relaxed mb-12 max-w-md">
             O que vamos te mostrar a seguir são exemplos da nossa capacidade técnica para transformar chapas de aço, alumínio, acrílico e outros diversos materiais, em peças de Identidade Visual capazes de representar com excelência a sua marca.
             <br/><br/>
             Para entregarmos projetos com resultado estético de alta qualidade contamos com pessoas muito bem treinadas e equipamentos de alta tecnologia e precisão alinhados à processos de produção cada dia mais eficientes para economizar tempo e matéria prima.
           </p>
        </motion.div>
        <motion.div 
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, margin: "-100px" }}
          variants={imageVariants}
          className="relative h-[500px] lg:h-auto overflow-hidden group"
        >
           <img 
             src="/identidade_visual_02.jpg" 
             alt="Comunicação Visual - NEOOH" 
             className="absolute inset-0 w-full h-full object-cover transition-transform duration-[2s] group-hover:scale-105 brightness-90"
           />
        </motion.div>
      </div>
    </section>
  );
};

export default About;