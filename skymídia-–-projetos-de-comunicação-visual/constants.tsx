import React from 'react';
import ProductGallery from './components/ProductGallery';
import ContactForm from './components/ContactForm';
import { Product, JournalArticle } from './types';

export const PRODUCTS: Product[] = [
  {
    id: 'p1',
    name: 'Aura Harmony',
    tagline: 'Listen naturally.',
    description: 'Audio that feels like the open air. Constructed with warm acoustic fabric and recycled sandstone composite.',
    longDescription: 'Experience sound as it was meant to be heard—unconfined and organic. The Aura Harmony headphones feature our proprietary open-air driver technology, encased in a breathable acoustic fabric that adapts to your temperature. The headband is crafted from a recycled sandstone composite, offering a unique, cool-to-the-touch texture that grounds you in the present moment.',
    price: 429,
    category: 'Audio',
    imageUrl: 'https://images.unsplash.com/photo-1505740420928-5e560c06d30e?auto=format&fit=crop&q=80&w=1000',
    gallery: [
      'https://images.unsplash.com/photo-1505740420928-5e560c06d30e?auto=format&fit=crop&q=80&w=1000',
      'https://images.unsplash.com/photo-1484704849700-f032a568e944?auto=format&fit=crop&q=80&w=1000',
      'https://images.unsplash.com/photo-1524678606372-565ae0f98944?auto=format&fit=crop&q=80&w=1000'
    ],
    features: ['Organic Noise Cancellation', '50h Battery', 'Natural Soundstage']
  },
  {
    id: 'p2',
    name: 'Aura Epoch',
    tagline: 'Moments, not minutes.',
    description: 'A timepiece designed for wellness. Ceramic casing with a strap made from sustainable vegan leather.',
    longDescription: 'Time is not a sequence of numbers, but a flow of moments. The Aura Epoch rethinks the smartwatch interface, using a calm E-Ink hybrid display that mimics paper. It tracks stress through skin temperature and heart rate variability, gently vibrating to remind you to breathe. The ceramic casing is hypoallergenic and smooth, polished by hand for 48 hours.',
    price: 349,
    category: 'Wearable',
    imageUrl: 'https://images.unsplash.com/photo-1523275335684-37898b6baf30?auto=format&fit=crop&q=80&w=1000',
    gallery: [
        'https://images.unsplash.com/photo-1523275335684-37898b6baf30?auto=format&fit=crop&q=80&w=1000',
        'https://images.unsplash.com/photo-1508685096489-7aacd43bd3b1?auto=format&fit=crop&q=80&w=1000'
    ],
    features: ['Stress Monitoring', 'E-Ink Hybrid Display', '7-Day Battery']
  },
  {
    id: 'p3',
    name: 'Aura Canvas',
    tagline: 'Capture the warmth.',
    description: 'A display that mimics the properties of paper. Soft on the eyes, vivid in color, and textured to the touch.',
    longDescription: 'Screens shouldn\'t feel like looking into a lightbulb. Aura Canvas uses a matte, nano-etched OLED panel that scatters ambient light, creating a display that looks and feels like high-quality magazine paper. Perfect for reading, sketching, or displaying art, it brings a tactile warmth to your digital life.',
    price: 1099,
    category: 'Mobile',
    imageUrl: 'https://images.unsplash.com/photo-1544816155-12df9643f363?auto=format&fit=crop&q=80&w=1000',
    gallery: [
        'https://images.unsplash.com/photo-1544816155-12df9643f363?auto=format&fit=crop&q=80&w=1000',
        'https://images.unsplash.com/photo-1585338107529-13afc5f02586?auto=format&fit=crop&q=80&w=1000'
    ],
    features: ['Paper-like OLED', 'Portrait Lens', 'Sandstone Texture']
  },
  {
    id: 'p4',
    name: 'Aura Essence',
    tagline: 'Return to nature.',
    description: 'An air purifier that doubles as a sculpture. Whisper quiet, diffusing subtle natural scents while cleaning your space.',
    longDescription: 'Clean air is the foundation of a clear mind. Aura Essence uses a moss-based bio-filter combined with HEPA technology to scrub pollutants from your home. It gently diffuses natural essential oils—cedar, bergamot, and rain—orchestrated to match the time of day.',
    price: 599,
    category: 'Home',
    imageUrl: 'https://images.pexels.com/photos/8092420/pexels-photo-8092420.jpeg?auto=compress&cs=tinysrgb&w=1260&h=750&dpr=1',
    gallery: [
        'https://images.pexels.com/photos/8092420/pexels-photo-8092420.jpeg?auto=compress&cs=tinysrgb&w=1260&h=750&dpr=1',
        'https://images.unsplash.com/photo-1513694203232-719a280e022f?auto=format&fit=crop&q=80&w=1000'
    ],
    features: ['Bio-HEPA Filter', 'Aromatherapy', 'Silent Night Mode']
  },
  {
    id: 'p5',
    name: 'Aura Beam',
    tagline: 'Light that breathes.',
    description: 'Smart circadian lighting that follows the sun. Casts a warm, candle-like glow in the evenings.',
    longDescription: 'Artificial light disrupts our natural rhythms. Aura Beam syncs with your local sunrise and sunset, providing cool, energizing light during the day and transitioning to a warm, amber glow free of blue light in the evening. Controls are touchless; a simple wave of the hand adjusts brightness.',
    price: 249,
    category: 'Home',
    imageUrl: 'https://images.unsplash.com/photo-1565814329452-e1efa11c5b89?auto=format&fit=crop&q=80&w=1000',
    gallery: [
        'https://images.unsplash.com/photo-1565814329452-e1efa11c5b89?auto=format&fit=crop&q=80&w=1000',
        'https://images.unsplash.com/photo-1540932296235-d84931b6370b?auto=format&fit=crop&q=80&w=1000'
    ],
    features: ['Circadian Rhythm Sync', 'Warm Dimming', 'Touchless Control']
  },
  {
    id: 'p6',
    name: 'Aura Scribe',
    tagline: 'Thought in motion.',
    description: 'A digital stylus with the friction of graphite. Charges wirelessly when magnetically attached to Aura Canvas.',
    longDescription: 'The connection between hand and brain is sacred. Aura Scribe features a custom elastomer tip that replicates the microscopic friction of graphite on paper. Weighted perfectly for balance, it disappears in your hand, leaving only your thoughts.',
    price: 129,
    category: 'Mobile',
    imageUrl: 'https://images.pexels.com/photos/2647376/pexels-photo-2647376.jpeg?auto=compress&cs=tinysrgb&w=1260&h=750&dpr=1',
    gallery: [
        'https://images.pexels.com/photos/2647376/pexels-photo-2647376.jpeg?auto=compress&cs=tinysrgb&w=1260&h=750&dpr=1',
        'https://images.unsplash.com/photo-1517260487576-8977430081d3?auto=format&fit=crop&q=80&w=1000'
    ],
    features: ['Zero Latency', 'Textured Tip', 'Wireless Charging']
  }
];

export const JOURNAL_ARTICLES: JournalArticle[] = [
    {
        id: 3,
        title: "Estruturas para Mídia OOH",
        date: "Outdoor Media",
        excerpt: "Engenharia e fabricação de estruturas robustas para publicidade externa de alto impacto.",
        image: "/estruturas_ooh_00.jpg",
        icon: "/paineis_para_midia_ooh.png",
        submenuIcon: "/icon_paineis_para_midia_ooh.png",
        content: (
            <>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    A Mídia OUT OF HOME ( OOH ) ocupa cada vez mais espaço nos aeroportos, terminais rodoviários e urbanos, parques, shoppings e centros comerciais.
                </p>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    A evolução dos painéis de led, cada vez maiores e com formatos personalizados, exige construções confiáveis, estáveis e de alta qualidade.
                </p>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Somos especialistas na construção e instalação de estruturas metálicas de grande porte com alta precisão e qualidade de acabamento e estamos prontos para atender com velocidade e capacidade técnica a sua empresa de OOH.
                </p>
                <ProductGallery
                    categories={[
                        {
                            title: 'MÍDIA OOH',
                            subtitle: 'Vídeos Wall | Painéis | Totens',
                            images: ['/estruturas_ooh_01.jpg','/estruturas_ooh_02.jpg','/estruturas_ooh_03.jpg','/estruturas_ooh_04.jpg','/estruturas_ooh_05.jpg','/estruturas_ooh_06.jpg','/estruturas_ooh_07.jpg','/estruturas_ooh_08.jpg','/estruturas_ooh_09.jpg','/estruturas_ooh_10.jpg','/estruturas_ooh_11.jpg','/estruturas_ooh_12.jpg','/estruturas_ooh_13.jpg','/estruturas_ooh_14.jpg','/estruturas_ooh_15.jpg','/estruturas_ooh_16.jpg','/estruturas_ooh_17.jpg','/estruturas_ooh_18.jpg','/estruturas_ooh_19.jpg','/estruturas_ooh_20.jpg','/estruturas_ooh_21.jpg']
                        }
                    ]}
                />
            </>
        )
    },
    {
        id: 2,
        title: "Comunicação Visual",
        date: "Identidade & Sinalização",
        excerpt: "Projetos completos de sinalização e branding que destacam sua marca no mercado.",
        image: "/comunicacao_visual_placaseletreiros_00.jpg",
        icon: "/comunicacao_visual.png",
        submenuIcon: "/icon_comunicacao_visual.png",
        content: (
            <>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Todas as formas, cores e reproduções da sua MARCA precisam seguir rigorosamente as características técnicas do seu Manual de Identidade Visual para que o resultado seja uma peça exclusiva e reconhecida por todos os públicos.
                    <br/><br/>
                    Para entregarmos projetos com resultado estético de alta qualidade contamos com pessoas muito bem treinadas e equipamentos de alta tecnologia e precisão alinhados à processos de produção cada dia mais eficientes para economizar tempo e matéria prima.
                </p>
                <ProductGallery categories={[{title: 'PLACAS E LETREIROS', subtitle: '2D | 3D | Com Iluminação | Sem Iluminação', images: ['/comunicacao_visual_placaseletreiros_01.jpg', '/comunicacao_visual_placaseletreiros_02.jpg', '/comunicacao_visual_placaseletreiros_03.jpg', '/comunicacao_visual_placaseletreiros_04.jpg', '/comunicacao_visual_placaseletreiros_05.jpg', '/comunicacao_visual_placaseletreiros_06.jpg', '/comunicacao_visual_placaseletreiros_07.jpg', '/comunicacao_visual_placaseletreiros_08.jpg', '/comunicacao_visual_placaseletreiros_09.jpg', '/comunicacao_visual_placaseletreiros_10.jpg', '/comunicacao_visual_placaseletreiros_11.jpg', '/comunicacao_visual_placaseletreiros_12.jpg', '/comunicacao_visual_placaseletreiros_13.jpg', '/comunicacao_visual_placaseletreiros_14.jpg', '/comunicacao_visual_placaseletreiros_15.jpg', '/comunicacao_visual_placaseletreiros_16.jpg']}, {title: 'NEON', subtitle: ' Tradicional | Led ', images: ['/comunicacao_visual_neon_01.jpeg', '/comunicacao_visual_neon_02.jpg', '/comunicacao_visual_neon_03.jpg', '/comunicacao_visual_neon_04.jpg']}, {title: 'TOTENS', subtitle: ' Sinalização | Institucionais | Promocionais ', images: ['/comunicacao_visual_totens_01.jpg', '/comunicacao_visual_totens_02.jpeg', '/comunicacao_visual_totens_03.jpg', '/comunicacao_visual_totens_04.jpg']}, {title: 'SINALIZAÇÃO', subtitle: ' Externa | Interna ', images: ['/comunicacao_visual_sinalizacao_01.jpg', '/comunicacao_visual_sinalizacao_02.jpg', '/comunicacao_visual_sinalizacao_03.jpg', '/comunicacao_visual_sinalizacao_04.jpg', '/comunicacao_visual_sinalizacao_05.jpg']}, {title: 'PLOTAGENS', subtitle: ' Plotter de Recorte | Grandes Formatos ', images: ['/comunicacao_visual_plotagens_01.jpg', '/comunicacao_visual_plotagens_02.jpg', '/comunicacao_visual_plotagens_03.jpg', '/comunicacao_visual_plotagens_04.jpg', '/comunicacao_visual_plotagens_05.jpg', '/comunicacao_visual_plotagens_06.jpg', '/comunicacao_visual_plotagens_07.jpg', '/comunicacao_visual_plotagens_08.jpg', '/comunicacao_visual_plotagens_09.jpg']}, {title: 'PAINÉIS IMPRESSOS', subtitle: ' Adesivo Acoplado | Lonas | Tecidos ', images: ['/comunicacao_visual_paineisimpressos_01.jpg', '/comunicacao_visual_paineisimpressos_02.jpg', '/comunicacao_visual_paineisimpressos_03.jpg', '/comunicacao_visual_paineisimpressos_04.jpg']}, {title: 'PROJETOS ESPECIAIS', subtitle: ' 100% Personalizados ',  images: ['/comunicacao_visual_projetosespeciais_01.jpg', '/comunicacao_visual_projetosespeciais_02.jpg', '/comunicacao_visual_projetosespeciais_03.jpg', '/comunicacao_visual_projetosespeciais_04.jpg']}]} />
            </>
        )
    },
    {
        id: 5,
        title: "Projetos Especiais de Arquitetura",
        date: "Arquitetura & Design",
        excerpt: "Execução de projetos arquitetônicos complexos e instalações artísticas sob medida.",
        image: "/projetos_especiais_arquitetura_00.jpg",
        icon: "/projetos_para_arquitetura.png",
        submenuIcon: "/icon_projetos_para_arquitetura.png",
        content: (
            <>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Quando a Comunicação Visual, o Design e a Arquitetura se encontram, projetos incríveis saem do papel.
                </p>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    São construções que exigem uma estrutura muito especializada em serralheria técnica, pintura (eletrostática e automotiva) elétrica, eletrônica e montagem para transformar desenhos técnicos em peças de alto valor estético e conceitual.
                </p>
                <ProductGallery
                    categories={[
                        {
                        title: 'TUDO AQUI',
                        subtitle: 'Serralheria Técnica | Pinturas Especiais | Elétrica e Eletrônica | Montagem',
                        images: ['/projetos_especiais_arquitetura_01.jpg','/projetos_especiais_arquitetura_02.jpg','/projetos_especiais_arquitetura_03.jpg', '/projetos_especiais_arquitetura_04.jpg','/projetos_especiais_arquitetura_05.jpg', '/projetos_especiais_arquitetura_06.jpg','/projetos_especiais_arquitetura_07.jpg', '/projetos_especiais_arquitetura_08.jpg','/projetos_especiais_arquitetura_09.jpg']}
                    ]}
                />
            </>
        )
    },
    {
        id: 4,
        title: "Marcenaria",
        date: "Mobiliário Sob Medida",
        excerpt: "Mobiliário corporativo e comercial executado com precisão e acabamento superior.",
        image: "/marcenaria_00.jpg",
        icon: "/marcenaria.png",
        submenuIcon: "/icon_marcenaria.png",
        content: (
            <>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Nossa marcenaria combina técnicas tradicionais com tecnologia CNC para criar peças exclusivas que otimizam espaços e reforçam a estética da sua marca.
                </p>
                <ProductGallery 
                    categories={[
                        {
                            title: 'MARCENARIA',
                            subtitle: 'x1',
                            images: ['/marcenaria_00.jpg']
                        }
                    ]}
                />
            </>
        )
    },
    {
        id: 6,
        title: "Totens Digitais e Interativos",
        date: "Tecnologia & Interação",
        excerpt: "Hardware de ponta para sinalização digital e autoatendimento with design sofisticado.",
        image: "/totens_00.jpg",
        icon: "/totens_digitais_e_interativos.png",
        submenuIcon: "/icon_totens_digitais_e_interativos.png",
        content: (
            <>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Pagar o estacionamento do shopping, comprar um sanduiche ou escolher um filme no cinema, de forma interativa, é uma realidade que, a cada dia, ganha mais espaço nos pontos de venda.
                </p>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Nesse mesmo ponto de venda o espaço físico dos cartazes, banners e painéis que divulgam as mensagens promocionais e institucionais se multiplicou com o uso das telas e painéis de led. Sim, estamos falando dos Totens Digitais e Interativos!
                </p>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Aqui na SKYMIDIA oferecemos a solução completa que parte do desenvolvimento técnico do projeto, construção e instalação no ponto veiculação. Uma solução 100% personalizada com a Identidade Visual da sua empresa.
                </p>
                <ProductGallery
                    categories={[
                        {
                            title: 'TOTENS',
                            subtitle: 'Autoatendimento | Digitais | Interativos',
                            images: ['/totens_01.jpeg','/totens_02.jpeg','/totens_03.jpeg','/totens_04.jpeg','/totens_05.jpeg','/totens_06.jpg']
                        }
                    ]}
                />
            </>
        )
    },
    {
        id: 1,
        title: "Carregadores de Celular",
        date: "Soluções de Energia",
        excerpt: "Estações de carregamento inteligentes e seguras para ambientes públicos e privados.",
        image: "/carregadores_celular_00.jpg",
        icon: "/carregadores_de_celular.png",
        submenuIcon: "/icon_carregadores_de_celular.png",
        content: (
            <>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Sim, temos mais celulares do que habitantes no Brasil. 270 MILHÕES de aparelhos, segundo a Anatel em dezembro de 2025.
                </p>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Esse número gigantesco se traduz em uma grande oportunidade comercial, afinal todos esses milhões de aparelhos são carregados diariamente, e isso está acontecendo cada vez mais fora de casa.
                </p>
                <p className="mb-12 text-brand-text text-center max-w-2xl mx-auto text-xl">
                    Imagine como um benefício simples e direto, como uma carga de celular, é capaz de fazer o consumidor ficar mais tempo no seu ponto de venda. Faz sentido? Se esse benefício é oferecido de forma mais inteligente, do que uma simples tomada - sim, faz todo sentido!
                </p>
                <ProductGallery
                    categories={[
                        {
                        title: 'SOB MEDIDA',
                        subtitle: 'Fotovoltaico | Indução | Por Cabo',
                        images: ['/carregadores_celular_01.jpg','/carregadores_celular_02.jpg','/carregadores_celular_03.jpg','/carregadores_celular_04.jpg','/carregadores_celular_05.jpg','/carregadores_celular_06.jpg','/carregadores_celular_07.jpg','/carregadores_celular_08.jpg','/carregadores_celular_09.jpg','/carregadores_celular_10.jpg','/carregadores_celular_11.jpg']
                        }
                    ]}
                />
            </>
        )
    },
    {
        id: 7,
        title: "LEGAL!",
        date: "Vamos Conversar?",
        excerpt: "Entre em contato com a nossa equipe e tire o seu projeto do papel.",
        image: "https://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&q=80&w=1000",
        content: (
            <div className="space-y-12">
                <p className="text-brand-text text-center max-w-2xl mx-auto text-xl leading-relaxed">
                    Agora que você conheceu um pouco do que somos capazes de fazer, transformado matérias-primas comuns em projetos exclusivos, vamos falar rapidamente sobre a nossa estrutura e como você pode entrar em contato com a gente para tirar o seu projeto do papel.
                </p>
                
                <div className="grid md:grid-cols-2 gap-12 items-center bg-brand-dark/5 p-8 md:p-12 rounded-sm border border-brand-hover/10">
                    <div className="space-y-6">
                        <p className="text-brand-text leading-relaxed">
                            Localizada no bairro Santo Antônio, nossa planta industrial possui mais de 1.000 m² de área de produção. O espaço é equipado com tecnologia de alta precisão e desempenho, incluindo estações de corte a laser, Router CNC, setores de montagem, serralheria técnica, cabine de pintura e laboratórios de elétrica e eletrônica. Além disso, contamos com uma equipe especializada em transporte e instalação de estruturas de grande porte, devidamente certificada conforme as normas de segurança vigentes.
                        </p>
                    </div>
                    <div className="aspect-video overflow-hidden rounded-sm">
                        <img 
                            src="/galpao1.png" 
                            alt="Planta Industrial Santo Antônio" 
                            className="w-full h-full object-cover"
                            referrerPolicy="no-referrer"
                        />
                    </div>
                </div>

                <div className="grid md:grid-cols-2 gap-12 items-center bg-brand-dark/5 p-8 md:p-12 rounded-sm border border-brand-hover/10">
                    <div className="aspect-video overflow-hidden rounded-sm order-2 md:order-1">
                        <img 
                            src="/galpao2.png" 
                            alt="Planta Industrial Nova Esperança" 
                            className="w-full h-full object-cover"
                            referrerPolicy="no-referrer"
                        />
                    </div>
                    <div className="space-y-6 order-1 md:order-2">
                        <p className="text-brand-text leading-relaxed">
                            No bairro Nova Esperança, mantemos outra planta industrial com mais de 500 m² de área de produção. Nessa unidade, são realizados serviços de solda, serralheria, pintura e montagem, garantindo a qualidade e a eficiência na fabricação de nossas estruturas.
                        </p>
                    </div>
                </div>

                <div className="flex flex-col items-center gap-12 pt-12">
                    <div className="text-center space-y-3 max-w-xl mx-auto">
                        <p className="text-2xl md:text-3xl font-serif text-brand-text leading-tight">
                            Pelo <span className="text-brand-hover font-bold">WhatsApp</span> é mais <span className="italic font-bold text-brand-hover">RÁPIDO</span>!
                        </p>
                        <p className="text-xs uppercase tracking-[0.2em] text-brand-text/40 font-medium">
                            Mas, fique à vontade para ligar ou enviar um e-mail, ok?
                        </p>
                    </div>
                    
                    <div className="grid md:grid-cols-2 gap-8 w-full">
                        <div className="bg-brand-dark/5 p-8 rounded-sm border border-brand-hover/10 flex flex-col items-center text-center space-y-4">
                            <span className="text-[10px] uppercase tracking-[0.2em] text-brand-hover font-bold">Comercial</span>
                            <h4 className="text-xl font-serif text-brand-text">RODRIGO DOMUSCI</h4>
                            <div className="space-y-2 pt-4 w-full">
                                <a 
                                    href="https://wa.me/5531983315015" 
                                    target="_blank" 
                                    rel="noopener noreferrer"
                                    className="flex items-center justify-center gap-2 text-sm hover:text-brand-hover transition-colors text-brand-text"
                                >
                                    <span className="opacity-40 uppercase tracking-widest text-[10px]">WhatsApp:</span>
                                    <span className="font-medium">+55 (31) 9 8331.5015</span>
                                </a>
                                <a 
                                    href="mailto:rodrigo@skymidia.com.br" 
                                    className="flex items-center justify-center gap-2 text-sm hover:text-brand-hover transition-colors text-brand-text"
                                >
                                    <span className="opacity-40 uppercase tracking-widest text-[10px]">E-mail:</span>
                                    <span className="font-medium">rodrigo@skymidia.com.br</span>
                                </a>
                            </div>
                        </div>

                        <div className="bg-brand-dark/5 p-8 rounded-sm border border-brand-hover/10 flex flex-col items-center text-center space-y-4">
                            <span className="text-[10px] uppercase tracking-[0.2em] text-brand-hover font-bold">Comercial - Arquitetura</span>
                            <h4 className="text-xl font-serif text-brand-text">FERNANDO RIBALDO</h4>
                            <div className="space-y-2 pt-4 w-full">
                                <a 
                                    href="https://wa.me/5531985869051" 
                                    target="_blank" 
                                    rel="noopener noreferrer"
                                    className="flex items-center justify-center gap-2 text-sm hover:text-brand-hover transition-colors text-brand-text"
                                >
                                    <span className="opacity-40 uppercase tracking-widest text-[10px]">WhatsApp:</span>
                                    <span className="font-medium">+55 (31) 9 8586.9051</span>
                                </a>
                                <a 
                                    href="mailto:fernando@skymidia.com.br" 
                                    className="flex items-center justify-center gap-2 text-sm hover:text-brand-hover transition-colors text-brand-text"
                                >
                                    <span className="opacity-40 uppercase tracking-widest text-[10px]">E-mail:</span>
                                    <span className="font-medium">fernando@skymidia.com.br</span>
                                </a>
                            </div>
                        </div>
                    </div>
                </div>

                <ContactForm />
            </div>
        )
    }
];

export const BRAND_NAME = 'SKYMÍDIA';
export const PRIMARY_COLOR = '#727377'; 
export const ACCENT_COLOR = '#15F0DB'; 
export const TEXT_COLOR = '#E6E7E9';