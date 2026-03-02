import React, { useState } from 'react';

const ContactForm: React.FC = () => {
  const [formData, setFormData] = useState({
    nome: '',
    email: '',
    telefone: '',
    mensagem: ''
  });
  const [status, setStatus] = useState<'idle' | 'sending' | 'success' | 'error'>('idle');
  const [errors, setErrors] = useState<Record<string, string>>({});

  const validate = () => {
    const newErrors: Record<string, string> = {};
    if (!formData.nome.trim()) newErrors.nome = 'Nome é obrigatório';
    if (!formData.email.trim()) {
      newErrors.email = 'E-mail é obrigatório';
    } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(formData.email)) {
      newErrors.email = 'E-mail inválido';
    }
    if (!formData.mensagem.trim()) newErrors.mensagem = 'Mensagem é obrigatória';
    
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!validate()) return;

    setStatus('sending');
    
    try {
      await new Promise(resolve => setTimeout(resolve, 1500));
      setStatus('success');
      setFormData({ nome: '', email: '', telefone: '', mensagem: '' });
    } catch (error) {
      setStatus('error');
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
    if (errors[name]) {
      setErrors(prev => {
        const newErrors = { ...prev };
        delete newErrors[name];
        return newErrors;
      });
    }
  };

  if (status === 'success') {
    return (
      <div className="bg-brand-hover/10 border border-brand-hover p-8 text-center rounded-sm animate-fade-in">
        <h4 className="text-2xl font-serif text-brand-hover mb-4">Mensagem Enviada!</h4>
        <p className="text-brand-text mb-6">Obrigado pelo contato. Nossa equipe retornará em breve.</p>
        <button 
          onClick={() => setStatus('idle')}
          className="text-xs uppercase tracking-widest font-bold text-brand-hover hover:underline"
        >
          Enviar outra mensagem
        </button>
      </div>
    );
  }

  return (
    <form onSubmit={handleSubmit} className="w-full max-w-4xl mx-auto space-y-6 pt-12 border-t border-brand-hover/10">
      <div className="text-center mb-8">
        <h3 className="text-xl font-serif uppercase tracking-widest text-brand-text">Envie um E-mail</h3>
      </div>

      <div className="grid md:grid-cols-3 gap-6">
        <div className="space-y-1">
          <label className="text-[10px] uppercase tracking-widest text-brand-text ml-1">Nome *</label>
          <input 
            type="text"
            name="nome"
            value={formData.nome}
            onChange={handleChange}
            className={`w-full bg-brand-dark/10 border ${errors.nome ? 'border-red-500' : 'border-brand-hover/20'} focus:border-brand-hover px-4 py-3 text-sm outline-none transition-colors text-brand-text`}
            placeholder="Seu nome completo"
          />
          {errors.nome && <span className="text-[10px] text-red-500 ml-1">{errors.nome}</span>}
        </div>

        <div className="space-y-1">
          <label className="text-[10px] uppercase tracking-widest text-brand-text ml-1">E-mail *</label>
          <input 
            type="email"
            name="email"
            value={formData.email}
            onChange={handleChange}
            className={`w-full bg-brand-dark/10 border ${errors.email ? 'border-red-500' : 'border-brand-hover/20'} focus:border-brand-hover px-4 py-3 text-sm outline-none transition-colors text-brand-text`}
            placeholder="seu@email.com"
          />
          {errors.email && <span className="text-[10px] text-red-500 ml-1">{errors.email}</span>}
        </div>

        <div className="space-y-1">
          <label className="text-[10px] uppercase tracking-widest text-brand-text ml-1">Telefone (Opcional)</label>
          <input 
            type="text"
            name="telefone"
            value={formData.telefone}
            onChange={handleChange}
            className="w-full bg-brand-dark/10 border border-brand-hover/20 focus:border-brand-hover px-4 py-3 text-sm outline-none transition-colors text-brand-text"
            placeholder="(00) 00000-0000"
          />
        </div>
      </div>

      <div className="space-y-1">
        <label className="text-[10px] uppercase tracking-widest text-brand-text ml-1">Mensagem *</label>
        <textarea 
          name="mensagem"
          value={formData.mensagem}
          onChange={handleChange}
          rows={6}
          className={`w-full bg-brand-dark/10 border ${errors.mensagem ? 'border-red-500' : 'border-brand-hover/20'} focus:border-brand-hover px-4 py-4 text-sm outline-none transition-colors text-brand-text resize-none`}
          placeholder="Como podemos ajudar no seu projeto?"
        />
        {errors.mensagem && <span className="text-[10px] text-red-500 ml-1">{errors.mensagem}</span>}
      </div>

      <div className="flex flex-col items-center gap-4">
        <button 
          type="submit"
          disabled={status === 'sending'}
          className="px-12 py-4 bg-brand-hover text-brand-bg font-bold uppercase tracking-widest hover:bg-brand-hover/80 transition-all duration-300 rounded-sm disabled:opacity-50"
        >
          {status === 'sending' ? 'Enviando...' : 'Enviar Mensagem'}
        </button>
        {status === 'error' && <p className="text-xs text-red-500">Ocorreu um erro. Tente novamente.</p>}
        {Object.keys(errors).length > 0 && <p className="text-xs text-red-500">Por favor, preencha todos os campos obrigatórios corretamente.</p>}
      </div>
    </form>
  );
};

export default ContactForm;
