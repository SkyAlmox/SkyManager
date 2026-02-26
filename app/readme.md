Aqui está um **diagrama conceitual em texto estruturado** para o seu aplicativo, considerando os módulos, conexões e perfis de usuário que você definiu:

---

## 📊 Diagrama Conceitual do Sistema de Almoxarifado

### 1. Perfis de Usuário
- **Administrador**
  - Acesso total a todos os módulos e relatórios.
  - Gerencia permissões e configurações.

- **Almoxarifado**
  - Acesso a: Produtos, Patrimônios, Fornecedores, Ordens de Produção.
  - Responsável por movimentações de estoque e acompanhamento de produção.

- **Vendas**
  - Acesso a: Clientes, Produtos, Ordens de Produção (apenas consulta).
  - Criação de pedidos e integração com produção.

- **Instalação**
  - Acesso a: Patrimônios e Ordens de Produção (execução e finalização).
  - Consulta a relatórios específicos de instalação.

---

### 2. Módulos Principais
- **Produtos**
  - Cadastro, estoque, movimentações.
  - Relacionado a Ordens de Produção e Fornecedores.

- **Clientes**
  - Cadastro e histórico de pedidos.
  - Relacionado a Vendas e Ordens de Produção.

- **Fornecedores**
  - Cadastro e histórico de fornecimento.
  - Relacionado a Produtos e Almoxarifado.

- **Patrimônios**
  - Controle de bens e equipamentos.
  - Relacionado a Instalação e manutenção.

- **Ordens de Produção**
  - Fluxo principal do Kanban.
  - Relacionado a Produtos, Clientes e Almoxarifado.

---

### 3. Dashboard
- **Filtros dinâmicos**: qualquer informação salva pode ser exibida.
- **Relatórios em tempo real**:
  - Estoque atual.
  - Ordens em andamento.
  - Movimentações recentes.
  - Alertas (estoque mínimo, ordens atrasadas).

---

### 4. Kanban de Produção
- Colunas fixas:
  1. **Ordens de Produção** (entrada de pedidos).
  2. **Separação de Material** (estoque → produção).
  3. **Produção** (execução).
  4. **Finalização** (conclusão da ordem).
  5. **Revisão** (quando necessário).

- Cada card no Kanban é uma **Ordem de Produção**, puxando dados de Produtos, Clientes e Almoxarifado.

---

### 5. Melhorias Futuras
- **Notificações**: alertas automáticos.
- **Busca inteligente**: localizar produtos/ordens rapidamente.
- **Exportação de relatórios**: PDF/Excel.
- **Histórico de movimentações**: rastreio completo de entradas e saídas.

---

<!-- prompt -->

Objetivo:
Criar um aplicativo web e mobile para gerenciamento de almoxarifado, com módulos interligados e funcionalidades de dashboard e kanban.
Stack Tecnológica:
- Frontend: React Native (para web e mobile, com componentes reutilizáveis).
- Backend: Node.js (API REST/GraphQL).
- Banco de Dados: PostgreSQL (relacional, para consistência e escalabilidade).
- Tempo real: WebSockets ou Firebase para atualização instantânea.
Perfis de Usuário:
- Administrador (acesso total).
- Almoxarifado (estoque, fornecedores, ordens).
- Vendas (clientes, pedidos, consulta ordens).
- Instalação (patrimônios, execução/finalização ordens).
Módulos:
- Produtos (cadastro, estoque, movimentações).
- Clientes (cadastro, histórico de pedidos).
- Fornecedores (cadastro, histórico de fornecimento).
- Patrimônios (controle de bens/equipamentos).
- Ordens de Produção (fluxo principal do kanban).
Dashboard:
- Relatórios em tempo real com filtros dinâmicos.
- Métricas: estoque atual, ordens em andamento, movimentações recentes, alertas.
Kanban de Produção:
- Colunas fixas:
- Ordens de Produção
- Separação de Material
- Produção
- Finalização
- Revisão
Funcionalidades Futuras:
- Notificações (estoque mínimo, ordens atrasadas).
- Busca inteligente (produtos, ordens).
- Exportação de relatórios (PDF/Excel).
- Histórico de movimentações (entradas/saídas).
