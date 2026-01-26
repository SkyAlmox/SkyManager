# 📦 Sky Mídia - Ecossistema de Gestão de Almoxarifado Inteligente

Este projeto visa a digitalização e automação do controle de estoque e ativos da **Sky Mídia (Belo Horizonte)**. O sistema utiliza Python para conectar processos físicos (ferramentas e insumos) a uma base de dados em nuvem via Excel Online.

## 🚀 Objetivo
Migrar o controle manual do almoxarifado para um ecossistema digital que permita:
- Controle de insumos consumíveis (lonas, tintas, adesivos).
- Gestão de ativos retornáveis (ferramentas e equipamentos).
- Monitoramento de estoque em tempo real via celular.

---

## 🏗️ Arquitetura do Sistema

O projeto é baseado em uma arquitetura de baixo custo e alta eficiência (Lean Stack):

1.  **Interface (Frontend/Mobile):** Streamlit Web App (Acessível via Navegador).
2.  **Lógica/Cérebro (Backend):** Python 3.12+.
3.  **Banco de Dados:** Excel Online (Office 365) integrado via API/Pandas.
4.  **Hardware de Captura:** Câmeras de dispositivos móveis para leitura de QR Codes.



---

## 📋 Planejamento de Funcionalidades

### Fase 1: Gestão de Consumíveis (Insumos)
- [ ] Estruturação da Planilha Mestra (Excel Online).
- [ ] Script Python para geração em lote de etiquetas QR Code.
- [ ] Interface para leitura e baixa de material (Saída de Estoque).
- [ ] Alertas automáticos de "Estoque Baixo".

### Fase 2: Gestão de Ativos (Ferramentas)
- [ ] Implementação de sistema de Check-out e Check-in.
- [ ] Criação de "QR Codes de Parede" para itens pequenos/peças manuais.
- [ ] Registro de histórico: "Com quem está a ferramenta X?".
- [ ] Criação de Kits (ex: Mala de Instalação 01).

### Fase 3: Inteligência e Relatórios
- [ ] Dashboard de consumo mensal.
- [ ] Integração com IA para previsão de compras (Baseado em demanda histórica).

---

## 🛠️ Especificações Técnicas

### Modelagem de Dados (Campos Principais)
| Campo | Descrição | Exemplo |
| :--- | :--- | :--- |
| `ID_ITEM` | Chave única (Primary Key) | `LONA_440_BR` |
| `NOME` | Descrição do produto | Lona Brilho 440g |
| `TIPO` | Consumível ou Ativo | Consumível |
| `SALDO` | Quantidade atual | 50.5 |
| `STATUS` | Localização ou disponibilidade | Em Viagem |

### Tecnologias Utilizadas
- **Linguagem:** Python
- **Bibliotecas Principais:**
  - `streamlit`: Para a interface web.
  - `pandas`: Para manipulação de dados.
  - `qrcode`: Para geração das etiquetas.
  - `opencv-python` / `streamlit-qrcode-scanner`: Para leitura via câmera.

---

## 📈 Roadmap de Implementação

1. **Semana 1:** Validação da planilha e geração dos primeiros QR Codes de teste.
2. **Semana 2:** Desenvolvimento do MVP de leitura e conexão com Excel.
3. **Semana 3:** Teste em ambiente real (Almoxarifado Sky Mídia).
4. **Semana 4:** Apresentação de resultados e métricas para a diretoria.

---

**Desenvolvido por:** [Seu Nome]
**Localização:** Belo Horizonte, MG
**Formação:** Sistemas de Informação