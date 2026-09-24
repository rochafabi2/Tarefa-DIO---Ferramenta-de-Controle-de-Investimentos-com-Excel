# Tarefa-DIO---Ferramenta-de-Controle-de-Investimentos-com-Excel
Foi criado um arquivo Excel para calcular investimentos, solicitado pelo prof. Felipão para finalização do módulo do curso Santander - Excel com IA e Claude

# Ferramenta de Controle de Investimentos com Excel & IA

Uma solução prática e automatizada desenvolvida em Excel, criada durante o bootcamp da **DIO (Digital Innovation One)**. O objetivo é facilitar o acompanhamento de aportes, distribuição de ativos, cálculo de rendimentos e organização do patrimônio financeiro.

---

## Sobre o Projeto


O projeto foi construído com a orientação do professor, onde exploramos a capacidade de geração de fórmulas, formatações, nomear intervalos e criar uniformidade visual. Foi usado para simular investimentos com três tipos de investidores (Moderado, Agressivo e Conservador), mas com o que foi aprendido posso gerar diversos tipos de simuladores.

##  Funcionalidades

- **Tabela Configurações**:
  - Definição de salário/renda mensal.
  - Rendinemnto da carteira
  - Regra de alocação automática (ex.: meta de 30% da renda para investimentos).
  **Tabela Investimento Mensal**:
    - Quanto investir por mês? (preenchido pelo usuário)
    - Por quantos anos?(preenchido pelo usuário)
    - Patrimônio acumulado? (calculado pelo Excel)
    - Dividendos Mensais? (rendimento, calculado pelo Excel)
- **Tabela Cenários**:
  - o usuário poderá ver, já calculado, quanto o aporte renderá em 05 cenários diferentes (02, 5, 10, 20 e 30 anos)
- **Tabela Perfil**:
  - Nessa tabela o usuário escolhe o seu perfil de investidor (Moderado, Avançado ou Agressivo) e o Excel sugere os tipos de investimentos, o percentual sugerido e quanto investir em cada categoria

- Há também um gráfico de pizza com o percentual sugerido de acordo com o perfil escolhido

##  Tecnologias e Ferramentas

- **Microsoft Excel**: Modelagem, formatação condicional e fórmulas financeiras/estatísticas.
Foi criada uma tabela de apoio Planilha 2.  Ela armazena as combinações de perfil de risco e os percentuais ideais para cada categoria de investimento.

## Fórmulas utilizadas
**Tabela de Configurações, aba App**
- Multiplicação simples de porcentagem (*)
- Fórmula: = Salário * 30%1
- Função: Multiplica o valor total informado da renda/salário por 30% (0,30) para calcular automaticamente a sugestão ideal de aporte mensal.

**Tabela de Investimento Mensal, aba App**
- Fórmula: = VF(Taxa; Anos * 12; -AporteMensal)
- Função: Simula o acúmulo de patrimônio ao longo do tempo (juros compostos), somando os depósitos mensais regulares com os rendimentos acumulados no período especificado.
- Fórmula: = Patrimônio Acumulado * Rendimento da Carteira1
- Função: Multiplica o valor total do patrimônio acumulado pela taxa mensal estimada para calcular a renda passiva mensal em Reais.

**Tabela Cenários, aba app**
- Fórmula: Aplicação da função VF adaptada aos prazos de 24, 60, 120, 240 e 360 meses (referentes a 2, 5, 10, 20 e 30 anos).
- Função: Projeta o patrimônio total acumulado para múltiplos horizontes temporais de longo prazo.
- Fórmula: Multiplicação do patrimônio projetado de cada período pela taxa de retorno mensal estimada.
- Função: Estima os dividendos mensais em Reais gerados pelo patrimônio acumulado em cada cenário de tempo.

**Tabela Perfil, aba app**
- Fórmula: = PROCV(Perfil & "-" & TipoFII; MatrizPlanilha2; ColunaPorcentagem; FALSO)
- Função: Realiza uma busca combinada na matriz de referência (Aba Planilha2) para identificar o percentual exato recomendado para cada segmento de FII de acordo com o perfil selecionado (Conservador, Moderado ou Agressivo).
- Fórmula: = Valor a ser investido por mês * Percentual Sugerido
- Função: Converte o percentual sugerido no valor exato em Reais a ser aplicado em cada categoria de fundo imobiliário.


## Como Utilizar

1. **Clone ou baixe o repositório**:
   ```bash

Abra o arquivo Ferramenta de Controle de Investimentos com Excel.xlsx no Microsoft Excel ou importe-o no Google Planilhas.
Acesse a aba de configurações e informe seus parâmetros (renda mensal, percentual de aporte desejado, etc.).
Comece a cadastrar suas posições e acompanhe as atualizações automáticas dos totais e rendimentos.
