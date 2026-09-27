# 📈 Simulação de Investimentos — Modelo Financeiro em Excel

Uma ferramenta analítica desenvolvida em Excel para projeção de património, simulação de aportes em **Fundos de Investimento Imobiliário (FIIs)** e recomendação automatizada do **perfil de risco do investidor**.

---

## 📋 Conteúdo

* [Descrição do Projeto](https://www.google.com/search?q=%23-descri%C3%A7%C3%A3o-do-projeto&utm_source=gemini)
* [Estrutura do Ficheiro](https://www.google.com/search?q=%23-estrutura-do-ficheiro&utm_source=gemini)
* [Funcionalidades e Fórmulas](https://www.google.com/search?q=%23-funcionalidades-e-f%C3%B3rmulas&utm_source=gemini)
* [Como Utilizar](https://www.google.com/search?q=%23-como-utilizar&utm_source=gemini)
* [Tecnologias / Funções Utilizadas](https://www.google.com/search?q=%23-tecnologias--fun%C3%A7%C3%B5es-utilizadas&utm_source=gemini)

---

## 🎯 Descrição do Projeto

O modelo tem como objetivo auxiliar na tomada de decisão financeira através de:

1. **Análise de Objetivos:** Cálculo da taxa de rentabilidade anual necessária para atingir determinado valor num prazo definido.
2. **Classificação de Perfil:** Sugestão automática do perfil de investidor (*Conservador*, *Moderado* ou *Agressivo*).
3. **Seleção Dinâmica de Portfólio:** Mapeamento de FIIs de acordo com o perfil selecionado e cálculo do rendimento médio esperado.
4. **Projeção Patrimonial:** Simulação do crescimento do património a longo prazo (2, 5, 10, 20 e 30 anos) considerando aportes periódicos.

---

## 📁 Estrutura do Ficheiro

O ficheiro encontra-se dividido em duas folhas de cálculo principais:

| Folha | Descrição |
| --- | --- |
| **`Página1`** | Painel principal com calculadoras de aportes, sugestão de perfil, projeção temporal e portfólio de FIIs ativo. |
| **`Apoio`** | Tabela relacional de suporte que armazena a lista de FIIs, taxas de rendimento por ativo e taxas de variação/risco associadas a cada perfil. |

---

## ⚙️ Funcionalidades e Fórmulas

### 1. Sugestão de Perfil do Investidor

A folha calcula a taxa de retorno anual necessária ($R$) para atingir o objetivo com base na fórmula de juros compostos:


$$\text{Rendimento Anual} = \left(\frac{\text{Valor Desejado}}{\text{Investimento Inicial}}\right)^{\frac{1}{\text{Anos}}} - 1$$

A partir desta taxa, o perfil é categorizado dinamicamente:

* **Conservador:** Rendimento $\le 9\%$
* **Moderado:** Rendimento entre $9\%$ e $15\%$
* **Agressivo:** Rendimento $\ge 15\%$

```excel
=IFS(C19<=0.09, "Conservador", C19>=0.15, "Agressivo", AND(C19>0.09, C19<0.15), "Moderado")

```

---

### 2. Portfólio Dinâmico de FIIs

Com base no perfil configurado na célula `C25`, a ferramenta procura automaticamente na folha `Apoio` a lista de ativos correspondentes e calcula o rendimento esperado de cada um através de procuras compostas (`VLOOKUP` / `PROCV`).

#### Rendimentos Médios por Perfil:

* **Conservador:** HCTR11 ($8\%$), XPIN11 ($10\%$), DEVA11 ($9\%$), TGAR11 ($9\%$) — **Média: $9,0\%$**
* **Moderado:** SPDE11 ($14\%$), VGRI11 ($11\%$), LIFE11 ($12\%$), BRFT11 ($12\%$) — **Média: $12,0\%$**
* **Agressivo:** VCRA11 ($15\%$), EGAF11 ($14\%$), AAZQ11 ($12\%$), EDGE11 ($15\%$) — **Média: $14,5\%$**

---

### 3. Projeção Patrimonial Futura

Utiliza a função de Valor Futuro (`FV` / `VF`) para projetar o acumulado ao longo de vários naos, considerando uma taxa anual média e aportes mensais/anuais:

```excel
=FV($C$28, E26, $C$26 * -1)

```

---

## 🚀 Como Utilizar

1. **Defina o Salário e Sugestão de Aporte:** Insira o salário bruto para calcular a sugestão de investimento mensal ($15\%$).
2. **Defina as Metas:** Preencha o **Valor Desejado**, **Investimento Inicial** e a quantidade de **Anos** para alcançar o objetivo.
3. **Consulte o Perfil Sugerido:** Verifique a categoria recomendada no painel principal.
4. **Selecione o Perfil Pretendido:** Ajuste a célula de **Perfil** (`C25`) para atualizar a carteira de FIIs e os retornos esperados.
5. **Analise as Projeções:** Acompanhe a evolução temporal do património na tabela de cenários futuros.

---

## 🛠️ Tecnologias / Funções Utilizadas

* **Microsoft Excel / Google Sheets**
* **Funções Lógicas:** `IFS` / `SE`, `AND` / `E`
* **Funções de Procura:** `VLOOKUP` / `PROCV`
* **Funções Estatísticas e Financeiras:** `MEDIAN` / `MEDIANA`, `FV` / `VF`
* **Modelos de Procura Condicional Multi-critério** (Combinação de Chaves Únicas)
