# Ferramenta-de-Controle-de-Investimentos-com-Excel
Utilizando o excel como apoio na tomada de decisões dentro de uma carteira de fundos imobiliários.
#  Ferramenta de Controle de Investimentos

Planilha em Excel para simulação de aportes mensais, projeção de patrimônio no longo prazo e sugestão de alocação de carteira em Fundos Imobiliários (FIIs), de acordo com o perfil de risco do investidor.


---

## ⚙️ Configurações gerais

No topo do simulador ficam os parâmetros de base usados pelos cálculos:

- **Salário**: renda mensal do investidor (célula nomeada `salario`).
- **Rentabilidade (A.M.)**: rentabilidade média mensal esperada da carteira (`rendimento_carteira`), usada para estimar dividendos.
- **Sugestão de aporte (mensal)**: calculada automaticamente como **30% do salário** (`=salario*30%`), servindo de referência para quanto investir todo mês.

---

## 💰 Investimento mensal

Área onde o usuário informa os dados da simulação e obtém os resultados:

- **Quanto investir por mês?** – valor do aporte mensal (`aporte_mensal`).
- **Por quanto tempo?** – prazo do investimento, em anos (`período`).
- **Taxa de rendimento mensal?** – taxa de retorno mensal esperada (`taxa`).
- **Patrimônio acumulado?** – calculado com a função financeira `FV` (Valor Futuro):
  `=FV(taxa, período*12, aporte_mensal*-1)`
- **Dividendos mensais?** – estimativa de renda passiva mensal gerada pelo patrimônio acumulado:
  `=montante*rendimento_carteira`

Ou seja, basta alterar o valor do aporte, o prazo e a taxa esperada para que o patrimônio final e os dividendos mensais projetados sejam recalculados automaticamente.

---

## 📈 Cenários de projeção

Tabela que projeta automaticamente o patrimônio acumulado e os dividendos mensais em diferentes horizontes de tempo, mantendo o aporte e a taxa definidos acima:

- 2 anos
- 5 anos
- 10 anos
- 20 anos
- 30 anos

Cada linha usa a mesma função `FV` variando apenas o número de meses, permitindo comparar rapidamente o efeito do tempo (juros compostos) sobre o patrimônio e a renda passiva.

Esses dados alimentam um **gráfico de pizza** com a distribuição sugerida da carteira por tipo de fundo (ver seção seguinte).

---

## 🧭 Perfil do investidor e alocação sugerida

- **Perfil do investidor**: campo com lista suspensa (dropdown) para escolher entre `Conservador`, `Moderado` e `Agressivo`.
- **Aporte mensal**: replica o valor definido em "Investimento mensal".

A partir do perfil escolhido, a planilha sugere automaticamente como distribuir o aporte mensal entre 6 tipos de Fundos Imobiliários:

| Tipo de FII |
|---|
| Papel |
| Tijolo |
| Híbridos |
| Fundos de Fundos (FOFs) |
| Desenvolvimento |
| Hotelarias |

Para cada tipo, a planilha calcula:
- **Percentual Sugerido** — buscado na aba `chave composta` via `VLOOKUP`, combinando perfil + tipo de fundo em uma chave composta (ex: `MODERADO-Tijolo`).
- **Valores** — o percentual sugerido aplicado sobre o aporte mensal (`=aporte_mensal*percentual`).

Um **gráfico de pizza** ilustra visualmente essa distribuição percentual entre os tipos de fundo.

### Exemplo de alocação por perfil

| Tipo de FII | Conservador | Moderado | Agressivo |
|---|---|---|---|
| Papel | 30% | 32% | 50% |
| Tijolo | 50% | 40% | 10% |
| Híbridos | 10% | 8% | 5% |
| Fundos de Fundos (FOFs) | 10% | 10% | 5% |
| Desenvolvimento | 0% | 5% | 20% |
| Hotelarias | 0% | 5% | 10% |

> Quanto mais conservador o perfil, maior a concentração em fundos de **Tijolo** (mais estáveis); quanto mais agressivo, maior a exposição a **Papel** e **Desenvolvimento** (maior risco/retorno).

---

## 🔑 Aba "chave composta"

Base de dados que sustenta a busca de percentuais sugeridos. Cada linha combina **Perfil + Tipo de Fundo** em uma chave única (coluna `CHAVE COMPOSTA`, ex.: `AGRESSIVO-Papel`), associada ao percentual sugerido de alocação. É essa tabela que o `VLOOKUP` da aba principal consulta.

---

##  Intervalos nomeados (Named Ranges)

Para deixar as fórmulas mais legíveis, o workbook usa os seguintes nomes:

| Nome | Célula | Significado |
|---|---|---|
| `salario` | F15 | Salário mensal |
| `rendimento_carteira` | F16 | Rentabilidade mensal esperada |
| `sugestao_de_aporte` | F17 | Sugestão de aporte (30% do salário) |
| `aporte_mensal` | D20 | Valor do aporte mensal informado |
| `período` | D21 | Prazo do investimento (anos) |
| `taxa` | D22 | Taxa de rendimento mensal |
| `montante` | D23 | Patrimônio acumulado projetado |

---

##  Como usar

1. Preencha **Salário** e **Rentabilidade (A.M.)** em "Configurações".
2. Em "Investimento mensal", informe quanto pretende investir por mês, por quanto tempo e a taxa esperada.
3. Confira o **Patrimônio acumulado** e os **Dividendos mensais** projetados.
4. Veja a tabela de **Cenários** para comparar o crescimento do patrimônio em 2, 5, 10, 20 e 30 anos.
5. Selecione seu **Perfil do investidor** (Conservador, Moderado ou Agressivo) no menu suspenso.
6. Confira a **alocação sugerida** entre os tipos de Fundos Imobiliários e o gráfico de pizza correspondente.

---

## 🛠 Tecnologia

- Microsoft Excel (`.xlsx`)
- Funções utilizadas: `FV`, `VLOOKUP`, intervalos nomeados, validação de dados (lista suspensa) e gráfico de pizza nativo do Excel.

---

## ⚠️ Aviso

Esta planilha é uma ferramenta de **simulação educacional** e não constitui recomendação de investimento. Rentabilidades passadas ou projetadas não garantem resultados futuros. Consulte um profissional certificado antes de tomar decisões de investimento.
