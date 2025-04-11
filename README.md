# 🧾 Relatório de Espelho de Pedido de Compras

Este projeto tem como objetivo gerar um relatório formatado de **espelho de pedido de compras**, utilizando recursos gráficos do `oImpFRMC` no **TOTVS Protheus (ADVPL/TL++)**.

---

## 📌 Funcionalidades

- Impressão de cabeçalhos com colunas fixas.
- Impressão de linhas de itens com dados organizados em colunas.
- Impressão de informações adicionais (descrição, fornecedor, observações etc.).
- Impressão da área de totais com colunas distintas para tributos e valores finais.

---

## 🧱 Estrutura das Funções

### `fImprimeItens()`

Responsável por:

- Gerar o cabeçalho das colunas dos itens.
- Iterar sobre os itens e desenhar:
  - Colunas com: `#`, `Item`, `NCM`, `CFOP`, `Entrega`, `Qtd`, `Valor Unit.`, `Desc.`, `Total`.
  - Informações detalhadas de cada item como **descrição**, **fornecedor**, **observação**, etc.

#### Parâmetros:
- `aItens`: array de itens com os dados.
- `nRow_`, `nCol_`, `nCol_R`: controle de posição da impressão no relatório.

---

### `fImprimeTotais()`

Responsável por:

- Exibir a tabela de totais e tributos ao final do relatório.
- Colunas para: `Valor Total`, `Frete`, `Seguro`, `IPI`, `PIS`, `COFINS`, `ICMS`, entre outros.

#### Parâmetros:
- `jTotais`: objeto JSON contendo os valores dos totais e impostos.
- `nRow_`, `nCol_`, `nCol_R`: controle de posição.

---

## 🧮 Exemplo de Item Renderizado

1 MPC0112600021 33029019 2101 14/04/25 7.00 285.15 0.00 1996.05

---

## ✨ Observações Técnicas

- A função `SayAlign()` é usada com:
  - **Posição X** e **largura** definidas dinamicamente pelas colunas.
  - Parâmetro de espaçamento vertical fixo (ex: `15`).
  - Alinhamento `ALIGN_H_LEFT` ou `ALIGN_H_RIGHT` conforme a necessidade.
- `nMargRel_` e `nColTot_` definem a margem inicial e o total da largura do relatório.

---

## 📎 Dependências

- Ambiente Protheus com suporte a relatórios via `oImpFRMC`.
- Fonte `oFont10` e `oFont12` devidamente carregadas.

---

## 👨‍💻 Autor

Desenvolvido por **[TigoP]**, contador com mais de 15 anos de experiência e apaixonado por tecnologia aplicada à gestão empresarial.

---
