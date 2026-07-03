# 💚 Planilha Financeira Pessoal

Planilha financeira completa, moderna e interativa, com visual profissional em **preto, verde e branco** (modo escuro elegante). Compatível com **Excel** e **Google Planilhas**.

📥 **Arquivo pronto para uso:** [`Planilha_Financeira_Pessoal.xlsx`](Planilha_Financeira_Pessoal.xlsx)

## 🌐 Versão site (novo!)

O mesmo controle financeiro em formato de **site interativo**: [`index.html`](index.html) — basta abrir o arquivo em qualquer navegador (funciona offline, sem instalar nada).

- Todos os módulos da planilha: dashboard, receitas, despesas, contas, cartões, metas, planejamento, relatórios e alertas.
- Gráficos animados (rosca, barras e linha) com tooltips.
- Lançamento de compras parceladas automático (divide em N meses).
- Filtros por mês, ano, categoria, conta, cartão e status.
- Dados salvos no navegador (localStorage) + exportação/importação de backup JSON e CSV.
- Para publicar na internet: ative o **GitHub Pages** nas configurações do repositório (Settings ▸ Pages ▸ Deploy from branch).

## 🗂️ Abas

| Aba | O que faz |
|---|---|
| 📊 **Dashboard** | Saldo atual, receitas, despesas e economia do mês, % de gastos, banner de situação financeira com cores automáticas, mini-alertas e 3 gráficos (pizza, barras e linha da evolução do patrimônio). Seletor de **mês/ano** que recalcula tudo. |
| 💰 **Receitas** | Data, descrição, categoria, valor, forma de recebimento, conta e status (Recebido/Pendente) — com listas suspensas e filtros. |
| 💸 **Despesas** | Data, descrição, categoria, valor, forma de pagamento, cartão, parcelas, vencimento e pago (Sim/Não). Vencidas ficam vermelhas e as próximas de vencer, amarelas. |
| 🏦 **Contas** | Saldo por banco com entradas, saídas e transferências entre contas — saldo atualizado automaticamente. |
| 💳 **Cartões** | Limite, fatura atual (soma automática das compras no crédito não pagas), limite disponível, % de uso com barra, melhor dia de compra e controle de parcelamentos. |
| 🎯 **Metas** | Valor desejado, guardado, restante, barra de progresso, aporte mensal e **previsão automática de conclusão**. |
| 📅 **Planejamento** | Orçamento por categoria × valor gasto no mês selecionado, com indicador 🟢/🟡/🔴 automático. |
| 📈 **Relatórios** | Receitas e despesas por mês, gastos por categoria, comparativo anual e ranking das 10 maiores despesas. |
| 🔔 **Alertas** | Avisos automáticos: orçamento estourado, contas a vencer em 7 dias, contas vencidas e metas atingidas. |
| ⚙️ **Config** | Listas editáveis (categorias, contas, cartões, formas de pagamento) usadas nas validações. |

## 🚀 Como usar

1. Abra o arquivo no Excel ou faça upload no Google Drive e abra com o Google Planilhas.
2. Lance suas receitas e despesas nas abas 💰 e 💸 usando os menus suspensos.
3. Escolha o **mês** e o **ano** no topo do 📊 Dashboard — todos os cálculos, gráficos e alertas se atualizam sozinhos.
4. Cadastre seus bancos em 🏦 Contas, cartões em 💳 Cartões, orçamentos em 📅 Planejamento e metas em 🎯 Metas.
5. A planilha vem com **dados de exemplo** (jan–jul/2026) para você ver tudo funcionando — basta substituí-los pelos seus.

## 🔒 Proteção

As células de fórmula estão protegidas **sem senha** para evitar apagamentos acidentais; apenas os campos de entrada ficam liberados. Para desproteger: *Revisão ▸ Desproteger Planilha*.

## 🛠️ Regenerar a planilha

O arquivo é gerado pelo script Python deste repositório:

```bash
pip install openpyxl
python3 gerar_planilha.py
```
