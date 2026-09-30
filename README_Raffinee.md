# Raffinée Analytics

Painel comercial da Raffinée: objetivo x realizado por cliente, vendedor e produto, clientes que compraram e que não compraram, kg e R$ acumulados por produto, apresentações e central de arquivos.

## Como o sistema funciona

```
index.html (GitHub Pages)  ──lê/grava──►  Apps Script (Web App)  ──►  Planilha "Raffinee Analytics"
        ▲                                        │                         (abas = base de dados)
        │ push                                   └──►  Google Drive (PDFs)
     GitHub  ──(Action, só se mudar apps-script/)──►  nova versão da mesma implantação
```

- **Planilha**: guarda os dados. Uma linha por registro, nunca uma coluna por mês.
- **Apps Script** (`apps-script/Code.gs`): lê e grava na planilha, sobe PDFs no Drive. Protegido por uma chave (aba CONFIG).
- **index.html**: toda a inteligência dos KPIs roda no navegador, em cima das tabelas que o Apps Script entrega.
- **GitHub**: fonte oficial do código. Cada push atualiza o index no GitHub Pages; quem estiver com o sistema aberto recebe o aviso "Nova versão disponível" (lido do `version.json`).

## Abas da planilha

| Aba | Conteúdo |
|---|---|
| BASE_VENDAS | ID_VENDA, DATA, ANO, MES, PEDIDO, COD_CLIENTE, CLIENTE, VENDEDOR, COD_PRODUTO, PRODUTO, CANAL, QTD, KG, VALOR, ORIGEM, IMPORTADO_EM |
| CLIENTES | COD_CLIENTE, CLIENTE, VENDEDOR, CANAL, CIDADE, UF, STATUS |
| PRODUTOS | COD_PRODUTO, PRODUTO, CATEGORIA, LINHA, PESO_KG, STATUS |
| VENDEDORES | COD_VENDEDOR, VENDEDOR, STATUS |
| METAS | ANO, MES, TIPO (EMPRESA / VENDEDOR / CLIENTE / PRODUTO), CHAVE (nome), META_VALOR, META_KG |
| ARQUIVOS | ID, TIPO, NOME, PERIODO, DATA_GERACAO, LINK_DRIVE, OBS |
| CONFIG | CHAVE_ACESSO, EMPRESA, PASTA_DRIVE_ID, DIAS_INATIVO, FATOR_ATRASO |
| LOG | Registro das importações, metas e cadastros |

As metas ficam numa aba única com a coluna TIPO, em vez de três abas separadas. Sem meta de EMPRESA, o sistema soma as metas dos vendedores.

## Implantação (uma vez)

1. **Repositório**: crie o repositório `Raffinee` no GitHub e suba estes arquivos.
2. **GitHub Pages**: Settings > Pages > Deploy from a branch > `main` / root. O link fica `https://<usuario>.github.io/Raffinee/`.
3. **Planilha**: crie uma planilha nova chamada `Raffinee Analytics` (separada das planilhas do grupo).
4. **Apps Script**: na planilha, Extensões > Apps Script. Cole o `Code.gs`, e em Configurações do projeto marque "Mostrar appsscript.json" e cole o `appsscript.json`.
5. Rode a função `setup()` uma vez e autorize. Ela cria todas as abas e gera a `CHAVE_ACESSO` na aba CONFIG.
6. Implantar > Nova implantação > App da Web. Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa**. Copie a URL `/exec`.
7. Abra o index uma vez com a conexão no link:
   `https://<usuario>.github.io/Raffinee/?api=URL_EXEC&k=CHAVE_ACESSO`
   O navegador guarda a conexão e o link some da barra. Também dá para colar em Configurações.
8. (Opcional) Crie uma pasta no Drive para os PDFs e coloque o ID em `PASTA_DRIVE_ID` na aba CONFIG.

Sem conexão configurada, o sistema abre com **dados de exemplo** (selo amarelo no título). Serve para apresentar ao João antes de ter a base real.

## Atualizações

- **Index**: altere, suba o `APP_VERSION` no `index.html` e o `versao` no `version.json`, acrescente a linha em `HISTORICO` e faça o push. Em até 5 minutos aparece "Nova versão disponível".
- **Backend**: pelo editor, sempre Gerenciar implantações > Editar > **Nova versão** na mesma implantação (a URL não muda).
  Ou automático pelo workflow `.github/workflows/deploy-apps-script.yml`, que faz exatamente isso quando algo em `apps-script/` muda. Segredos: `CLASPRC_JSON`, `SCRIPT_ID`, `DEPLOYMENT_ID`.

## Regras dos indicadores

- **Comparação com o ano anterior**: sempre o mesmo período (jan até o mês selecionado) e, no mês corrente, até o mesmo dia.
- **Meta acumulada**: soma das metas mensais; o mês corrente entra proporcional aos dias corridos.
- **Nível da meta** segue os filtros: cliente > produto > vendedor > empresa. Com filtro só de canal não há meta.
- **Ritmo necessário**: (meta anual − realizado) ÷ meses restantes, comparado à média dos 3 últimos meses fechados.
- **Situação do cliente** (na data de referência):
  - Comprou no mês: tem compra no mês selecionado.
  - Inativo: mais de 120 dias sem comprar.
  - Em risco: passou 1,5× da própria frequência de recompra (sem histórico suficiente: mais de 45 dias).
  - Ainda não comprou: os demais.
- **Cobertura** = compraram no mês ÷ carteira sem os inativos.
- **Clientes ativos** = compraram nos últimos 90 dias.
- **Movimento do produto**: crescimento saudável (R$ e kg sobem), puxado por valor (R$ sobe, kg parado ou caindo), crescimento de volume (kg sobe mais que R$), volume sem preço, queda, novo.

## Importação

Lançamentos / Importações aceita CSV (`;` ou `,`) e Excel. Modelo em `modelos/modelo_vendas.csv`. O `modelos/modelo_metas.csv` mostra o formato para colar direto na aba METAS (as metas também podem ser digitadas em Planejamento / Metas). Linhas com o mesmo PEDIDO + COD_PRODUTO + DATA não são gravadas duas vezes. O modo "Substituir" apaga a BASE_VENDAS antes de gravar.
