# Safe Harbor Financeiro · App de Gestão Financeira de Projetos

Frontend para o controle financeiro dos projetos do Safe Harbor GPOD. O app lê e grava **direto nas planilhas** `Database.xlsx` e `Config.xlsx` de uma pasta sincronizada pelo OneDrive/SharePoint. Não há servidor nem banco de dados: tudo roda no navegador e nenhum dado financeiro sai da máquina de quem usa.

> ⚠️ **Este repositório é público.** Ele contém só o código do app. Planilhas, backups e capturas de tela com dados reais **nunca** devem ser enviados (o `.gitignore` já bloqueia `*.xlsx`, `*.png` e as pastas de teste).

## Funcionalidades

| Área | O que faz |
|---|---|
| **Dashboard** | 4 visões prontas (executiva, rentabilidade, time e horas, contratos), 13 tipos de gráfico, filtros por período e por dimensão, ordenação, clique para filtrar e pontos de atenção automáticos. O layout é personalizável e pode ser exportado e importado em `.json`. |
| **Projetos e contratos** | Criar, editar e excluir projetos e contratos (aba `2. Propostas`). Um cliente novo ganha sua aba `BL_` a partir do `BL_Template`. |
| **Time e alocação** | Alocações por mês com papel e senioridade filtrados por Innovation Studio e ratecard e custo preenchidos automaticamente pela planilha **DRP** (aba `Custo Referência`), conforme o ano do mês. |
| **Ajuste de horas** | Ajuste mensal das horas executadas, agrupado por cliente › projeto › profissional, com motivo gravado na coluna Obs. |
| **Uso em equipe** | Trava de edição (uma pessoa edita por vez), gravação automática, backups com limpeza após 3 dias, detecção de alterações externas e recuperação de itens perdidos. |

## Como usar

1. Sincronize a pasta **Financeiro** do SharePoint no computador (OneDrive).
2. Baixe [`dist/SafeHarbor_Financeiro.html`](dist/SafeHarbor_Financeiro.html) e abra no **Chrome ou no Edge para computador**.
3. Clique em **Conectar pasta Financeiro** e escolha a pasta. O app abre em modo leitura; clique em **Editar** para fazer alterações.

A pasta precisa conter `Database*.xlsx` e `Config*.xlsx`, e deve ter também a planilha **DRP** da gestão (`DRP*.xlsx`, ex.: `DRP AIR BR V20260925.xlsx`). Toda vez que o app abre, ele confere o DRP, mostra a versão em uso e lembra de mantê-la atualizada; com mais de um DRP na pasta, usa o de versão mais recente (`Vaaaammdd` no nome). O app cria nela a pasta `_backups_safeharbor/` e o arquivo `_safeharbor.lock`.

## Estrutura

```
src/
  template.html   Interface (HTML, CSS e JS das telas, dashboard e trava de edição)
  core.js         Motor da planilha: leitura e gravação do XLSX, fórmulas, cálculos e Config
assets/
  airr-logo.svg   Logo AI/R (já embutido no template)
dist/             Arquivos gerados pelo build (não editar à mão)
  index.html                  Versão com JSZip via CDN
  SafeHarbor_Financeiro.html  Versão autossuficiente (JSZip embutido), para usar localmente
tests/            Testes de ponta a ponta (Playwright)
build.py          Monta dist/ a partir de src/
```

## Desenvolvimento

```bash
npm install          # instala o JSZip
python3 build.py     # gera dist/index.html e dist/SafeHarbor_Financeiro.html
npm run serve        # serve dist/ em http://localhost:8765
```

Edite apenas `src/` e rode o build. O marcador `/*__CORE__*/` em `src/template.html` é substituído pelo conteúdo de `src/core.js`.

### Regras de negócio principais

- **DRP é a fonte de Innovation Studios, perfis, custos e ratecards** (aba `Custo Referência`, somente leitura; colunas lidas pelo nome, anos 2026 em diante). O custo padrão é o *PC para Ratecard (com overhead)*; o *Custo Total (Sem Overhead)* fica disponível no formulário. Sem DRP, o app usa a aba `Perfis` do Config como alternativa.
- **Config.xlsx é a fonte única** de impostos, dias úteis e listas. A aba `1. Config` da Database é um espelho técnico, atualizado pelo app para que as fórmulas das abas `BL_` continuem funcionando no Excel.
- **Cálculo por linha:** bruto = horas × R$/h × fator · impostos = bruto × (PIS + COFINS + INSS + ISSQN), zero se Isento · líquido = bruto − impostos · custo = horas × custo/h · resultado = líquido − custo · MB = resultado ÷ líquido.
- **Tributação:** vem do campo *Funding* do contrato (Sim = Isento). Projeto sem contrato é Tributado.
- **Faixas de margem:** abaixo de 20% vermelho · de 20% a 28,9% amarelo · 29% ou mais verde (`MB_RED` e `MB_GREEN` em `src/template.html`).

### Testes

Os testes usam Playwright (Python) e precisam de cópias das planilhas em `tests/fixtures/` (veja o [README da pasta](tests/fixtures/README.md); esses arquivos não são versionados).

```bash
pip install playwright && playwright install chromium
npm run serve &                      # em outro terminal
python3 tests/fin_test.py            # fluxo completo com a pasta Financeiro
python3 tests/lock_test.py           # trava de edição entre dois usuários
python3 tests/dash_test.py           # dashboard configurável
python3 tests/pf_test.py             # perfis por iStudio e ratecard automático
python3 tests/mb_test.py             # faixas de margem
python3 tests/prune_test.py          # limpeza de backups
```

Para usar o Chrome instalado no lugar do Chromium do Playwright, defina `CHROME_PATH`. As saídas vão para `tests/out/` (ignorada pelo Git).

## Requisitos e limitações

- Leitura e gravação na pasta usam a *File System Access API*: **Chrome ou Edge para computador**. Não funciona em Safari, Firefox ou celular.
- A trava de edição depende da sincronização do OneDrive, que leva alguns segundos.
- O nome informado na trava serve só para exibição; não há autenticação.
