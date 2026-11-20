Identidade visual da **AI/R — Agentic AI Engineering**, extraída do template oficial de apresentações 2026. É uma marca monocromática e técnica: títulos enormes e leves em Neue Haas Grotesk Display, rótulos e números em Roboto Mono, uma rampa de cinzas sem nenhuma cor de destaque, formas retas com um único canto cortado e a seta ↗ como único ícone.

Todas as medidas estão em px num canvas de slide **1920×1080** (o dobro dos pontos do template, que tem 960×540pt). Para outro tamanho de canvas, escale proporcionalmente.

## Regras inegociáveis do template

O próprio guia do template ("Template Guide") estabelece regras que valem para qualquer peça gerada com este sistema:

- **Não altere o layout dos slides.** O logo fica sempre no canto superior direito, no mesmo tamanho; não o redimensione nem o mova.
- **Não insira imagens sem aprovação** do time de Internal Comms (comms@aircompany.ai) ou de Marketing (marketing@aircompany.ai). Use só as imagens do grupo `Backgrounds`. Não use vetores, GIFs, vídeos ou elementos que não façam parte da identidade.
- **Não mude cores de fundo nem tamanhos de fonte** de títulos e descrições. Respeite as guias de alinhamento.
- **Mantenha os fundos padrão das seções**; não troque as imagens de fundo.
- **Use só as fontes aprovadas:** Neue Haas Grotesk Display Pro e Roboto Mono.
- **Parcerias:** em apresentações de parceria use o logo `airr-compass-white.png` / `airr-compass-black.png` (AI/R + compass.uol, chamado no guia de "AI/R + Powerhouse") na capa e no cabeçalho. **Nunca misture slides com o logo AI/R e com o logo de parceria na mesma apresentação.**

## Conteúdo e tom

- O template escreve em **inglês**, em frases curtas e assertivas, no imperativo para quem monta o slide: "Insert a strong title here", "Replace this with a powerful, clear title that immediately communicates the core idea."
- O título carrega a ideia inteira: uma afirmação completa, não um rótulo. Prefira "Agentic delivery cuts lead time in half" a "Resultados".
- Rótulos de seção, tags e números vão em mono: "Insert section title here", "CONTACT ME", "HQs / SAN FRANCISCO". Caixa alta só em rótulos curtos (`mono-label`, `label-bold`).
- Números como protagonistas: "0.0+", "000+", "0x", "#0", sempre em mono com uma legenda curta em minúsculas embaixo ("text").
- Sem emoji, sem exclamação. O fechamento usa uma frase-assinatura: "Ready to meet you where you are and push you forward ↗".

## Fundamentos visuais

### Cor

- Fundo `surface` (branco no tema claro, preto no escuro) e todo texto em `ink`. Não existe cor de marca além do preto, branco e cinzas.
- Superfícies sólidas escuras: `surface-graphite` (#404040) e `surface-carbon` (#222222), sempre com texto branco.
- Cards, tags e barras usam a rampa `gray-050` → `gray-750`. Em grades de vários cards, **varie o tom em degraus da rampa** (claro → escuro ou alternado) em vez de repetir o mesmo cinza; é assim que o template cria ritmo sem cor.
- Texto preto sobre `gray-050` a `gray-400`; texto branco sobre `gray-450` a `gray-750`. Sobre `gray-400` só use branco a partir de 24px.
- `label-internal` só no rótulo de classificação; `logo-tagline` só dentro do logo; `link` só em hiperlinks no tema claro.

### Tipografia

- Títulos em `title-xl` (88px) ou `title-lg` (80px), peso XThin (200), alinhados à esquerda na `space-margin-x`, com no máximo três linhas. Títulos longos descem para `title-md` (64px).
- Corpo em `body` (32px) e `body-sm` (28px), também XThin. Nunca use negrito no corpo; o negrito (`label-bold`) é só para rótulos curtos em caixa alta ou o início de uma regra.
- Mono (`mono-*`) para rótulo de seção (`mono-tag`), tags, títulos de card, anos e todos os números.
- Neue Haas Grotesk Display é uma fonte comercial e **não acompanha este sistema**: instale-a localmente. Na falta dela, a pilha `display` usa Inter Tight (Google Fonts), que é a substituta mais próxima. Roboto Mono vem do Google Fonts.

### Grid e espaçamento

- Margem esquerda fixa de `space-margin-x` (136px) para título, corpo e grids.
- Cabeçalho: rótulo de seção em `mono-tag` na margem esquerda, hairline `rule` de 2px de `space-rule-x` (640px) até 1590px na altura `space-rule-y` (90px) e o logo à direita, a `space-logo-right` da borda.
- Título começa em `space-title-top` (224px). Conteúdo ocupa a metade direita ou a faixa inferior.
- Cards lado a lado com `space-card-gap`; tags e linhas de tags com `space-tag-gap`; grades densas e colunas de linha do tempo com `space-tile-gap`. Respiro interno de `space-card-pad` em cards e `space-tag-pad` em tags e células.

### Formas

- Tudo é reto (`radius-none`). Sem sombras, sem bordas, sem gradientes desenhados.
- A assinatura formal é o **canto superior direito cortado em 45°**: `notch-tag` nas tags cinza e `notch-header` nos cabeçalhos de coluna e barras de título.
- Faixas horizontais dividem o slide: uma metade superior escura (sólida ou foto) e uma faixa inferior branca com tags ou indicadores.

### Imagens

- Use apenas as imagens do grupo `Backgrounds`, sempre escurecidas (já vêm assim) e em sangria total.
- Slides de seção são foto em tela cheia com uma seta `arrow-white.svg` pequena no terço inferior direito, sem texto.
- Imagens dentro de cards ficam em retângulos retos, sem moldura.

### Classificação

- Todo slide interno traz o rótulo "Internal" em `label-internal`, 14px, nos cantos superior esquerdo e inferior direito (componente `ClassificationLabel`).

## Iconografia

- O único ícone é a **seta diagonal ↗** (`arrow-black.svg` para fundos claros, `arrow-white.svg` para escuros). Ela abre cada item de lista, cada card e cada tag com texto, e aparece gigante, em `gray-650`, como motivo decorativo no slide escuro.
- Não use outras bibliotecas de ícones nem emoji. Para marcadores sem seta, use a barra vertical branca de 6px (componente `BarList`).

## Logos

- `airr-logo-black.svg` em fundos claros e `airr-logo-white.svg` em fundos escuros, sempre no canto superior direito do cabeçalho.
- `airr-logo-cover-white.png` grande na capa; `airr-wordmark-horizontal-white.png` no slide de encerramento.
- `airr-compass-white.png` / `airr-compass-black.png` apenas em apresentações de parceria, e então em todos os slides.
