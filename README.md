# Acervo TEA — Dossiê Científico do Autismo

**Versão 1.2 — capítulo jurídico ampliado sobre Lei Berenice Piana, Educação, Saúde, Assistência Social e ECA; glossário ampliado (114 entradas).**

Site estático responsivo e sem dependências externas obrigatórias.

## Abrir localmente
Abra `index.html` no navegador. A busca, navegação, tema e controles de leitura funcionam localmente. O modo instalável/offline (PWA) passa a funcionar quando o site é servido por HTTPS ou localhost.

## Hospedar
Você pode enviar **todo o conteúdo desta pasta** para:
- GitHub Pages
- Netlify
- Cloudflare Pages
- Vercel (site estático)
- cPanel/FTP de hospedagem comum

Não há etapa de build. O arquivo inicial é `index.html`.

## Atualizar o conteúdo rapidamente
### Opção A — editar o conteúdo diretamente
O conteúdo que o site renderiza está em `content/conteudo.js`. É um arquivo de dados gerado a partir do Word.

### Opção B — fluxo recomendado
1. Edite `fonte/Dossie_Cientifico_Autismo_Estado_da_Arte_1911_2026_Direitos_Educacao_Saude_ECA.docx` no Microsoft Word/LibreOffice.
2. Tenha Python 3 e o pacote `python-docx` instalados.
3. Execute `ATUALIZAR_SITE.bat` no Windows.
4. O script reconstrói `content/conteudo.js` preservando a estrutura de títulos, listas, tabelas e referências.

> O gerador de atualização incluído é intencionalmente simples. Antes de publicar uma nova versão, revise o conteúdo e atualize também a data/versão no documento-fonte.

## Estrutura
- `index.html` — interface principal
- `assets/css/styles.css` — visual responsivo
- `assets/js/app.js` — busca e interações
- `content/conteudo.js` — conteúdo integral
- `content/siglas.js` — glossário estruturado de siglas, abreviações e símbolos
- `SIGLAS_E_ABREVIACOES.md` — referência rápida das 114 entradas por extenso
- `downloads/` — PDF e Word para visitantes
- `fonte/` — Word usado como fonte editorial
- `tools/` — script auxiliar de atualização
- `manifest.webmanifest` + `sw.js` — PWA/offline

## Recursos
- busca instantânea sem servidor;
- navegação por 16 partes e subseções;
- responsivo para desktop, tablet e smartphone;
- modo claro/escuro;
- ajuste de tamanho de fonte;
- progresso de leitura;
- impressão;
- compartilhamento de link;
- referências `[Rxx]` clicáveis;
- glossário com 105 siglas/abreviações/símbolos por extenso;
- significado de siglas disponível por hover/foco e toque no celular;
- PWA instalável;
- sem cookies, analytics ou trackers por padrão;
- acessibilidade de teclado e respeito a `prefers-reduced-motion`.

## Nota
O site é educacional e não substitui avaliação ou orientação profissional individual.


### Versão 1.2
- 52 páginas no dossiê-fonte.
- 93 referências essenciais.
- 114 siglas/abreviações no glossário.
- Seção jurídica ampliada com Lei nº 12.764/2012, Decreto nº 8.368/2014, LBI, LOAS/BPC e ECA.
