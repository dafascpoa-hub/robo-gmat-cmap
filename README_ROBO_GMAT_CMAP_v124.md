# Robô GMAT CMAP v124

Correção sobre a v123:

- Mantém busca por texto e frames.
- Adiciona fallback por coordenada para o menu antigo do GMAT:
  - menu `Relatórios` em torno de x=570, y=126 no viewport 1366x768.
- Usa o trecho inspecionado do GMAT: `left:505px; top:1px` para calibrar o ponto.
- Mantém retorno de diagnóstico em caso de erro.

Substitua no GitHub:

- `src/index.js`
- `package.json`
- `wrangler.toml`
- este README, se desejar.
