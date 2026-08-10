# Robô GMAT CMAP — v149

## Descoberta decisiva no HAR
O POST de `gerarInformacoesPlanilhaPesquisa.do` retorna HTML. Isso é normal.

No HTML retornado existe a função:
`printXLS('perform=run')`

Ela abre:
`gerarInformacoesPlanilhaXLS.do?perform=run&null`

Esse é o endpoint real da exportação.

## Ajuste
- preserva login, perfil e relatório 2085;
- preserva o POST real;
- após o HTML, chama o endpoint XLS real usando a mesma sessão;
- tenta primeiro via `fetch` same-origin autenticado;
- valida os bytes como planilha;
- mantém `window.open()` como fallback;
- preserva radar/CDP.

## Implantação
Substitua `src/index.js`, `package.json` e `wrangler.toml`.
Confirme `"versao": "v149"` e faça somente um teste.
