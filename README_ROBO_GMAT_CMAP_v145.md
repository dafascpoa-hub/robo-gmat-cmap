# Robô GMAT CMAP — v145

A v144 confirmou POST, sessão e campos corretos, porém a requisição saía identificada como HeadlessChrome/Linux.

## Alterações
- Chrome 128 normal em Windows 10;
- pt-BR;
- `navigator.webdriver` não exposto;
- `Browser.setDownloadBehavior` corrigido com `allowAndName` e `downloadPath`;
- restante do fluxo 2085 preservado.

## Implantação
Substitua `src/index.js`, `package.json` e `wrangler.toml`.
Confirme `"versao": "v145"` e faça somente um teste.
