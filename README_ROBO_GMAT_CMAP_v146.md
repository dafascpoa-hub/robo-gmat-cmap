# Robô GMAT CMAP — v146

## Diagnóstico da v145
A v145 registrou que a tentativa de apresentar o navegador como Chrome normal foi executada, porém os headers efetivamente enviados continuaram:
- `HeadlessChrome/128`
- `Linux`
- `en-US`

Portanto, a alteração anterior não chegou à camada de rede.

## Alteração da v146
- mantém todo o fluxo do relatório 2085;
- mantém o POST vivo e correto;
- aplica `Network.setUserAgentOverride` na mesma sessão CDP que observa o POST;
- aplica `Network.setExtraHTTPHeaders` para `Accept-Language`;
- mantém monitoramento de download e captura de resposta;
- registra os headers efetivos para confirmar se a alteração realmente chegou à requisição.

## Implantação
Substitua:
- `src/index.js`
- `package.json`
- `wrangler.toml`

Depois confirme:
`"versao": "v146"`

Faça somente um teste em CMAP > Atualizar estoque.
