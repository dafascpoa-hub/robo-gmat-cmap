# Robô GMAT CMAP — v152

## Motivo
A v151 eliminou a reutilização do frame do 2085, mas o Browser Rendering fechou a conexão ao tentar criar uma nova página (`Protocol error: Connection closed`).

## Solução
O 2033 não depende mais de uma segunda aba.

Depois que o navegador faz login e gera o 2085, a v152:
1. lê o `JSESSIONID` da sessão autenticada;
2. executa diretamente o POST oficial do relatório 2033:
   `/gmat/uc2033/consultaMateriaisConsumoPesquisa.do`;
3. usa órgão 87, almoxarifado 845, operação 11 e período móvel de 12 meses;
4. confirma o HTML retornado;
5. chama o endpoint real:
   `/gmat/uc2033/consultaMateriaisConsumoPlanilha.do?perform=run&null`;
6. valida e retorna o XLS.

Assim o Browser Rendering é usado para login + 2085, mas o 2033 é concluído por HTTP na mesma sessão.

## Não alterado
- fluxo 2085 já funcional;
- credenciais/secrets;
- wrangler.toml;
- contrato JSON esperado pelo Portal v127.

## Implantação
Substitua `src/index.js`, `package.json` e `wrangler.toml`.
Confirme `"versao": "v152"` antes de um único novo teste.
