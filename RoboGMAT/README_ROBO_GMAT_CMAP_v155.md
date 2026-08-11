# Robô GMAT CMAP — v155

## Ajuste do período do relatório 2033
O fluxo funcional da v154 foi preservado.

A única alteração funcional foi o cálculo do período do relatório 2033:

- data final = 2 meses antes da data atual;
- data inicial = 12 meses antes da data final.

Exemplo em 11/08/2026:
- início: 11/06/2025
- fim: 11/06/2026

Isso mantém uma janela de 12 meses com defasagem de 2 meses.

## Não alterado
- login;
- relatório 2085;
- captura do JSESSIONID;
- geração/captura do 2033;
- endpoints;
- retorno das duas planilhas;
- Portal v127;
- secrets e wrangler.toml.
