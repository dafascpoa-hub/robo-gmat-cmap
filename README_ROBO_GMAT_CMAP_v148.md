# Robô GMAT CMAP — v148

## Base do ajuste
O cURL manual que realmente baixa `PlanilhaMateriais.xls` mostrou que o POST, Referer, Sec-Fetch-Dest/Mode/Site/User e Content-Type já eram iguais aos do robô.

As diferenças concretas eram:
- Chrome manual: 150; robô: 128;
- `sec-ch-ua` diferente;
- `Accept-Language` do robô estava duplicando os `q=` após o override CDP;
- cookies extras `_ga/_gid` existem no Chrome, mas não fazem parte da autenticação do GMAT; o `JSESSIONID` continua sendo a sessão relevante.

## Alteração v148
- usa exatamente o User-Agent do cURL manual (Chrome 150 / Windows 10);
- replica `sec-ch-ua`, `sec-ch-ua-mobile` e `sec-ch-ua-platform`;
- corrige o `Accept-Language` para evitar duplicação de `q=`;
- mantém o POST vivo com o `idAlmoxarifado` da sessão atual;
- mantém radar de rede, detecção XLS/XLSX/CSV/octet-stream e stream CDP.

## Implantação
Substituir:
- `src/index.js`
- `package.json`
- `wrangler.toml`

Confirmar:
`"versao": "v148"`

Depois fazer apenas um teste.
