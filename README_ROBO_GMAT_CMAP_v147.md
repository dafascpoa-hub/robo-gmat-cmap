# Robô GMAT CMAP — v147

## Objetivo
Parar de assumir que o XLS vem obrigatoriamente na resposta do POST 2085.

## Ajuste
- radar completo de rede antes do clique;
- observa requests, responses, redirects, resourceType, MIME type e Content-Disposition;
- detecta XLS, XLSX, CSV e application/octet-stream;
- captura por Network.getResponseBody;
- captura alternativa por Fetch.takeResponseBodyAsStream + IO.read;
- mantém Browser.downloadWillBegin/downloadProgress;
- rejeita o HTML de ~24 KB;
- preserva login, perfil, relatório 2085 e POST vivo.

## Implantação
Substitua src/index.js, package.json e wrangler.toml.
Confirme `"versao": "v147"` e faça somente um teste.
