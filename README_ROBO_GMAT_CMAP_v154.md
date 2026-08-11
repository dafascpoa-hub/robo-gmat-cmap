# Robô GMAT CMAP — v154

## Descoberta do diagnóstico v153
O relatório 2085 foi baixado corretamente (`application/vnd.ms-excel`, ~198 KB), mas logo depois o target do Chromium foi fechado/destruído. Por isso não havia página viva para continuar o 2033.

O próprio diagnóstico, porém, mostra que ANTES do fechamento o CDP recebeu os headers efetivos do POST do 2085, incluindo o `JSESSIONID`.

## Estratégia v154
- mantém a geração/captura do 2085 já funcional;
- durante o POST real do 2085, guarda em memória o header `Cookie` efetivo;
- não expõe o valor do cookie em mensagens adicionais;
- depois que o XLS 2085 é capturado, não tenta mais acessar a página/frames/popup;
- usa o JSESSIONID preservado para:
  1. POST oficial do 2033 com `perform=SHEET`;
  2. órgão 87;
  3. almoxarifado 845;
  4. operação 11;
  5. período móvel de 12 meses;
  6. GET do endpoint real `consultaMateriaisConsumoPlanilha.do?perform=run&null`;
- valida o segundo XLS e devolve as duas planilhas ao Portal v127.

## Implantação
Substituir `src/index.js`, `package.json` e `wrangler.toml`.
Confirmar `"versao": "v154"` antes de um único teste.
