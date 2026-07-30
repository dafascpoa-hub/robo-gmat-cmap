# Robô GMAT CMAP v121

Este pacote contém o Worker separado que faz a automação do GMAT para a CMAP.

## Fluxo automatizado

1. Acessa o GMAT.
2. Preenche `#username`.
3. Preenche `#password`.
4. Clica em `Entrar`.
5. Seleciona o perfil `SMAS - CONSUMO Almoxarifado`.
6. Abre o menu `Relatórios`.
7. Clica em `2085 - Geração Planilha Estoque`.
8. Clica em `Gerar Planilha`.
9. Captura o arquivo Excel gerado.
10. Retorna JSON para o Portal DA/SMAS com nome, tamanho, SHA-256, base64 do XLS e etapas executadas.

## Arquivos

- `src/index.js` — Worker completo com Puppeteer.
- `worker_robo_gmat_cmap_v121.txt` — mesmo código em TXT para copiar.
- `package.json` — dependências.
- `wrangler.toml` — modelo de configuração.

## Variáveis/secrets necessários

No Worker `robo-gmat-cmap`:

### Texto simples

- `CMAP_GMAT_URL`
  - `https://gmat.procempa.com.br/gmat/login/login.do?toLogin=true`

Opcionalmente:

- `CMAP_GMAT_PERFIL`
  - `SMAS - CONSUMO Almoxarifado`

- `CMAP_GMAT_RELATORIO`
  - `2085 - Geração Planilha Estoque`

### Segredos

- `CMAP_GMAT_USUARIO`
- `CMAP_GMAT_SENHA`
- `CMAP_GMAT_ROBO_TOKEN`

## Binding do navegador

O código aceita qualquer um destes nomes:

- `BROWSER`
- `browser`
- `NAVEGADOR`
- `navegador`

## Teste

Depois de implantado, acessar:

`https://robo-gmat-cmap.dafascpoa.workers.dev/`

Deve retornar JSON de diagnóstico.

Para executar a automação, enviar POST para a raiz do Worker, com header:

`Authorization: Bearer <CMAP_GMAT_ROBO_TOKEN>`

Body:

```json
{
  "codigoJob": "CMAP-GMAT-TESTE"
}
```

## Conexão com o Portal DA/SMAS

No Worker/Pages do Portal DA, configurar:

- `CMAP_GMAT_ROBO_ENDPOINT`
  - URL do Worker robô, por exemplo:
  - `https://robo-gmat-cmap.dafascpoa.workers.dev/`

- `CMAP_GMAT_ROBO_TOKEN`
  - o mesmo token configurado no Worker robô.

## Observação

Esta v121 captura o XLS e retorna o arquivo em base64.
A próxima etapa é ajustar o Portal DA/SMAS para processar esse base64, gravar a posição atual, histórico compacto e indicadores da CMAP.
