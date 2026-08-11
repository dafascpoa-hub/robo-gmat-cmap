# Robô GMAT CMAP — v153

## Estratégia
A v153 passa a imitar o comportamento observado no Chrome manual.

### Fluxo
1. Login e perfil no GMAT.
2. Gera o relatório 2085.
3. Captura `PlanilhaMateriais.xls`.
4. Fecha explicitamente qualquer popup/janela extra criada pelo download.
5. Recupera/confirma uma página principal viva do GMAT.
6. Abre o relatório 2033 no navegador.
7. Preenche almoxarifado 845, operação 11 e período móvel de 12 meses.
8. Clica no botão REAL `Gerar Planilha`.
9. Detecta o popup do segundo download.
10. Captura a planilha 2033 usando a mesma sessão autenticada.
11. Fecha o popup do 2033.
12. Retorna as duas planilhas ao Portal v127.

## Proteções
- não cria outro Browser Run;
- não depende de cookie extraído depois de contexto destruído;
- fecha páginas extras entre os relatórios;
- se a página principal original tiver sido fechada, procura outra página viva do GMAT no mesmo browser.

## Implantação
Substitua:
- `src/index.js`
- `package.json`
- `wrangler.toml`

Confirme `"versao": "v153"` e faça somente um teste.
