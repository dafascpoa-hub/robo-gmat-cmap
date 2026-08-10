# Robô GMAT CMAP — v151

## Correção do erro `Attempted to use detached Frame`
A v150 capturava corretamente o 2085 e depois tentava reutilizar a mesma página/frameset legado do GMAT para abrir o 2033.

O GMAT usa frames antigos e, depois da geração do 2085, algumas referências podem ficar destacadas. A v151:

- mantém o 2085 exatamente como estava;
- abre uma NOVA ABA para o relatório 2033;
- a nova aba usa o MESMO browser/contexto e portanto preserva o `JSESSIONID`;
- isso NÃO cria outro Browser Run;
- confirma o formulário 2033 antes de preencher;
- se ocorrer `detached Frame`, `Target closed` ou `Execution context destroyed`, fecha a aba e tenta uma segunda aba uma única vez;
- depois fecha a aba 2033 e retorna as duas planilhas ao Portal.

## Implantação
Substitua no repositório do robô:
- `src/index.js`
- `package.json`
- `wrangler.toml`

Confirme no endpoint:
`"versao": "v151"`

O Portal v127 não precisa ser alterado para este ajuste.
