# Robô GMAT CMAP v126

Versão limpa, sem duplicidade de funções.

Corrige o erro de build:

`O símbolo "allFrames" foi originalmente declarado aqui`

Mantém:

- login no GMAT;
- seleção do perfil `SMAS - CONSUMO Almoxarifado`;
- tentativa de abrir `Relatórios` por texto/frame;
- fallback por coordenada para o menu antigo do GMAT;
- diagnóstico com etapas, texto visível e screenshot em caso de falha.

Substitua no GitHub do robô:

- `src/index.js`
- `package.json`
- `wrangler.toml`
