# Robô GMAT CMAP — v150

A mesma execução gera:
- 2085 — Estoque;
- 2033 — Materiais Consolidados por Consumo.

O 2033 usa órgão 87, almoxarifado 845, operação 11 e período móvel de 12 meses. O fluxo usa `perform=SHEET`/`printPerform=SHEET` e depois o endpoint real `consultaMateriaisConsumoPlanilha.do?perform=run&null`, conforme o HAR manual.
