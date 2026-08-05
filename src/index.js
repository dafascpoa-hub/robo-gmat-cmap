import puppeteer from "@cloudflare/puppeteer";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization"
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: { "Content-Type": "application/json;charset=utf-8", ...CORS_HEADERS }
  });
}

function textValue(v) {
  return v === undefined || v === null ? "" : String(v);
}

function getBrowserBinding(env) {
  return env.BROWSER || env.browser || env.NAVEGADOR || env.navegador;
}

function assertToken(request, env) {
  const configured = textValue(env.CMAP_GMAT_ROBO_TOKEN).trim();
  if (!configured) return true;
  const header = request.headers.get("Authorization") || "";
  const token = header.replace(/^Bearer\s+/i, "").trim();
  return token && token === configured;
}

async function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function sha256Hex(buffer) {
  const digest = await crypto.subtle.digest("SHA-256", buffer);
  return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, "0")).join("");
}

function parseFileName(headers) {
  const cd = headers["content-disposition"] || headers["Content-Disposition"] || "";
  const utf = cd.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf) {
    try { return decodeURIComponent(utf[1].replace(/"/g, "")); } catch (e) {}
  }
  const normal = cd.match(/filename="?([^";]+)"?/i);
  if (normal) return normal[1];
  return "PlanilhaMateriais.xls";
}

function isExcelContentType(contentType) {
  const ct = String(contentType || "").toLowerCase();
  return ct.includes("excel") ||
         ct.includes("spreadsheet") ||
         ct.includes("application/vnd.ms-excel") ||
         ct.includes("application/vnd.openxmlformats-officedocument") ||
         ct.includes("application/octet-stream") ||
         ct.includes("text/csv");
}

function isExcelResponse(response) {
  const h = response.headers();
  const ct = textValue(h["content-type"]).toLowerCase();
  const cd = textValue(h["content-disposition"]).toLowerCase();
  const url = textValue(response.url()).toLowerCase();
  return (
    cd.includes("attachment") ||
    cd.includes(".xls") ||
    ct.includes("excel") ||
    ct.includes("spreadsheet") ||
    ct.includes("vnd.ms-excel") ||
    url.includes("printxls") ||
    url.includes(".xls")
  );
}

function allFrames(page) {
  try {
    return [page.mainFrame(), ...page.frames().filter(f => f !== page.mainFrame())];
  } catch (e) {
    return [page.mainFrame()];
  }
}

async function waitText(page, text, timeout = 20000) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    for (const frame of allFrames(page)) {
      try {
        const ok = await frame.evaluate(
          (targetText) => document.body && document.body.innerText && document.body.innerText.includes(targetText),
          text
        );
        if (ok) return true;
      } catch (e) {}
    }
    await wait(300);
  }
  throw new Error(`Texto não encontrado: ${text}`);
}

async function pageHasAnyText(page, textos) {
  const list = Array.isArray(textos) ? textos : [textos];
  for (const frame of allFrames(page)) {
    try {
      const txt = await frame.evaluate(() => document.body ? document.body.innerText || "" : "");
      if (list.some(t => txt.includes(t))) return true;
    } catch (e) {}
  }
  return false;
}

async function pageAnyFrameUrlHas(page, trecho) {
  for (const frame of allFrames(page)) {
    try {
      if (String(frame.url() || "").includes(trecho)) return true;
    } catch (e) {}
  }
  return false;
}

async function confirmarTelaRelatorio2085(page) {
  if (await pageAnyFrameUrlHas(page, "uc2085")) return true;
  return pageHasAnyText(page, [
    "Geração de Informações de Materiais em Planilha",
    "Critérios de Pesquisa",
    "Gerar Planilha"
  ]);
}

async function clickGerarPlanilhaEmQualquerFrame(page) {
  const seletores = [
    'input.buttonToolbar[value="Gerar Planilha"]',
    'input[value="Gerar Planilha"]',
    'input[value*="Gerar"]',
    'input[name="submitAction"]',
    'button'
  ];

  for (const frame of allFrames(page)) {
    for (const sel of seletores) {
      try {
        const handles = await frame.$$(sel);
        for (const h of handles) {
          const ok = await h.evaluate(el => {
            const txt = String(el.value || el.innerText || el.textContent || "");
            const visible = !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
            return visible && (!txt || txt.includes("Gerar") || txt.includes("Planilha"));
          }).catch(() => false);
          if (!ok) continue;
          await h.click();
          return true;
        }
      } catch (e) {}
    }
  }

  try {
    await clickByText(page, "Gerar Planilha", { timeout: 8000 });
    return true;
  } catch (e) {
    return false;
  }
}

async function dumpVisibleText(page) {
  const chunks = [];
  for (const frame of allFrames(page)) {
    try {
      const url = frame.url();
      const text = await frame.evaluate(() => document.body ? document.body.innerText.slice(0, 1400) : "");
      chunks.push(`FRAME ${url}\n${text}`);
    } catch (e) {}
  }
  return chunks.join("\n---\n").slice(0, 3500);
}

async function clickByText(page, text, options = {}) {
  const timeoutMs = options.timeout || 15000;
  const exact = !!options.exact;
  const started = Date.now();
  let lastSeen = "";
  while (Date.now() - started < timeoutMs) {
    for (const frame of allFrames(page)) {
      try {
        const ok = await frame.evaluate(({ targetText, exact }) => {
          const normalize = (s) => String(s || "").replace(/\s+/g, " ").trim();
          const wanted = normalize(targetText);
          const candidates = Array.from(document.querySelectorAll("td,div,a,span,button,input,li,tr"));
          const el = candidates.find(e => {
            const visible = !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
            const txt = normalize(e.innerText || e.value || e.textContent);
            return visible && (exact ? txt === wanted : txt.includes(wanted));
          });
          if (!el) return false;
          el.scrollIntoView({ block: "center", inline: "center" });
          const r = el.getBoundingClientRect();
          const x = r.left + r.width / 2;
          const y = r.top + r.height / 2;
          el.dispatchEvent(new MouseEvent("mouseover", { bubbles: true, clientX: x, clientY: y }));
          el.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: x, clientY: y }));
          el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, clientX: x, clientY: y }));
          el.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, clientX: x, clientY: y }));
          el.click();
          return true;
        }, { targetText: text, exact });
        if (ok) return true;
      } catch (e) {}
    }
    try {
      lastSeen = await page.evaluate(() => document.body ? document.body.innerText.slice(0, 800) : "");
    } catch (e) {}
    await wait(300);
  }
  throw new Error(`Elemento com texto não encontrado: ${text}. Texto visível: ${lastSeen}`);
}

async function hoverByText(page, text, options = {}) {
  const timeoutMs = options.timeout || 15000;
  const started = Date.now();
  let lastSeen = "";
  while (Date.now() - started < timeoutMs) {
    for (const frame of allFrames(page)) {
      try {
        const result = await frame.evaluate((targetText) => {
          const normalize = (s) => String(s || "").replace(/\s+/g, " ").trim();
          const wanted = normalize(targetText);
          const candidates = Array.from(document.querySelectorAll("td,div,a,span,button,input,li,tr"));
          const el = candidates.find(e => {
            const visible = !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
            const txt = normalize(e.innerText || e.value || e.textContent);
            return visible && txt.includes(wanted);
          });
          if (!el) return null;
          el.scrollIntoView({ block: "center", inline: "center" });
          const r = el.getBoundingClientRect();
          const x = r.left + r.width / 2;
          const y = r.top + r.height / 2;
          el.dispatchEvent(new MouseEvent("mouseover", { bubbles: true, clientX: x, clientY: y }));
          el.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: x, clientY: y }));
          return { x, y };
        }, text);
        if (result) {
          try { await page.mouse.move(result.x, result.y); } catch (e) {}
          await wait(1200);
          return true;
        }
      } catch (e) {}
    }
    try {
      lastSeen = await page.evaluate(() => document.body ? document.body.innerText.slice(0, 800) : "");
    } catch (e) {}
    await wait(300);
  }
  throw new Error(`Elemento para hover não encontrado: ${text}. Texto visível: ${lastSeen}`);
}

async function clickPerfilGMAT(page, perfil) {
  // A tela de perfil do GMAT é uma tabela antiga. Clicar em qualquer texto parecido
  // pode acertar outra linha; por isso tentamos primeiro a célula exata SMAS.
  const tentativas = [
    () => clickByText(page, perfil, { timeout: 8000, exact: true }),
    () => clickByText(page, perfil, { timeout: 8000 }),
    async () => {
      const clicked = await page.evaluate((perfil) => {
        const norm = s => String(s || "").replace(/\s+/g, " ").trim();
        const alvo = norm(perfil);
        const cells = Array.from(document.querySelectorAll("td"));
        const td = cells.find(c => norm(c.textContent) === alvo);
        if (!td) return false;
        const tr = td.closest("tr") || td;
        tr.scrollIntoView({ block: "center", inline: "center" });
        const r = tr.getBoundingClientRect();
        const x = r.left + r.width / 2;
        const y = r.top + r.height / 2;
        tr.dispatchEvent(new MouseEvent("mouseover", { bubbles: true, clientX: x, clientY: y }));
        tr.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, clientX: x, clientY: y }));
        tr.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, clientX: x, clientY: y }));
        tr.click();
        td.click();
        return true;
      }, perfil);
      if (!clicked) throw new Error("Célula exata do perfil não localizada por script.");
      return true;
    },
    async () => {
      // coordenada aproximada da segunda linha de perfil, conforme tela enviada.
      await page.mouse.click(730, 468);
      return true;
    }
  ];

  let ultimoErro = "";
  for (const tentativa of tentativas) {
    try {
      await Promise.allSettled([
        tentativa(),
        page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 })
      ]);
      await wait(1200);
      const estado = await page.evaluate(() => {
        const txt = document.body ? document.body.innerText || "" : "";
        return {
          texto: txt.slice(0, 1200),
          temMenuPrincipal:
            txt.includes("Relatórios") &&
            txt.includes("Operações Estoque") &&
            txt.includes("Materiais"),
          temIndex:
            String(location.href || "").includes("index.jsp"),
          aindaPerfil:
            txt.includes("SELEÇÃO DE PERFIL") || txt.includes("ÓRGÃO")
        };
      }).catch(() => ({ texto:"", temMenuPrincipal:false, temIndex:false, aindaPerfil:false }));

      // No GMAT antigo pode sobrar texto de frame/cache. Se o menu principal apareceu,
      // a seleção do perfil foi concluída, mesmo que algum texto antigo continue no DOM.
      if (estado.temMenuPrincipal || estado.temIndex) return true;

      ultimoErro = "Clique executado, mas o menu principal do GMAT não apareceu. Texto visto: " + estado.texto.slice(0, 500);
    } catch (e) {
      ultimoErro = e && e.message ? e.message : String(e);
    }
  }
  throw new Error("Não foi possível selecionar o perfil GMAT. " + ultimoErro);
}

async function hoverGMATRelatoriosPorCoordenada(page) {
  const pontos = [
    { x: 570, y: 126 },
    { x: 555, y: 126 },
    { x: 585, y: 126 },
    { x: 570, y: 120 },
    { x: 570, y: 132 }
  ];

  for (const p of pontos) {
    await page.mouse.move(p.x, p.y);
    await wait(900);
    const apareceu = await page.evaluate(() => {
      const txt = document.body ? document.body.innerText || "" : "";
      return txt.includes("2085 - Geração Planilha Estoque") || txt.includes("Geração Planilha Estoque");
    }).catch(() => false);
    if (apareceu) return true;
  }
  return false;
}

async function clickGMATRelatorio2085PorCoordenada(page) {
  // Fallback para menu legado do GMAT.
  // Com viewport 1366x768, o menu Relatórios fica por volta de x=570 y=126,
  // e o item 2085 aparece na sexta linha do submenu, por volta de x=735 y=252.
  const tentativas = [
    { menuX: 570, menuY: 126, itemX: 740, itemY: 253 },
    { menuX: 555, menuY: 126, itemX: 735, itemY: 253 },
    { menuX: 585, menuY: 126, itemX: 760, itemY: 253 },
    { menuX: 570, menuY: 132, itemX: 740, itemY: 258 },
    { menuX: 570, menuY: 120, itemX: 740, itemY: 248 }
  ];

  for (const t of tentativas) {
    await page.mouse.move(t.menuX, t.menuY);
    await wait(900);
    await page.mouse.move(t.itemX, t.itemY);
    await wait(250);
    await page.mouse.click(t.itemX, t.itemY);
    await wait(1800);

    const saiu = await confirmarTelaRelatorio2085(page);
    if (saiu) return true;
  }
  return false;
}

async function findFrameBoxByUrl(page, urlPart) {
  try {
    return await page.evaluate((urlPart) => {
      const frames = Array.from(document.querySelectorAll("frame,iframe"));
      const f = frames.find(el => String(el.src || "").includes(urlPart));
      if (!f) return null;
      const r = f.getBoundingClientRect();
      return { left: r.left, top: r.top, width: r.width, height: r.height, src: f.src };
    }, urlPart);
  } catch (e) {
    return null;
  }
}

async function clickGMATRelatorio2085NoFrameCabecalho(page) {
  // O GMAT usa frames antigos. O menu fica no frame cabecalho.jsp.
  const frame = allFrames(page).find(f => String(f.url()).includes("cabecalho.jsp"));
  if (!frame) return false;

  const frameBox = await findFrameBoxByUrl(page, "cabecalho.jsp");
  if (!frameBox) return false;

  const relBox = await frame.evaluate(() => {
    const norm = s => String(s || "").replace(/\s+/g, " ").trim();
    const elems = Array.from(document.querySelectorAll("div,td,span,a"));
    const el = elems.find(e => {
      const txt = norm(e.innerText || e.textContent);
      const visible = !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
      return visible && txt === "Relatórios";
    }) || elems.find(e => {
      const txt = norm(e.innerText || e.textContent);
      const visible = !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
      return visible && txt.includes("Relatórios");
    });
    if (!el) return null;
    const r = el.getBoundingClientRect();
    return { left: r.left, top: r.top, width: r.width, height: r.height, text: norm(el.innerText || el.textContent) };
  }).catch(() => null);

  if (!relBox) return false;

  const menuX = frameBox.left + relBox.left + Math.max(20, Math.min(relBox.width / 2, relBox.width - 10));
  const menuY = frameBox.top + relBox.top + Math.max(8, Math.min(relBox.height / 2, relBox.height - 3));

  // Primeiro tenta abrir o submenu por mouse real.
  await page.mouse.move(menuX, menuY);
  await wait(700);
  await page.mouse.move(menuX + 20, menuY + 2);
  await wait(900);

  // Depois tenta encontrar o item 2085 que pode ter sido criado no próprio frame.
  const itemBox = await frame.evaluate(() => {
    const norm = s => String(s || "").replace(/\s+/g, " ").trim();
    const elems = Array.from(document.querySelectorAll("div,td,span,a"));
    const el = elems.find(e => {
      const txt = norm(e.innerText || e.textContent);
      const visible = !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
      return visible && txt.includes("2085") && txt.includes("Geração Planilha Estoque");
    });
    if (!el) return null;
    const r = el.getBoundingClientRect();
    return { left: r.left, top: r.top, width: r.width, height: r.height, text: norm(el.innerText || el.textContent) };
  }).catch(() => null);

  if (itemBox) {
    const itemX = frameBox.left + itemBox.left + Math.max(20, Math.min(itemBox.width / 2, itemBox.width - 10));
    const itemY = frameBox.top + itemBox.top + Math.max(8, Math.min(itemBox.height / 2, itemBox.height - 3));
    await page.mouse.move(itemX, itemY);
    await wait(250);
    await page.mouse.click(itemX, itemY);
    await wait(1800);
  } else {
    // Se o item não ficou visível para o DOM, usa offset relativo ao próprio menu.
    // Pela tela real: 2085 é a 6ª linha do submenu, cerca de 127px abaixo do topo do botão Relatórios.
    const itemX = menuX + 175;
    const itemY = menuY + 126;
    await page.mouse.move(itemX, itemY);
    await wait(250);
    await page.mouse.click(itemX, itemY);
    await wait(1800);
  }

  const saiu = await confirmarTelaRelatorio2085(page);
  return !!saiu;
}

async function atualizarEstoqueGMAT(request, env) {
  if (!assertToken(request, env)) {
    return json({ ok: false, erro: "Token inválido ou ausente." }, 401);
  }

  const browserBinding = getBrowserBinding(env);
  if (!browserBinding) {
    return json({
      ok: false,
      erro: "Binding do navegador não encontrado. Configure o binding como BROWSER ou NAVEGADOR.",
      bindingsDisponiveis: Object.keys(env)
    }, 500);
  }

  const url = textValue(env.CMAP_GMAT_URL || "https://gmat.procempa.com.br/gmat/login/login.do?toLogin=true").trim();
  const usuario = textValue(env.CMAP_GMAT_USUARIO).trim();
  const senha = textValue(env.CMAP_GMAT_SENHA);
  const perfil = textValue(env.CMAP_GMAT_PERFIL || "SMAS - CONSUMO Almoxarifado").trim();
  const relatorio = textValue(env.CMAP_GMAT_RELATORIO || "2085 - Geração Planilha Estoque").trim();

  if (!usuario || !senha) {
    return json({ ok: false, erro: "CMAP_GMAT_USUARIO e/ou CMAP_GMAT_SENHA não configurados." }, 500);
  }

  let payload = {};
  try { payload = await request.json(); } catch (e) {}
  const codigoJob = textValue(payload.codigoJob || "").trim();

  const etapas = [];
  const log = (etapa) => etapas.push({ etapa, em: new Date().toISOString() });

  let browser;
  let page;
  let captured = null;

  try {
    log("abrindo navegador");
    browser = await puppeteer.launch(browserBinding);
    page = await browser.newPage();
    await page.setViewport({ width: 1366, height: 768 });

    const paginasMonitoradas = new Set();

    async function anexarCapturaArquivo(pg, origem = "pagina") {
      if (!pg || paginasMonitoradas.has(pg)) return;
      paginasMonitoradas.add(pg);

      try {
        pg.on("response", async (response) => {
          try {
            if (captured || !isExcelResponse(response)) return;
            const headers = response.headers();
            const ab = await response.arrayBuffer();
            if (!ab || !ab.byteLength) return;
            captured = {
              url: response.url(),
              status: response.status(),
              headers,
              arrayBuffer: ab,
              arquivoNome: parseFileName(headers),
              origem
            };
          } catch (e) {}
        });
      } catch (e) {}
    }

    await anexarCapturaArquivo(page, "pagina-principal");

    try {
      browser.on("targetcreated", async (target) => {
        try {
          const pg = await target.page();
          if (pg) {
            await pg.setViewport({ width: 1366, height: 768 }).catch(() => {});
            await anexarCapturaArquivo(pg, "popup");
            etapas.push({
              etapa: "popup detectado",
              em: new Date().toISOString(),
              detalhe: String(pg.url() || target.url() || "").slice(0, 500)
            });
          }
        } catch (e) {}
      });
    } catch (e) {}

    log("acessando GMAT");
    await page.goto(url, { waitUntil: "networkidle0", timeout: 45000 });

    log("preenchendo login");
    await page.waitForSelector("#username, input[name='nomeUsuario']", { timeout: 20000 });
    await page.click("#username, input[name='nomeUsuario']", { clickCount: 3 });
    await page.type("#username, input[name='nomeUsuario']", usuario, { delay: 20 });
    await page.click("#password, input[name='senhaUsuario']", { clickCount: 3 });
    await page.type("#password, input[name='senhaUsuario']", senha, { delay: 20 });

    log("enviando login");
    await Promise.allSettled([
      page.click('input[type="submit"][value="Entrar"], input.botao'),
      page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 })
    ]);

    log("selecionando perfil");
    await waitText(page, "SELEÇÃO DE PERFIL", 30000).catch(() => {});
    await waitText(page, perfil, 30000);
    await clickPerfilGMAT(page, perfil);

    log("confirmando tela principal");
    try {
      await waitText(page, "Relatórios", 20000);
    } catch (e) {
      const txt = await dumpVisibleText(page);
      throw new Error("Após selecionar o perfil, a tela principal do GMAT não foi confirmada. " + txt.slice(0, 900));
    }

    log("abrindo menu relatórios");
    await wait(1200);
    try {
      await hoverByText(page, "Relatórios", { timeout: 8000 });
    } catch (e) {
      etapas.push({
        etapa: "aviso",
        em: new Date().toISOString(),
        detalhe: "Não achou Relatórios por texto/frame. Tentando coordenada fixa do menu GMAT."
      });
      const okCoord = await hoverGMATRelatoriosPorCoordenada(page);
      if (!okCoord) {
        throw new Error("Menu Relatórios não abriu por texto nem por coordenada.");
      }
    }
    await wait(800);

    log("selecionando relatório 2085");
    try {
      await clickByText(page, relatorio, { timeout: 12000 });
    } catch (e) {
      etapas.push({
        etapa: "aviso",
        em: new Date().toISOString(),
        detalhe: "Não clicou relatório por texto após primeira abertura. Reabrindo menu por coordenada e tentando clique direto no item 2085."
      });
      await hoverGMATRelatoriosPorCoordenada(page);
      await wait(800);

      try {
        await clickByText(page, relatorio, { timeout: 6000 });
      } catch (e2) {
        let okCoordItem = await clickGMATRelatorio2085NoFrameCabecalho(page);
        if (!okCoordItem) okCoordItem = await clickGMATRelatorio2085PorCoordenada(page);
        if (!okCoordItem) {
          throw new Error("Não foi possível clicar no relatório 2085 por texto, por frame cabecalho.jsp nem por coordenada. " + (e2 && e2.message ? e2.message : String(e2)));
        }
      }
    }

    await Promise.allSettled([
      page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 }),
      waitText(page, "Gerar Planilha", 30000)
    ]);

    if (!(await confirmarTelaRelatorio2085(page))) {
      const txt = await dumpVisibleText(page);
      throw new Error("Relatório 2085 não confirmado após clique. Texto visível: " + txt.slice(0, 900));
    }

    log("baixando planilha por POST no contexto do navegador");
    const browser2085 = await baixarPlanilha2085FetchNoBrowser(page, etapas);
    if (browser2085) {
      captured = browser2085;
    }

    if (!captured) {
      log("baixando planilha por POST direto");
      const direto2085 = await baixarPlanilha2085PostDireto(page, etapas);
      if (direto2085) {
        captured = direto2085;
      }
    }

    log("gerando planilha");
    const paginasAntes = await browser.pages().catch(() => []);
    for (const pg of paginasAntes) await anexarCapturaArquivo(pg, "pagina-existente");

    const clicouGerar = captured ? true : await clickGerarPlanilhaEmQualquerFrame(page);
    if (!clicouGerar) {
      const txt = await dumpVisibleText(page);
      throw new Error("Botão Gerar Planilha não encontrado em nenhum frame. Texto visível: " + txt.slice(0, 900));
    }

    await Promise.allSettled([
      page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 })
    ]);

    log("aguardando popup/download");
    const started = Date.now();
    let popupVisto = false;
    let urlsPaginas = [];
    while (!captured && Date.now() - started < 90000) {
      try {
        const paginas = await browser.pages();
        urlsPaginas = [];
        for (const pg of paginas) {
          await anexarCapturaArquivo(pg, "pagina-ou-popup");
          urlsPaginas.push(await pg.url().catch(() => ""));
        }
        if (paginas.length > paginasAntes.length) popupVisto = true;
      } catch (e) {}
      await wait(600);
    }

    if (!captured) {
      log("tentando submit direto do formulário");
      const capturadoDireto = await tentarSubmitDiretoRelatorio2085(page, captured, etapas);
      if (capturadoDireto) {
        captured = capturadoDireto;
      }
    }

    if (!captured) {
      log("tentando fallback exaustivo do relatório 2085");
      const capturadoExaustivo = await tentarVariosSubmits2085(page, etapas);
      if (capturadoExaustivo) {
        captured = capturadoExaustivo;
      }
    }

    if (!captured) {
      log("tentando post por formulário/navegação real");
      const capturadoFormReal = await submeterPOST2085ComoFormularioReal(page, browser, etapas);
      if (capturadoFormReal) captured = capturadoFormReal;
    }

    if (!captured) {
      const title = await page.title().catch(() => "");
      const text = await dumpVisibleText(page).catch(async () => {
        return await page.evaluate(() => document.body ? document.body.innerText.slice(0, 1000) : "").catch(() => "");
      });
      throw new Error(
        "A planilha não foi capturada após clique, popup, submit direto, submit exaustivo e URLs prováveis. " +
        (popupVisto ? "Popup/janela foi detectado, mas o arquivo não foi interceptado. " : "Nenhum popup novo foi confirmado. ") +
        "URLs abertas: " + urlsPaginas.filter(Boolean).join(" | ").slice(0, 900) +
        " | Última página: " + title + " | Texto visível: " + String(text).slice(0, 700)
      );
    }

    log("planilha capturada");
    const hash = await sha256Hex(captured.arrayBuffer);
    const tamanho = captured.arrayBuffer.byteLength;
    const bytes = new Uint8Array(captured.arrayBuffer);
    let binary = "";
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    const base64 = btoa(binary);

    return json({
      ok: true,
      codigoJob,
      sistema: "GMAT",
      perfil,
      relatorio,
      arquivoNome: captured.arquivoNome || "PlanilhaMateriais.xls",
      arquivoTamanho: tamanho,
      arquivoHash: hash,
      contentType: captured.headers["content-type"] || "application/vnd.ms-excel",
      origemCaptura: captured.origem || "",
      urlCaptura: captured.url || "",
      arquivoBase64: base64,
      etapas,
      concluidoEm: new Date().toISOString()
    });
  } catch (e) {
    let screenshotBase64 = "";
    try {
      if (page) {
        screenshotBase64 = await page.screenshot({ encoding: "base64", fullPage: false }) || "";
      }
    } catch (s) {}

    let textoVisivel = "";
    try {
      if (page) textoVisivel = await dumpVisibleText(page);
    } catch (t) {}

    return json({
      ok: false,
      codigoJob,
      erro: e && e.message ? e.message : String(e),
      etapas,
      textoVisivel,
      screenshotBase64
    }, 500);
  } finally {
    try { if (page) await page.close(); } catch (e) {}
    try { if (browser) await browser.close(); } catch (e) {}
  }
}

async function getCookiesHeaderDaPagina(page) {
  try {
    const cookies = await page.cookies();
    return cookies.map(c => `${c.name}=${c.value}`).join("; ");
  } catch (e) {
    return "";
  }
}

function absolutizarUrlGMAT(baseUrl, action) {
  try {
    return new URL(action || "", baseUrl || "https://gmat.procempa.com.br/gmat/").toString();
  } catch (e) {
    return String(action || "");
  }
}

async function tentarSubmitDiretoRelatorio2085(page, capturedRef, etapas) {
  // Plano B para GMAT legado: se o clique no botão não expõe download/response,
  // coleta o formulário do relatório no frame uc2085 e faz POST/GET direto com cookies da sessão.
  let alvo = null;

  for (const frame of allFrames(page)) {
    try {
      const urlFrame = String(frame.url() || "");
      const dados = await frame.evaluate(() => {
        const norm = s => String(s || "").replace(/\s+/g, " ").trim();
        const forms = Array.from(document.querySelectorAll("form"));
        const candidates = forms.map((form, idx) => {
          const fd = new FormData(form);
          const params = [];
          for (const [k, v] of fd.entries()) params.push([k, String(v)]);
          const buttons = Array.from(form.querySelectorAll("input,button")).map(el => ({
            name: el.name || "",
            value: el.value || el.innerText || el.textContent || "",
            type: el.type || el.tagName
          }));
          return {
            idx,
            action: form.getAttribute("action") || "",
            method: (form.getAttribute("method") || "GET").toUpperCase(),
            text: norm(form.innerText || form.textContent || "").slice(0, 1000),
            params,
            buttons
          };
        });
        return candidates;
      });

      const form = dados.find(f =>
        String(f.text || "").includes("Critérios de Pesquisa") ||
        String(f.text || "").includes("Geração de Informações de Materiais") ||
        String(f.action || "").includes("uc2085") ||
        JSON.stringify(f.buttons || []).includes("Gerar")
      ) || dados[0];

      if (form && (urlFrame.includes("uc2085") || String(form.action || "").includes("uc2085") || String(form.text || "").includes("Critérios"))) {
        alvo = { frameUrl: urlFrame, form };
        break;
      }
    } catch (e) {}
  }

  if (!alvo) {
    etapas.push({ etapa: "submit direto", em: new Date().toISOString(), detalhe: "Nenhum formulário do relatório 2085 encontrado nos frames." });
    return null;
  }

  const form = alvo.form;
  const baseUrl = alvo.frameUrl || page.url();
  const url = absolutizarUrlGMAT(baseUrl, form.action || baseUrl);
  const cookie = await getCookiesHeaderDaPagina(page);

  const params = new URLSearchParams();
  for (const [k, v] of (form.params || [])) {
    if (k) params.append(k, v);
  }

  // Tenta incluir submitAction/Gerar Planilha se não veio no FormData.
  const temSubmit = Array.from(params.keys()).some(k => /submit|action|acao/i.test(k));
  if (!temSubmit) {
    const btn = (form.buttons || []).find(b => String(b.value || "").includes("Gerar") || String(b.name || "").includes("submit"));
    if (btn && btn.name) params.append(btn.name, btn.value || "Gerar Planilha");
    else params.append("submitAction", "Gerar Planilha");
  }

  etapas.push({
    etapa: "submit direto",
    em: new Date().toISOString(),
    detalhe: `Tentando ${form.method || "GET"} direto em ${url.slice(0, 400)} com ${Array.from(params.keys()).length} campos.`
  });

  const headers = {
    "Cookie": cookie,
    "Referer": baseUrl,
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/vnd.ms-excel,application/octet-stream,text/html,*/*"
  };

  let resp;
  if ((form.method || "GET").toUpperCase() === "POST") {
    headers["Content-Type"] = "application/x-www-form-urlencoded";
    resp = await fetch(url, { method: "POST", headers, body: params.toString(), redirect: "follow" });
  } else {
    const u = new URL(url);
    for (const [k, v] of params.entries()) u.searchParams.append(k, v);
    resp = await fetch(u.toString(), { method: "GET", headers, redirect: "follow" });
  }

  const respHeaders = {};
  resp.headers.forEach((v, k) => respHeaders[k] = v);
  const ab = await resp.arrayBuffer();

  etapas.push({
    etapa: "submit direto resposta",
    em: new Date().toISOString(),
    detalhe: `HTTP ${resp.status}; content-type=${respHeaders["content-type"] || ""}; bytes=${ab ? ab.byteLength : 0}; url=${resp.url.slice(0, 500)}`
  });

  if (ab && ab.byteLength) {
    const contentType = String(respHeaders["content-type"] || "");
    const cd = String(respHeaders["content-disposition"] || "");
    const pareceArquivo =
      isExcelContentType(contentType) ||
      /attachment|xls|xlsx|csv|octet/i.test(cd) ||
      ab.byteLength > 5000;

    if (pareceArquivo) {
      return {
        url: resp.url,
        status: resp.status,
        headers: respHeaders,
        arrayBuffer: ab,
        arquivoNome: parseFileName(respHeaders),
        origem: "submit-direto-frame"
      };
    }
  }

  return null;
}


async function dumpFormularios2085(page) {
  const out = [];
  for (const frame of allFrames(page)) {
    try {
      const url = String(frame.url() || "");
      const forms = await frame.evaluate(() => {
        const norm = s => String(s || "").replace(/\s+/g, " ").trim();
        return Array.from(document.querySelectorAll("form")).map((form, idx) => {
          const fields = [];
          const fd = new FormData(form);
          for (const [k, v] of fd.entries()) fields.push([k, String(v).slice(0, 200)]);
          const inputs = Array.from(form.querySelectorAll("input,select,textarea,button")).map(el => ({
            tag: el.tagName,
            type: el.type || "",
            name: el.name || "",
            id: el.id || "",
            value: String(el.value || el.innerText || el.textContent || "").slice(0, 200)
          }));
          return {
            idx,
            action: form.getAttribute("action") || "",
            method: (form.getAttribute("method") || "GET").toUpperCase(),
            text: norm(form.innerText || form.textContent || "").slice(0, 1200),
            fields,
            inputs
          };
        });
      });
      out.push({ frameUrl: url, forms });
    } catch (e) {}
  }
  return out;
}

async function fetchComSessaoGMAT(page, url, options = {}) {
  const cookie = await getCookiesHeaderDaPagina(page);
  const headers = {
    "Cookie": cookie,
    "Referer": options.referer || page.url(),
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/vnd.ms-excel,application/octet-stream,text/csv,text/html,*/*",
    ...(options.headers || {})
  };
  return fetch(url, { ...options, headers, redirect: "follow" });
}

async function respostaComoCaptura(resp, origem, etapas) {
  const respHeaders = {};
  resp.headers.forEach((v, k) => respHeaders[k] = v);
  const ab = await resp.arrayBuffer();
  const ct = String(respHeaders["content-type"] || "");
  const cd = String(respHeaders["content-disposition"] || "");
  const pareceArquivo =
    isExcelContentType(ct) ||
    /attachment|xls|xlsx|csv|octet/i.test(cd) ||
    (ab && ab.byteLength > 5000 && !ct.includes("text/html"));

  etapas.push({
    etapa: origem + " resposta",
    em: new Date().toISOString(),
    detalhe: `HTTP ${resp.status}; content-type=${ct}; content-disposition=${cd}; bytes=${ab ? ab.byteLength : 0}; url=${String(resp.url || "").slice(0, 700)}`
  });

  if (pareceArquivo && ab && ab.byteLength) {
    return {
      url: resp.url,
      status: resp.status,
      headers: respHeaders,
      arrayBuffer: ab,
      arquivoNome: parseFileName(respHeaders),
      origem
    };
  }

  return null;
}

async function tentarVariosSubmits2085(page, etapas) {
  const dump = await dumpFormularios2085(page);
  const candidatos = [];

  for (const frameInfo of dump) {
    for (const form of frameInfo.forms || []) {
      const texto = `${form.text || ""} ${form.action || ""} ${frameInfo.frameUrl || ""}`;
      const relevante =
        texto.includes("uc2085") ||
        texto.includes("Geração de Informações de Materiais") ||
        texto.includes("Critérios de Pesquisa") ||
        texto.includes("Gerar Planilha");
      if (!relevante) continue;

      const baseUrl = frameInfo.frameUrl || page.url();
      const actionUrl = absolutizarUrlGMAT(baseUrl, form.action || baseUrl);

      const baseParams = new URLSearchParams();
      for (const [k, v] of (form.fields || [])) {
        if (k) baseParams.append(k, v);
      }

      const btns = (form.inputs || []).filter(i => 
        String(i.value || "").includes("Gerar") ||
        String(i.name || "").match(/submit|action|acao/i) ||
        String(i.id || "").match(/submit|gerar/i)
      );

      const variantes = [];

      variantes.push({ nome: "form-original", method: form.method || "GET", url: actionUrl, params: new URLSearchParams(baseParams), referer: baseUrl });

      for (const btn of btns) {
        const p = new URLSearchParams(baseParams);
        if (btn.name) p.set(btn.name, btn.value || "Gerar Planilha");
        variantes.push({ nome: "form-botao-" + (btn.name || btn.id || "semnome"), method: form.method || "GET", url: actionUrl, params: p, referer: baseUrl });
      }

      // Variações comuns em sistemas Java antigos.
      for (const [k, v] of [
        ["submitAction", "Gerar Planilha"],
        ["method", "gerarPlanilha"],
        ["acao", "Gerar Planilha"],
        ["action", "Gerar Planilha"]
      ]) {
        const p = new URLSearchParams(baseParams);
        p.set(k, v);
        variantes.push({ nome: "var-" + k, method: form.method || "POST", url: actionUrl, params: p, referer: baseUrl });
      }

      for (const v of variantes) candidatos.push(v);
    }
  }

  etapas.push({
    etapa: "submit exaustivo",
    em: new Date().toISOString(),
    detalhe: `Formulários/frame encontrados: ${dump.length}; candidatos de submit: ${candidatos.length}`
  });

  for (const c of candidatos.slice(0, 20)) {
    try {
      etapas.push({
        etapa: "submit tentativa",
        em: new Date().toISOString(),
        detalhe: `${c.nome}; ${c.method}; ${String(c.url).slice(0, 500)}; campos=${Array.from(c.params.keys()).join(",").slice(0, 500)}`
      });

      let resp;
      if (String(c.method || "GET").toUpperCase() === "POST") {
        resp = await fetchComSessaoGMAT(page, c.url, {
          method: "POST",
          referer: c.referer,
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: c.params.toString()
        });
      } else {
        const u = new URL(c.url);
        for (const [k, v] of c.params.entries()) u.searchParams.set(k, v);
        resp = await fetchComSessaoGMAT(page, u.toString(), { method: "GET", referer: c.referer });
      }

      const cap = await respostaComoCaptura(resp, "submit-exaustivo-" + c.nome, etapas);
      if (cap) return cap;
    } catch (e) {
      etapas.push({
        etapa: "submit tentativa erro",
        em: new Date().toISOString(),
        detalhe: `${c.nome}: ${e && e.message ? e.message : String(e)}`
      });
    }
  }

  // Tentativas diretas de URLs prováveis, mantendo cookies da sessão.
  const urls = [];
  for (const frame of allFrames(page)) {
    const u = String(frame.url() || "");
    if (u.includes("uc2085")) {
      const base = u.split("?")[0];
      urls.push(base);
      urls.push(base.replace("gerarInformacoesPlanilhaPesquisa.do", "gerarInformacoesPlanilha.do"));
      urls.push(base.replace("Pesquisa.do", ".do"));
      urls.push(base + "?submitAction=Gerar+Planilha");
      urls.push(base + "?method=gerarPlanilha");
    }
  }

  const unicas = Array.from(new Set(urls.filter(Boolean))).slice(0, 10);
  etapas.push({
    etapa: "urls diretas",
    em: new Date().toISOString(),
    detalhe: unicas.join(" | ").slice(0, 1200)
  });

  for (const u of unicas) {
    try {
      const resp = await fetchComSessaoGMAT(page, u, { method: "GET", referer: page.url() });
      const cap = await respostaComoCaptura(resp, "url-direta", etapas);
      if (cap) return cap;
    } catch (e) {
      etapas.push({
        etapa: "url direta erro",
        em: new Date().toISOString(),
        detalhe: `${u}: ${e && e.message ? e.message : String(e)}`
      });
    }
  }

  etapas.push({
    etapa: "formularios coletados",
    em: new Date().toISOString(),
    detalhe: JSON.stringify(dump).slice(0, 3500)
  });

  return null;
}


function paramsPOST2085PadraoGMAT() {
  const params = new URLSearchParams();

  // Capturado manualmente via DevTools / Copy as cURL no GMAT em 03/08/2026.
  // Não inclui cookies fixos: os cookies são sempre os da sessão atual do robô.
  params.append("perform", "printXLS");
  params.append("actionForward", "success");
  params.append("strutsFormName", "gerarInformacoesPlanilhaPesquisaForm");
  params.append("user", "");
  params.append("dominio", "");
  params.append("validate", "true");
  params.append("printPerform", "printXLS");
  params.append("pesquisar", "false");
  params.append("chave", "");
  params.append("idAlmoxarifado", "br.com.procempa.tesouro.psi.gmat.combo.ComboAlmoxarifado@6b6d4e81");
  params.append("tipo", "popularCombo");
  params.append("idOrgao.id", "87");
  params.append("idAlmoxarifado.id", "845");
  params.append("filtro.nmIniMaterial", "");
  params.append("filtro.nmCtmMaterial", "");
  params.append("idEspCla.id", "");
  params.append("idSubEspCla.id", "");
  params.append("filtro.ctrlValidade", "");
  params.append("idArea.id", "");
  params.append("idEstante.id", "");
  params.append("idPrateleira.id", "");
  params.append("idTipoMaterial.id", "");
  params.append("filtro.regPreco", "");
  params.append("filtro.situacao", "");
  params.append("filtro.dtIniMovEntrada", "");
  params.append("filtro.dtFimMovEntrada", "");
  params.append("filtro.dtIniMovSaida", "");
  params.append("filtro.dtFimMovSaida", "");
  params.append("filtro.listarMaterialLote", "SIM");
  params.append("filtro.listarMaterialLote", "");
  params.append("defaultSearch.pageSize", "0");
  params.append("defaultSearch.orderField", "");
  params.append("defaultSearch.orderDirection", "");

  return params;
}


function arrayBufferFromBase64GMAT(base64) {
  const bin = atob(String(base64 || ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

function magicArquivoGMAT(ab) {
  try {
    const b = new Uint8Array(ab || new ArrayBuffer(0));
    if (b.length >= 8 && b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0) return "xls-cdf";
    if (b.length >= 4 && b[0] === 0x50 && b[1] === 0x4B && b[2] === 0x03 && b[3] === 0x04) return "xlsx-zip";
    if (b.length >= 1 && b[0] === 0x3C) return "html";
  } catch (e) {}
  return "desconhecido";
}

function textoAmostraGMAT(ab, limite = 1600) {
  try {
    return new TextDecoder("utf-8").decode((ab || new ArrayBuffer(0)).slice(0, limite));
  } catch (e) {
    try { return new TextDecoder("iso-8859-1").decode((ab || new ArrayBuffer(0)).slice(0, limite)); } catch (e2) { return ""; }
  }
}

function htmlTemCabecalhoPlanilhaGMAT(texto) {
  const t = String(texto || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ").toLowerCase();
  return (t.includes("cod mat") || t.includes("codigo mat")) &&
         t.includes("nome mat") &&
         t.includes("qtde estoque");
}

function capturaParecePlanilhaGMAT(ab, headers = {}, url = "") {
  const ct = String(headers["content-type"] || headers["Content-Type"] || "").toLowerCase();
  const cd = String(headers["content-disposition"] || headers["Content-Disposition"] || "").toLowerCase();
  const magic = magicArquivoGMAT(ab);
  if (magic === "xls-cdf" || magic === "xlsx-zip") return { ok:true, motivo:magic };
  if (/attachment|\.xls|\.xlsx|octet|excel|spreadsheet/i.test(cd) || isExcelContentType(ct)) return { ok:true, motivo:"headers-excel" };
  if (magic === "html" || ct.includes("text/html")) {
    const amostra = textoAmostraGMAT(ab, 2500);
    if (htmlTemCabecalhoPlanilhaGMAT(amostra)) return { ok:true, motivo:"html-com-cabecalho-gmat" };
    return { ok:false, motivo:"html-sem-cabecalho-gmat", amostra };
  }
  return { ok:false, motivo:"sem-magic-xls", amostra:textoAmostraGMAT(ab, 800) };
}

async function baixarPlanilha2085FetchNoBrowser(page, etapas) {
  const endpoint = "https://gmat.procempa.com.br/gmat/uc2085/gerarInformacoesPlanilhaPesquisa.do";
  const body = paramsPOST2085PadraoGMAT().toString();

  etapas.push({ etapa:"post xls via browser", em:new Date().toISOString(), detalhe:"Executando fetch dentro do contexto autenticado do GMAT para capturar o corpo real da resposta." });

  let resultado = null;
  const frames = allFrames(page);
  for (const frame of frames) {
    try {
      const frameUrl = String(frame.url() || "");
      if (!frameUrl.includes("gmat.procempa.com.br")) continue;
      resultado = await frame.evaluate(async ({ endpoint, body }) => {
        const resp = await fetch(endpoint, {
          method: "POST",
          credentials: "include",
          redirect: "follow",
          headers: {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/vnd.ms-excel,application/octet-stream,*/*;q=0.8",
            "Content-Type": "application/x-www-form-urlencoded",
            "Cache-Control": "max-age=0"
          },
          body
        });
        const headers = {};
        resp.headers.forEach((v,k)=>headers[k]=v);
        const ab = await resp.arrayBuffer();
        const bytes = new Uint8Array(ab);
        let base64 = "";
        const chunk = 0x8000;
        for (let i=0;i<bytes.length;i+=chunk) {
          let s = "";
          const part = bytes.subarray(i, i+chunk);
          for (let j=0;j<part.length;j++) s += String.fromCharCode(part[j]);
          base64 += btoa(s);
        }
        // O loop acima concatena blocos base64 independentes; para evitar corromper, faz conversão única em chunks binários.
        let binary = "";
        for (let i=0;i<bytes.length;i+=chunk) {
          const part = bytes.subarray(i, i+chunk);
          binary += String.fromCharCode(...part);
        }
        base64 = btoa(binary);
        return { ok: resp.ok, status: resp.status, url: resp.url, headers, bytes: bytes.length, base64 };
      }, { endpoint, body });
      break;
    } catch (e) {
      etapas.push({ etapa:"post xls via browser erro", em:new Date().toISOString(), detalhe:`Frame ${String(frame.url()||"").slice(0,300)}: ${e && e.message ? e.message : String(e)}` });
    }
  }

  if (!resultado) {
    etapas.push({ etapa:"post xls via browser", em:new Date().toISOString(), detalhe:"Nenhum frame GMAT permitiu executar fetch interno." });
    return null;
  }

  const ab = arrayBufferFromBase64GMAT(resultado.base64);
  const headers = resultado.headers || {};
  const avaliacao = capturaParecePlanilhaGMAT(ab, headers, resultado.url || endpoint);
  etapas.push({
    etapa:"post xls via browser resposta",
    em:new Date().toISOString(),
    detalhe:`HTTP ${resultado.status}; bytes=${resultado.bytes}; magic=${magicArquivoGMAT(ab)}; avaliacao=${avaliacao.motivo}; content-type=${headers["content-type"]||""}; content-disposition=${headers["content-disposition"]||""}; url=${String(resultado.url||"").slice(0,500)}`
  });

  if (!resultado.ok || !avaliacao.ok) {
    etapas.push({ etapa:"post xls via browser não capturou arquivo", em:new Date().toISOString(), detalhe:`${avaliacao.motivo}; amostra=${String(avaliacao.amostra||"").slice(0,1000)}` });
    return null;
  }

  return {
    url: resultado.url,
    status: resultado.status,
    headers,
    arrayBuffer: ab,
    arquivoNome: parseFileName(headers) || "PlanilhaMateriais.xls",
    origem: "post-xls-browser-v137"
  };
}

async function baixarPlanilha2085PostDireto(page, etapas) {
  const endpoint = "https://gmat.procempa.com.br/gmat/uc2085/gerarInformacoesPlanilhaPesquisa.do";
  const referer = "https://gmat.procempa.com.br/gmat/uc2085/gerarInformacoesPlanilhaPesquisa.do";
  const cookie = await getCookiesHeaderDaPagina(page);
  const body = paramsPOST2085PadraoGMAT().toString();

  etapas.push({
    etapa: "post xls direto",
    em: new Date().toISOString(),
    detalhe: `POST direto no endpoint do relatório 2085; campos=${Array.from(paramsPOST2085PadraoGMAT().keys()).length}; cookie=${cookie ? "presente" : "ausente"}`
  });

  const resp = await fetch(endpoint, {
    method: "POST",
    redirect: "follow",
    headers: {
      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/vnd.ms-excel,application/octet-stream,*/*;q=0.8",
      "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
      "Cache-Control": "max-age=0",
      "Content-Type": "application/x-www-form-urlencoded",
      "Cookie": cookie,
      "Origin": "https://gmat.procempa.com.br",
      "Referer": referer,
      "Sec-Fetch-Dest": "frame",
      "Sec-Fetch-Mode": "navigate",
      "Sec-Fetch-Site": "same-origin",
      "Sec-Fetch-User": "?1",
      "Upgrade-Insecure-Requests": "1",
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36"
    },
    body
  });

  const headers = {};
  resp.headers.forEach((v, k) => headers[k] = v);
  const ab = await resp.arrayBuffer();
  const contentType = String(headers["content-type"] || "");
  const contentDisposition = String(headers["content-disposition"] || "");

  etapas.push({
    etapa: "post xls direto resposta",
    em: new Date().toISOString(),
    detalhe: `HTTP ${resp.status}; bytes=${ab ? ab.byteLength : 0}; content-type=${contentType}; content-disposition=${contentDisposition}; url=${String(resp.url || "").slice(0, 500)}`
  });

  const avaliacao = capturaParecePlanilhaGMAT(ab, headers, resp.url || endpoint);

  if (!resp.ok || !ab || !ab.byteLength || !avaliacao.ok) {
    etapas.push({
      etapa: "post xls direto não capturou arquivo",
      em: new Date().toISOString(),
      detalhe: `Resposta não parece planilha real. magic=${magicArquivoGMAT(ab)}; motivo=${avaliacao.motivo}; amostra=${String(avaliacao.amostra||"").slice(0, 1000)}`
    });
    return null;
  }

  return {
    url: resp.url,
    status: resp.status,
    headers,
    arrayBuffer: ab,
    arquivoNome: parseFileName(headers) || "PlanilhaMateriais.xls",
    origem: "post-xls-direto-v137"
  };
}


async function submeterPOST2085ComoFormularioReal(page, browser, etapas) {
  // GMAT legado: o download real ocorre por navegação de formulário, não por fetch.
  // Esta rotina cria um formulário POST real dentro do contexto do GMAT e submete
  // para um target oculto, capturando a resposta de rede binária.
  const endpoint = "https://gmat.procempa.com.br/gmat/uc2085/gerarInformacoesPlanilhaPesquisa.do";
  const params = paramsPOST2085PadraoGMAT();

  let capturado = null;
  const paginas = await browser.pages().catch(() => [page]);

  async function anexar(pg, origem) {
    try {
      pg.on("response", async (response) => {
        try {
          if (capturado) return;
          const url = String(response.url() || "");
          if (!url.includes("/gmat/uc2085/gerarInformacoesPlanilhaPesquisa.do")) return;

          const headers = response.headers();
          const ab = await response.arrayBuffer();
          const magic = ab && ab.byteLength >= 8
            ? Array.from(new Uint8Array(ab.slice(0, 8))).map(b => b.toString(16).padStart(2, "0")).join(" ")
            : "";

          const ct = String(headers["content-type"] || "");
          const cd = String(headers["content-disposition"] || "");
          const xlsBinario = magic.startsWith("d0 cf 11 e0");
          const pareceArquivo = xlsBinario ||
            isExcelContentType(ct) ||
            /attachment|xls|xlsx|octet/i.test(cd) ||
            (ab && ab.byteLength > 80000 && !ct.toLowerCase().includes("text/html"));

          etapas.push({
            etapa: "form post navegação resposta",
            em: new Date().toISOString(),
            detalhe: `origem=${origem}; HTTP ${response.status()}; bytes=${ab ? ab.byteLength : 0}; magic=${magic}; content-type=${ct}; content-disposition=${cd}; url=${url.slice(0, 700)}`
          });

          if (response.ok() && pareceArquivo && ab && ab.byteLength) {
            capturado = {
              url,
              status: response.status(),
              headers,
              arrayBuffer: ab,
              arquivoNome: parseFileName(headers) || "PlanilhaMateriais.xls",
              origem: xlsBinario ? "form-post-navegacao-xls-binario-v138" : "form-post-navegacao-v138"
            };
          }
        } catch (e) {}
      });
    } catch (e) {}
  }

  for (const pg of paginas) await anexar(pg, "pagina-existente");

  try {
    browser.on("targetcreated", async (target) => {
      try {
        const pg = await target.page();
        if (pg) await anexar(pg, "targetcreated");
      } catch (e) {}
    });
  } catch (e) {}

  etapas.push({
    etapa: "form post navegação",
    em: new Date().toISOString(),
    detalhe: `Criando formulário POST real para endpoint 2085 com ${Array.from(params.keys()).length} campos.`
  });

  // Executa dentro de um frame/página GMAT. Preferir frame uc2085; senão página principal.
  let frameAlvo = allFrames(page).find(f => String(f.url() || "").includes("uc2085")) ||
                  allFrames(page).find(f => String(f.url() || "").includes("gmat.procempa.com.br/gmat")) ||
                  page.mainFrame();

  await frameAlvo.evaluate((endpoint, entries) => {
    const old = document.getElementById("__portal_da_form_2085_v138");
    if (old) old.remove();

    let iframe = document.getElementById("__portal_da_target_2085_v138");
    if (!iframe) {
      iframe = document.createElement("iframe");
      iframe.name = "__portal_da_target_2085_v138";
      iframe.id = "__portal_da_target_2085_v138";
      iframe.style.position = "fixed";
      iframe.style.left = "-9999px";
      iframe.style.top = "-9999px";
      iframe.style.width = "10px";
      iframe.style.height = "10px";
      document.body.appendChild(iframe);
    }

    const form = document.createElement("form");
    form.id = "__portal_da_form_2085_v138";
    form.method = "POST";
    form.action = endpoint;
    form.target = "__portal_da_target_2085_v138";
    form.enctype = "application/x-www-form-urlencoded";
    form.acceptCharset = "UTF-8";

    for (const [k, v] of entries) {
      const input = document.createElement("input");
      input.type = "hidden";
      input.name = k;
      input.value = v;
      form.appendChild(input);
    }

    document.body.appendChild(form);
    form.submit();
  }, endpoint, Array.from(params.entries()));

  const started = Date.now();
  while (!capturado && Date.now() - started < 60000) {
    await wait(500);
  }

  if (capturado) return capturado;

  etapas.push({
    etapa: "form post navegação não capturou",
    em: new Date().toISOString(),
    detalhe: "Formulário real foi submetido, mas nenhuma resposta reconhecida como XLS binário/arquivo foi capturada em 60s."
  });

  return null;
}


export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (request.method === "GET") {
      return json({
        ok: true,
        servico: "Robô GMAT CMAP",
        versao: "v138",
        navegadorConfigurado: !!getBrowserBinding(env),
        urlGMATConfigurada: !!env.CMAP_GMAT_URL,
        usuarioConfigurado: !!env.CMAP_GMAT_USUARIO,
        senhaConfigurada: !!env.CMAP_GMAT_SENHA,
        tokenConfigurado: !!env.CMAP_GMAT_ROBO_TOKEN,
        endpoints: {
          executar: "POST /",
          diagnostico: "GET /"
        }
      });
    }

    if (request.method !== "POST") {
      return json({ ok: false, erro: "Método não permitido." }, 405);
    }

    return atualizarEstoqueGMAT(request, env);
  }
};
