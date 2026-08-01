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

    const saiu = await page.evaluate(() => {
      const txt = document.body ? document.body.innerText || "" : "";
      return txt.includes("Geração de Informações de Materiais em Planilha") ||
             txt.includes("Critérios de Pesquisa") ||
             txt.includes("Gerar Planilha") ||
             txt.includes("Almoxarifado");
    }).catch(() => false);

    if (saiu) return true;
  }
  return false;
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

    page.on("response", async (response) => {
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
          arquivoNome: parseFileName(headers)
        };
      } catch (e) {}
    });

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
        const okCoordItem = await clickGMATRelatorio2085PorCoordenada(page);
        if (!okCoordItem) {
          throw new Error("Não foi possível clicar no relatório 2085 por texto nem por coordenada. " + (e2 && e2.message ? e2.message : String(e2)));
        }
      }
    }

    await Promise.allSettled([
      page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 }),
      waitText(page, "Gerar Planilha", 30000)
    ]);

    log("gerando planilha");
    await page.waitForSelector('input.buttonToolbar[value="Gerar Planilha"], input[name="submitAction"]', { timeout: 30000 });
    await Promise.allSettled([
      page.click('input.buttonToolbar[value="Gerar Planilha"], input[name="submitAction"]'),
      page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 })
    ]);

    const started = Date.now();
    while (!captured && Date.now() - started < 45000) {
      await wait(500);
    }

    if (!captured) {
      const title = await page.title().catch(() => "");
      const text = await page.evaluate(() => document.body ? document.body.innerText.slice(0, 1000) : "").catch(() => "");
      throw new Error("A planilha não foi capturada. Última página: " + title + " | " + text.slice(0, 300));
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

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (request.method === "GET") {
      return json({
        ok: true,
        servico: "Robô GMAT CMAP",
        versao: "v129",
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
