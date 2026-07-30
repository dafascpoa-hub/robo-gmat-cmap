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

async function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function clickByText(page, text, options = {}) {
  const timeoutMs = options.timeout || 15000;
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const ok = await page.evaluate((targetText) => {
      const normalize = (s) => String(s || "").replace(/\s+/g, " ").trim();
      const wanted = normalize(targetText);
      const candidates = Array.from(document.querySelectorAll("td,div,a,span,button,input,li"));
      const el = candidates.find(e => {
        const visible = !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
        const txt = normalize(e.innerText || e.value || e.textContent);
        return visible && txt.includes(wanted);
      });
      if (!el) return false;
      el.scrollIntoView({ block: "center", inline: "center" });
      el.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
      el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
      el.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
      el.click();
      return true;
    }, text);
    if (ok) return true;
    await wait(300);
  }
  throw new Error(`Elemento com texto não encontrado: ${text}`);
}

async function hoverByText(page, text, options = {}) {
  const timeoutMs = options.timeout || 15000;
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const box = await page.evaluate((targetText) => {
      const normalize = (s) => String(s || "").replace(/\s+/g, " ").trim();
      const wanted = normalize(targetText);
      const candidates = Array.from(document.querySelectorAll("td,div,a,span,button,input,li"));
      const el = candidates.find(e => {
        const visible = !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
        const txt = normalize(e.innerText || e.value || e.textContent);
        return visible && txt.includes(wanted);
      });
      if (!el) return null;
      el.scrollIntoView({ block: "center", inline: "center" });
      const r = el.getBoundingClientRect();
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    }, text);
    if (box) {
      await page.mouse.move(box.x, box.y);
      await wait(600);
      return true;
    }
    await wait(300);
  }
  throw new Error(`Elemento para hover não encontrado: ${text}`);
}

async function waitText(page, text, timeout = 20000) {
  await page.waitForFunction(
    (targetText) => document.body && document.body.innerText && document.body.innerText.includes(targetText),
    { timeout },
    text
  );
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
    await waitText(page, perfil, 30000);
    await Promise.allSettled([
      clickByText(page, perfil, { timeout: 20000 }),
      page.waitForNavigation({ waitUntil: "networkidle0", timeout: 30000 })
    ]);

    log("abrindo menu relatórios");
    await wait(900);
    await hoverByText(page, "Relatórios", { timeout: 20000 });
    await wait(500);

    log("selecionando relatório 2085");
    await clickByText(page, relatorio, { timeout: 20000 });
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
        const shot = await page.screenshot({ encoding: "base64", fullPage: false });
        screenshotBase64 = shot || "";
      }
    } catch (s) {}

    return json({
      ok: false,
      codigoJob,
      erro: e && e.message ? e.message : String(e),
      etapas,
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
        versao: "v121",
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
