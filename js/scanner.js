const fs = require("fs");
const path = require("path");

function read(file) {
  return fs.existsSync(file) ? fs.readFileSync(file, "utf8") : "";
}

function has(html, pattern) {
  return pattern.test(html);
}

const root = path.resolve(__dirname, "..");
const indexFile = path.join(root, "index.html");
const appFile = path.join(root, "js", "app.js");
const cssFile = path.join(root, "css", "styles.css");
const configFile = path.join(root, "js", "supabase-config.js");

const indexHtml = read(indexFile);
const appJs = read(appFile);
const css = read(cssFile);
const report = [];

report.push({ ok: has(indexHtml, /id="loginForm"/), check: "Tela de login", detail: "#loginForm" });
report.push({ ok: has(indexHtml, /id="filtersToggleBtn"/) && has(indexHtml, /id="advancedFilters"/) && /function setAdvancedFiltersOpen/.test(appJs) && /filter-search-row/.test(css), check: "Filtros expansíveis", detail: "Busca fixa com filtros recolhíveis" });
report.push({ ok: has(indexHtml, /id="saleModal"/) && has(indexHtml, /id="newSaleBtn"/) && /function openSaleModal/.test(appJs) && appJs.includes("Editar pago mensal") && appJs.includes("Editar venda/aditivo") && /closeSaleModal\(false\)/.test(appJs) && /closeSaleModal\(false\)/.test(appJs), check: "Modal de vendas", detail: "Vendas/Pagos em modal" });
report.push({ ok: has(indexHtml, /id="clientModal"/) && has(indexHtml, /id="newClientBtn"/) && /function openClientModal/.test(appJs) && appJs.includes("Editar cliente"), check: "Modal de cliente", detail: "Cadastro/edição em modal" });
report.push({ ok: has(indexHtml, /id="cancelForm"/), check: "Formulário de cancelamento", detail: "#cancelForm" });
report.push({ ok: has(indexHtml, /id="cancelModal"/) && has(indexHtml, /id="newCancelBtn"/) && /function openCancelModal/.test(appJs), check: "Modal de cancelamento", detail: "Cadastro/edição em modal" });
report.push({ ok: has(indexHtml, /data-view="resumo"/) && has(indexHtml, /id="summaryTable"/) && /function renderMonthlySummary/.test(appJs) && /renderMonthlySummary\(\);/.test(appJs), check: "Resumo mensal", detail: "Aba e tabela consolidadas" });
report.push({ ok: has(indexHtml, /id="clearCancelsBtn"/), check: "Botão Limpar tudo em cancelamentos", detail: "#clearCancelsBtn" });
report.push({ ok: has(indexHtml, /<th[^>]*>Ações<\/th>/), check: "Coluna Ações", detail: "Cabeçalho da tabela" });
report.push({ ok: /data-cancel-edit/.test(appJs) && /data-cancel-delete/.test(appJs), check: "Editar/Excluir cancelamentos", detail: "Eventos no js/app.js" });
report.push({ ok: /data-sale-edit/.test(appJs) && /data-sale-delete/.test(appJs), check: "Editar/Excluir vendas", detail: "Eventos no js/app.js" });
report.push({ ok: /data-receivable-edit/.test(appJs) && /data-receivable-delete/.test(appJs), check: "Editar/Excluir pagos mensais", detail: "Eventos no js/app.js" });
report.push({ ok: /function canonicalMonth/.test(appJs), check: "Normalização de meses", detail: "Aceita mês completo e abreviado" });
report.push({ ok: /option:not\(:first-child\)/.test(appJs), check: "Filtro de mês sem duplicação", detail: "Limpa opções antes de recriar" });
report.push({ ok: fs.existsSync(configFile), check: "Configuração Supabase", detail: "js/supabase-config.js" });
report.push({ ok: has(indexHtml, /id="loadingScreen"/) && /body\.loading \.loading-screen/.test(css), check: "Tela de carregamento", detail: "Aguarda banco antes de exibir painel" });
report.push({ ok: !/conectel-data\.js/.test(indexHtml), check: "Sem carga local de dados", detail: "index.html não carrega data/conectel-data.js" });
report.push({ ok: /ghost-btn\.danger/.test(css), check: "Estilo do botão de perigo", detail: "css/styles.css" });
report.push({ ok: /\[hidden\]\s*\{\s*display: none !important;/.test(css) && /function openCancelModal\(title = "Novo cancelamento"\) \{\s*if \(!canManage\(\)\) return;/.test(appJs) && /function editCancel\(key\) \{\s*if \(!canManage\(\)\) return;/.test(appJs), check: "Bloqueio visualizador", detail: "Ações admin escondidas e modais bloqueadas" });

const encodingSuspects = [];
[["index.html", indexHtml], ["js/app.js", appJs], ["css/styles.css", css]].forEach(([name, content]) => {
  const checkContent = name === "js/app.js" ? content.replace(/if \(!\/�\/\.test\(utf8\)\) return utf8;/g, "") : content;
  if (/Ã[^\sA-Z]|Â|�/.test(checkContent)) encodingSuspects.push(name);
});
report.push({ ok: encodingSuspects.length === 0, check: "Acentuação visível", detail: encodingSuspects.length ? encodingSuspects.join(", ") : "OK" });

console.log("Relatório rápido de verificação:");
let failed = 0;
report.forEach((item) => {
  const status = item.ok ? "[OK] " : "[WARN]";
  if (!item.ok) failed++;
  console.log(`${status} ${item.check} - ${item.detail}`);
});

if (failed) {
  console.log("\nExistem avisos para revisar antes de publicar.");
  process.exit(1);
}
console.log("\nTudo básico parece OK.");




