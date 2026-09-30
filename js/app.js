(function () {
  const supabaseSettings = window.CONNECTEL_SUPABASE || {};
  const supabaseClient = window.supabase && supabaseSettings.url && supabaseSettings.anonKey
    ? window.supabase.createClient(supabaseSettings.url, supabaseSettings.anonKey)
    : null;
  const raw = {};
  const manualKey = "conectel.manual.clients.v1";
  const manualSalesKey = "conectel.manual.sales.v2";
  const manualReceivablesKey = "conectel.manual.receivables.v2";
  const manualCancelsKey = "conectel.manual.cancels.v2";
  const salesEditsKey = "conectel.sales.edits.v2";
  const receivableEditsKey = "conectel.receivable.edits.v2";
  const cancelEditsKey = "conectel.cancel.edits.v2";
  const receivableStatusKey = "conectel.receivable.status.v1";
  const overrideKey = "conectel.client.overrides.v1";
  const authKey = "conectel.session.v1";
  const money = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
  const users = {
    admin: { password: "admin123", role: "admin", label: "Admin" },
    visualizador: { password: "visualizar123", role: "viewer", label: "Visualizador" }
  };

  const state = {
    view: "dashboard",
    search: "",
    status: "",
    month: "",
    system: "",
    service: "",
    finance: "",
    consultant: "",
    reason: "",
    ranking: 10,
    salesMode: "vendas",
    salesMonth: "",
    cancelMonth: "",
    editingKey: "",
    editingCancelKey: "",
    editingSaleKey: "",
    editingReceivableKey: "",
    session: readSession(),
    overrides: supabaseClient ? {} : readOverrides(),
    receivableStatus: supabaseClient ? {} : readReceivableStatus(),
    manualSales: supabaseClient ? [] : readManualSales(),
    manualReceivables: supabaseClient ? [] : readManualReceivables(),
    manualCancels: supabaseClient ? [] : readManualCancels(),
    salesEdits: supabaseClient ? {} : readStoredMap(salesEditsKey),
    receivableEdits: supabaseClient ? {} : readStoredMap(receivableEditsKey),
    cancelEdits: supabaseClient ? {} : readStoredMap(cancelEditsKey),
    showDuplicates: false,
    manual: supabaseClient ? [] : readManual()
  };


  const monthOrder = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
  ];

  function normalizeText(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .trim();
  }

  function canonicalMonth(value) {
    const text = normalizeText(value);
    if (!text) return "";
    const aliases = {
      JAN: "JANEIRO",
      JANEIRO: "JANEIRO",
      FEV: "FEVEREIRO",
      FEVEREIRO: "FEVEREIRO",
      MAR: "MARÇO",
      MARCO: "MARÇO",
      MARÇO: "MARÇO",
      ABR: "ABRIL",
      ABRIL: "ABRIL",
      MAI: "MAIO",
      MAIO: "MAIO",
      JUN: "JUNHO",
      JUNHO: "JUNHO",
      JUL: "JULHO",
      JULHO: "JULHO",
      AGO: "AGOSTO",
      AGOSTO: "AGOSTO",
      SET: "SETEMBRO",
      SETEMBRO: "SETEMBRO",
      OUT: "OUTUBRO",
      OUTUBRO: "OUTUBRO",
      NOV: "NOVEMBRO",
      NOVEMBRO: "NOVEMBRO",
      DEZ: "DEZEMBRO",
      DEZEMBRO: "DEZEMBRO"
    };
    const token = text.split(/[\/\-\s]+/).find(Boolean) || text;
    return aliases[text] || aliases[token] || aliases[token.slice(0, 3)] || text;
  }
  function numberValue(value) {
    if (value === null || value === undefined || value === "") return 0;
    if (typeof value === "number") return value;
    const rawValue = String(value).trim().replace(/[^\d,.-]/g, "");
    if (!rawValue) return 0;
    if (rawValue.includes(",")) {
      return Number(rawValue.replace(/\./g, "").replace(",", ".")) || 0;
    }
    return Number(rawValue) || 0;
  }

  function formatMoney(value) {
    return money.format(numberValue(value));
  }

  function formatCompactMoney(value) {
    const amount = numberValue(value);
    if (Math.abs(amount) < 1000000) return formatMoney(amount);
    return new Intl.NumberFormat("pt-BR", {
      style: "currency",
      currency: "BRL",
      notation: "compact",
      maximumFractionDigits: 1
    }).format(amount);
  }

  function excelDate(value) {
    const serial = Number(value);
    if (!Number.isFinite(serial) || serial <= 20000) return value || "";
    const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    return date.toLocaleDateString("pt-BR", { timeZone: "UTC" });
  }

  function toKey(header, index) {
    const base = normalizeText(header || `COLUNA_${index + 1}`)
      .replace(/[^A-Z0-9]+/g, "_")
      .replace(/^_|_$/g, "") || `COLUNA_${index + 1}`;
    return base;
  }

  function rowsToObjects(sheet, fallbackHeaders) {
    const rows = sheet?.rows || [];
    if (!rows.length) return [];
    const firstRow = rows[0] || [];
    const hasHeader = firstRow.some((cell) => /EMPRESA|VALOR|STATUS|CONSULTOR|BASE/i.test(String(cell || "")));
    const headers = hasHeader ? firstRow : fallbackHeaders;
    const start = hasHeader ? 1 : 0;
    const seen = {};
    const keys = headers.map((header, index) => {
      const key = toKey(header, index);
      seen[key] = (seen[key] || 0) + 1;
      return seen[key] > 1 ? `${key}_${seen[key]}` : key;
    });

    return rows.slice(start).map((row) => {
      const obj = {};
      keys.forEach((key, index) => {
        obj[key] = row[index] ?? "";
      });
      obj._sheet = sheet.name || "";
      return obj;
    }).filter((item) => Object.values(item).some((value) => value !== ""));
  }

  function asArray(value) {
    if (!value) return [];
    return Array.isArray(value) ? value : [value];
  }

  const baseSheets = asArray(raw["BASE-CONECTEL"]);
  const cancelSheets = asArray(raw["CANCELADOS-2026"]);
  const salesSheets = asArray(raw.VENDAS);

  const importedBaseRows = baseSheets
    .flatMap((sheet) => rowsToObjects(sheet).map((item) => ({
      ...item,
      MES_REFERENCIA: canonicalMonth(getField(item, ["MES_REFERENCIA", "MES", "M_S", "MES_", "ENTRADA"]) || sheet.name)
    })));
  let baseRows = buildBaseRows();
  const blockRows = [];
  const baseCancelRows = [];
  const manualCancelRows = state.manualCancels.map((item, index) => normalizeManualCancel(item, index));
  const cancelRows = cancelSheets
    .flatMap((sheet) => rowsToObjects(sheet))
    .filter((item) => monthIndex(getItemMonth(item)) >= monthIndex("JUNHO"))
    .concat(manualCancelRows);
  function rowsFromSalesSheets(predicate, typeLabel, statusLabel) {
    return salesSheets.filter(predicate).flatMap((sheet) => rowsToObjects(sheet).map((item) => ({
      ...item,
      ORIGEM: sheet.name || "Planilha",
      TIPO_REGISTRO: typeLabel,
      STATUS_FINANCEIRO: statusLabel
    })));
  }

  const manualSalesRows = state.manualSales.map((item, index) => normalizeManualSale(item, index));
  const manualReceivableRows = state.manualReceivables.map((item, index) => normalizeManualReceivable(item, index));
  const salesMonthlyRows = rowsFromSalesSheets((sheet) => normalizeText(sheet.name) !== "PAGOS", "Venda/aditivo", "Venda").concat(manualSalesRows);
  const paidMonthlyRows = rowsFromSalesSheets((sheet) => normalizeText(sheet.name) === "PAGOS", "Mensalidade", "Pago").concat(manualReceivableRows);
  const salesRows = salesSheets.flatMap((sheet) => rowsToObjects(sheet).map((item) => ({
    ...item,
    ORIGEM: sheet.name || "Planilha"
  }))).concat(manualSalesRows);

  let clientStatusCancelRows = buildClientStatusCancelRows();
  let baseMonthlyReceivableRows = buildBaseMonthlyReceivableRows();

  function buildClientStatusCancelRows() {
    return [];
  }

  function buildBaseMonthlyReceivableRows() {
    return dedupeClientsByMonth(baseRows).filter((item) => !isCanceledClient(item)).filter((item) => !hasCancelRecord(item)).map((item) => ({
      ...item,
      CONSULTOR_CANONICO: "",
      MES: getItemMonth(item),
      TIPO_REGISTRO: "Mensalidade",
      STATUS_FINANCEIRO: monthlyFinanceStatus(item),
      VENDA: getField(item, ["SERVICO", "SERVI_O"]),
      ORIGEM: "Carteira"
    }));
  }

  const overdueMonthlyRows = baseRows
    .filter((item) => normalizeText(getField(item, ["FINANCEIRO", "FINANCEIRO_2"])).includes("ATRASO"))
    .map((item) => ({
      ...item,
      CONSULTOR_CANONICO: "",
      MES: getItemMonth(item),
      TIPO_REGISTRO: "Mensalidade",
      STATUS_FINANCEIRO: "Atraso",
      VENDA: getField(item, ["SERVICO", "SERVI_O"])
    }));

  salesRows.concat(salesMonthlyRows, paidMonthlyRows, cancelRows, clientStatusCancelRows).forEach((item) => {
    item.CONSULTOR_CANONICO = canonicalConsultant(getField(item, ["CONSULTOR"]));
  });

  function isTotalSalesRow(item) {
    return normalizeText(getField(item, ["CONSULTOR"])) === "CONECTEL";
  }

  function individualSalesRows(rows) {
    return rows.filter((item) => !isTotalSalesRow(item));
  }

  function totalSalesRows(rows) {
    return individualSalesRows(rows);
  }

  function canManage() {
    return state.session?.role === "admin";
  }

  function canonicalFinanceStatus(value) {
    const text = normalizeText(value);
    if (text.includes("PAGO")) return "Pago";
    if (text.includes("ATRASO")) return "Atraso";
    if (text.includes("PENDENTE") || text.includes("RECEBER")) return "Pendente";
    if (text.includes("VENDA")) return "Venda";
    return value || "Pendente";
  }

  function normalizeFinanceRows(rows) {
    return rows.map((item) => ({
      ...item,
      STATUS_FINANCEIRO: canonicalFinanceStatus(getField(item, ["STATUS_FINANCEIRO", "FINANCEIRO", "FINANCEIRO_2"]))
    }));
  }
  function monthlyFinanceStatus(item) {
    const finance = normalizeText(getField(item, ["FINANCEIRO", "FINANCEIRO_2"]));
    if (finance.includes("PAGO")) return "Pago";
    if (finance.includes("ATRASO")) return "Atraso";
    return "Pendente";
  }

  function isCanceledClient(item) {
    const status = normalizeText(getField(item, ["STATUS"]));
    return status.includes("CANCEL");
  }

  function isDowngradeClient(item) {
    const status = normalizeText(getField(item, ["STATUS"]));
    return status.includes("DOWNGRADE");
  }

  function isCancelOrDowngradeClient(item) {
    return isCanceledClient(item) || isDowngradeClient(item);
  }

  function isCancellationRecord(item) {
    const status = normalizeText(getField(item, ["STATUS", "MOTIVO"]));
    return status.includes("CANCEL");
  }

  function isDowngradeRecord(item) {
    const status = normalizeText(getField(item, ["STATUS", "MOTIVO"]));
    return status.includes("DOWNGRADE");
  }

  function hasCancelRecord(item) {
    const company = normalizeText(getField(item, ["EMPRESA"]));
    const month = getItemMonth(item);
    if (!company || !month) return false;
    return cancelRows.concat(clientStatusCancelRows, baseCancelRows).some((row) => (
      isCancellationRecord(row) &&
      normalizeText(getField(row, ["EMPRESA"])) === company &&
      getItemMonth(row) === month
    ));
  }

  function allCancelRows() {
    return applyEdits(cancelRows.concat(clientStatusCancelRows, baseCancelRows), cancelKey, state.cancelEdits, "_cancelKey");
  }

  function clientKey(item) {
    return `${normalizeText(getField(item, ["EMPRESA"]))}::${getItemMonth(item) || "ATUAL"}`;
  }

  function normalizeManualClient(item) {
    const normalized = {
      ...item,
      MES_REFERENCIA: item.MES_REFERENCIA || state.month || "JULHO",
      _manual: true
    };
    normalized._rowKey = normalized._rowKey || clientKey(normalized) || `manual-${Date.now()}`;
    return normalized;
  }

  function buildBaseRows() {
    const imported = importedBaseRows
      .map((item) => {
        const key = item._rowKey || clientKey(item);
        const override = state.overrides[key];
        if (override?.deleted) return null;
        return { ...item, ...(override?.data || {}), _rowKey: key };
      })
      .filter(Boolean);
    const manual = state.manual.map(normalizeManualClient).filter((item) => !item._deleted);
    return imported.concat(manual);
  }

  function refreshBaseRows() {
    baseRows = buildBaseRows();
    clientStatusCancelRows = buildClientStatusCancelRows();
    baseMonthlyReceivableRows = buildBaseMonthlyReceivableRows();
  }

  function dedupeClients(rows) {
    const map = new Map();
    rows.forEach((item) => {
      const key = normalizeText(getField(item, ["EMPRESA"]));
      if (!key) return;
      map.set(key, item);
    });
    return [...map.values()];
  }

  function dedupeClientsByMonth(rows) {
    const map = new Map();
    rows.forEach((item) => {
      const company = normalizeText(getField(item, ["EMPRESA"]));
      const month = getItemMonth(item);
      if (!company || !month) return;
      map.set(`${company}::${month}`, item);
    });
    return [...map.values()];
  }


  function stableRowKey(prefix, item, index = 0) {
    return `${prefix}-${normalizeText(getField(item, ["EMPRESA"]))}-${getItemMonth(item) || "MES"}-${normalizeText(getField(item, ["VENDA", "SERVICO", "MOTIVO", "RECORRENCIA"]))}-${numberValue(getField(item, ["VALOR"]))}-${index}`;
  }

  function saleKey(item, index = 0) {
    return item._saleKey || stableRowKey("sale", item, index);
  }

  function normalizeManualSale(item, index = 0) {
    const normalized = {
      ...item,
      CONSULTOR_CANONICO: canonicalConsultant(getField(item, ["CONSULTOR_CANONICO", "CONSULTOR"])),
      ORIGEM: item.ORIGEM || "Manual",
      TIPO_REGISTRO: item.TIPO_REGISTRO || "Venda/aditivo",
      STATUS_FINANCEIRO: item.STATUS_FINANCEIRO || "Venda",
      _manualSale: true,
      _saleKey: item._saleKey || `sale-${Date.now()}-${Math.random().toString(16).slice(2)}`
    };
    normalized._saleKey = saleKey(normalized, index);
    return normalized;
  }

  function normalizeManualReceivable(item, index = 0) {
    const normalized = {
      ...item,
      CONSULTOR_CANONICO: canonicalConsultant(getField(item, ["CONSULTOR_CANONICO", "CONSULTOR"])),
      ORIGEM: item.ORIGEM || "Manual",
      TIPO_REGISTRO: item.TIPO_REGISTRO || "Mensalidade",
      STATUS_FINANCEIRO: item.STATUS_FINANCEIRO || "Pendente",
      _manualReceivable: true
    };
    normalized._receivableKey = receivableKey(normalized) || stableRowKey("receivable", normalized, index);
    return normalized;
  }

  function applyEdits(rows, keyFn, editMap, keyName) {
    return rows.map((item, index) => {
      const key = keyFn(item, index);
      const saved = editMap[key];
      if (saved?.deleted) return null;
      return { ...item, ...(saved?.data || {}), [keyName]: key };
    }).filter(Boolean);
  }

  function saveEdits() {
    if (supabaseClient) return;
    localStorage.setItem(salesEditsKey, JSON.stringify(state.salesEdits));
    localStorage.setItem(receivableEditsKey, JSON.stringify(state.receivableEdits));
    localStorage.setItem(cancelEditsKey, JSON.stringify(state.cancelEdits));
  }
  function monthlyKey(item) {
    return `${normalizeText(getField(item, ["EMPRESA"]))}::${getItemMonth(item)}`;
  }

  function receivableKey(item) {
    return item._receivableKey || monthlyKey(item);
  }

  function applySavedReceivableStatus(rows) {
    return rows.map((item) => {
      const savedStatus = state.receivableStatus[receivableKey(item)];
      return savedStatus ? { ...item, STATUS_FINANCEIRO: savedStatus } : item;
    });
  }

  function downgradeAmountFor(item) {
    const company = normalizeText(getField(item, ["EMPRESA"]));
    const month = getItemMonth(item);
    if (!company || !month) return 0;
    return cancelRows.concat(baseCancelRows)
      .filter((row) => isDowngradeRecord(row))
      .filter((row) => normalizeText(getField(row, ["EMPRESA"])) === company && getItemMonth(row) === month)
      .reduce((total, row) => total + numberValue(getField(row, ["VALOR"])), 0);
  }

  function applyMonthlyAdjustments(rows) {
    const usedDiscounts = new Set();
    return rows.map((item) => {
      const key = monthlyKey(item);
      const discount = usedDiscounts.has(key) ? 0 : downgradeAmountFor(item);
      if (!discount) return item;
      usedDiscounts.add(key);
      const originalValue = numberValue(getField(item, ["VALOR"]));
      const adjustedValue = Math.max(0, originalValue - discount);
      return {
        ...item,
        VALOR: adjustedValue,
        STATUS_FINANCEIRO: getField(item, ["STATUS_FINANCEIRO"]) || monthlyFinanceStatus(item),
        TIPO_REGISTRO: `${getField(item, ["TIPO_REGISTRO"]) || "Mensalidade"} com downgrade`,
        VENDA: `${getField(item, ["VENDA", "SERVICO", "SERVI_O"]) || "Mensalidade"} - downgrade ${formatMoney(discount)}`
      };
    });
  }

  function projectedReceivableMonths() {
    const latestMonth = latestBaseMonth();
    const latestIndex = monthIndex(latestMonth);
    const months = new Set();
    paidMonthlyRows.concat(salesMonthlyRows).forEach((item) => {
      const month = getItemMonth(item);
      if (monthIndex(month) > latestIndex) months.add(month);
    });
    if (state.salesMonth && monthIndex(state.salesMonth) > latestIndex) months.add(state.salesMonth);
    return [...months].filter((month) => !baseRows.some((item) => getItemMonth(item) === month));
  }

  function buildProjectedMonthlyReceivableRows() {
    const latestMonth = latestBaseMonth();
    if (!latestMonth) return [];
    const sourceRows = dedupeClients(baseRows.filter((item) => getItemMonth(item) === latestMonth))
      .filter((item) => !isCanceledClient(item))
      .filter((item) => !hasCancelRecord(item));
    return projectedReceivableMonths().flatMap((month) => (
      sourceRows
        .map((item) => ({
          ...item,
          MES: month,
          MES_REFERENCIA: month,
          TIPO_REGISTRO: "Mensalidade projetada",
          STATUS_FINANCEIRO: "Pendente",
          ORIGEM: `Espelho de ${monthLabel(latestMonth)}`,
          VENDA: getField(item, ["SERVICO", "SERVI_O"])
        }))
        .filter((item) => !hasCancelRecord(item))
    ));
  }

  function latestBaseMonth() {
    return baseRows
      .map((item) => getItemMonth(item))
      .filter(Boolean)
      .sort((a, b) => monthIndex(b) - monthIndex(a))[0] || "";
  }

  function scopedBaseRows() {
    const latestMonth = latestBaseMonth();
    const rows = state.month
      ? baseRows.filter((item) => getItemMonth(item) === normalizeText(state.month))
      : baseRows.filter((item) => !latestMonth || getItemMonth(item) === latestMonth);
    return state.showDuplicates ? rows : dedupeClients(rows);
  }

  function readStoredMap(key) {
    try {
      return JSON.parse(localStorage.getItem(key) || "{}");
    } catch (error) {
      return {};
    }
  }

  function readManual() {
    try {
      return JSON.parse(localStorage.getItem(manualKey) || "[]");
    } catch (error) {
      return [];
    }
  }

  function saveManual() {
    if (supabaseClient) return;
    localStorage.setItem(manualKey, JSON.stringify(state.manual));
  }

  function readManualSales() {
    try {
      return JSON.parse(localStorage.getItem(manualSalesKey) || "[]");
    } catch (error) {
      return [];
    }
  }

  function saveManualSales() {
    if (supabaseClient) return;
    localStorage.setItem(manualSalesKey, JSON.stringify(state.manualSales));
  }

  function readManualReceivables() {
    try {
      return JSON.parse(localStorage.getItem(manualReceivablesKey) || "[]");
    } catch (error) {
      return [];
    }
  }

  function saveManualReceivables() {
    if (supabaseClient) return;
    localStorage.setItem(manualReceivablesKey, JSON.stringify(state.manualReceivables));
  }

  function readManualCancels() {
    try {
      return JSON.parse(localStorage.getItem(manualCancelsKey) || "[]");
    } catch (error) {
      return [];
    }
  }

  function saveManualCancels() {
    if (supabaseClient) return;
    localStorage.setItem(manualCancelsKey, JSON.stringify(state.manualCancels));
  }
  function readReceivableStatus() {
    try {
      return JSON.parse(localStorage.getItem(receivableStatusKey) || "{}");
    } catch (error) {
      return {};
    }
  }

  function saveReceivableStatus() {
    if (supabaseClient) return;
    localStorage.setItem(receivableStatusKey, JSON.stringify(state.receivableStatus));
  }

  function cancelKey(item, index = 0) {
    return item._cancelKey || `cancel-${normalizeText(getField(item, ["EMPRESA"]))}-${getItemMonth(item) || "MES"}-${index}`;
  }

  function normalizeManualCancel(item, index = 0) {
    return {
      ...item,
      CONSULTOR_CANONICO: canonicalConsultant(getField(item, ["CONSULTOR_CANONICO", "CONSULTOR"]) || "POS"),
      ORIGEM: item.ORIGEM || "Manual",
      _manualCancel: true,
      _cancelKey: cancelKey(item, index)
    };
  }

  function persistManualCancels() {
    state.manualCancels = state.manualCancels.map(normalizeManualCancel);
    saveManualCancels();
  }
  function readOverrides() {
    try {
      return JSON.parse(localStorage.getItem(overrideKey) || "{}");
    } catch (error) {
      return {};
    }
  }

  function saveOverrides() {
    if (supabaseClient) return;
    localStorage.setItem(overrideKey, JSON.stringify(state.overrides));
  }

  function readSession() {
    try {
      localStorage.removeItem(authKey);
      return JSON.parse(sessionStorage.getItem(authKey) || "null");
    } catch (error) {
      return null;
    }
  }

  function saveSession() {
    localStorage.removeItem(authKey);
    if (state.session) {
      sessionStorage.setItem(authKey, JSON.stringify(state.session));
    } else {
      sessionStorage.removeItem(authKey);
    }
  }

  function exportValue(row, fields) {
    if (fields === "MES") return monthLabel(getItemMonth(row));
    if (fields === "VENCIMENTO") return excelDate(getField(row, ["VENCIMENTO"]));
    if (fields === "VALOR") return numberValue(getField(row, ["VALOR"]));
    return getField(row, Array.isArray(fields) ? fields : [fields]);
  }

  function rowsToCsv(rows, columns) {
    const escapeCell = (value) => `"${String(value ?? "").replace(/"/g, '""')}"`;
    const header = columns.map((column) => escapeCell(column.label)).join(";");
    const body = rows.map((row) => columns.map((column) => escapeCell(exportValue(row, column.fields))).join(";"));
    return `\ufeff${[header].concat(body).join("\r\n")}`;
  }

  function downloadFile(name, content, type = "text/csv;charset=utf-8") {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function currentExportDataset() {
    const baseColumns = [
      { label: "Empresa", fields: "EMPRESA" },
      { label: "Filial", fields: "FILIAL" },
      { label: "Sistema", fields: "SISTEMA" },
      { label: "Serviço", fields: ["SERVICO", "SERVI_O"] },
      { label: "Status", fields: "STATUS" },
      { label: "Financeiro", fields: ["FINANCEIRO", "FINANCEIRO_2"] },
      { label: "Valor", fields: "VALOR" },
      { label: "Vencimento", fields: "VENCIMENTO" },
      { label: "Mês", fields: "MES" }
    ];
    if (state.view === "clientes") return { name: "clientes", rows: scopedBaseRows().filter(matchesFilters), columns: baseColumns };
    if (state.view === "financeiro") return { name: "financeiro", rows: scopedBaseRows().concat(blockRows).filter(matchesFilters), columns: baseColumns };
    if (state.view === "vendas") {
      const rows = getSalesWorkspaceRows().filter(matchesFilters).filter((item) => !state.salesMonth || getItemMonth(item) === state.salesMonth);
      return {
        name: state.salesMode === "pagos" ? "pagos-mensais" : "vendas-mensais",
        rows,
        columns: [
          { label: "Tipo", fields: "TIPO_REGISTRO" },
          { label: "Consultor", fields: ["CONSULTOR_CANONICO", "CONSULTOR"] },
          { label: "Mês", fields: "MES" },
          { label: "Status", fields: ["STATUS_FINANCEIRO", "FINANCEIRO"] },
          { label: "Valor", fields: "VALOR" },
          { label: "Empresa", fields: "EMPRESA" },
          { label: "Venda", fields: ["VENDA", "SERVICO", "SERVI_O"] }
        ]
      };
    }
    if (state.view === "resumo") {
      return {
        name: "resumo-mensal",
        rows: buildMonthlySummaryRows(),
        columns: [
          { label: "Mês", fields: "MES" },
          { label: "Clientes", fields: "CLIENTES" },
          { label: "Receita carteira", fields: "RECEITA_CARTEIRA" },
          { label: "Vendas", fields: "VENDAS" },
          { label: "Pagos", fields: "PAGOS" },
          { label: "Em aberto", fields: "ABERTO" },
          { label: "Cancelamentos", fields: "CANCELAMENTOS" },
          { label: "Downgrades", fields: "DOWNGRADES" },
          { label: "Saldo", fields: "SALDO" }
        ]
      };
    }
    if (state.view === "cancelamentos") {
      return {
        name: "cancelamentos",
        rows: allCancelRows().filter(matchesFilters),
        columns: [
          { label: "Empresa", fields: "EMPRESA" },
          { label: "Mês", fields: "MES" },
          { label: "Status", fields: "STATUS" },
          { label: "Motivo", fields: ["MOTIVO", "RECORRENCIA"] },
          { label: "Consultor", fields: ["CONSULTOR_CANONICO", "CONSULTOR"] },
          { label: "Valor", fields: "VALOR" }
        ]
      };
    }
    return { name: "dashboard-clientes", rows: scopedBaseRows().filter(matchesFilters), columns: baseColumns };
  }

  function exportCurrentView() {
    if (state.view === "dashboard") {
      exportDashboardDoc();
      return;
    }
    const dataset = currentExportDataset();
    const date = new Date().toISOString().slice(0, 10);
    downloadFile(`dados-conectel-${dataset.name}-${date}.csv`, rowsToCsv(dataset.rows, dataset.columns));
  }

  function reportRows(title, rows, formatter = (value) => value) {
    const max = Math.max(...rows.map((item) => item.value), 1);
    return `
      <h2>${escapeHtml(title)}</h2>
      <table>
        <thead><tr><th>Item</th><th>Valor</th><th>Gráfico</th></tr></thead>
        <tbody>
          ${rows.map((item) => `
            <tr>
              <td>${escapeHtml(item.label)}</td>
              <td>${escapeHtml(formatter(item.value))}</td>
              <td><div class="bar"><span style="width:${Math.max(4, (item.value / max) * 100)}%"></span></div></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }

  function exportDashboardDoc() {
    const currentBaseRows = scopedBaseRows();
    const filteredBase = currentBaseRows.filter(matchesFilters);
    const filteredSales = individualSalesRows(salesMonthlyRows).filter(matchesFilters);
    const filteredCancels = allCancelRows().filter(matchesFilters);
    const overdue = filteredBase.filter((item) => normalizeText(getField(item, ["FINANCEIRO", "FINANCEIRO_2"])).includes("ATRASO"));
    const metrics = [
      { label: "Clientes na base", value: uniqueCount(filteredBase, "EMPRESA") },
      { label: "Receita ativa", value: sumBy(filteredBase, "VALOR"), money: true },
      { label: "Em atraso", value: sumBy(overdue, "VALOR"), money: true },
      { label: "Vendas", value: sumBy(filteredSales, "VALOR"), money: true },
      { label: "Cancelamentos/downgrades", value: filteredCancels.length }
    ];
    const topClients = filteredBase
      .map((item) => ({ label: getField(item, ["EMPRESA"]) || "Não informado", value: numberValue(getField(item, ["VALOR"])) }))
      .filter((item) => item.value > 0)
      .sort((a, b) => b.value - a.value)
      .slice(0, state.ranking);
    const html = `
      <html>
        <head>
          <meta charset="UTF-8">
          <style>
            body { font-family: Arial, sans-serif; color: #102033; }
            h1 { margin-bottom: 4px; }
            h2 { margin-top: 26px; color: #163b63; }
            table { width: 100%; border-collapse: collapse; margin-top: 10px; }
            th, td { border: 1px solid #c7d6e8; padding: 8px; font-size: 12px; }
            th { background: #eaf4ff; text-align: left; }
            .cards { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; }
            .card { border: 1px solid #c7d6e8; padding: 10px; background: #f7fbff; }
            .card strong { display: block; font-size: 20px; }
            .bar { height: 12px; background: #e8eef6; border-radius: 999px; overflow: hidden; }
            .bar span { display: block; height: 100%; background: #45aeea; }
          </style>
        </head>
        <body>
          <h1>Dashboard Conectel BI Operacional</h1>
          <p>Exportado em ${new Date().toLocaleDateString("pt-BR")}</p>
          <div class="cards">
            ${metrics.map((item) => `<div class="card"><span>${escapeHtml(item.label)}</span><strong>${item.money ? formatMoney(item.value) : item.value}</strong></div>`).join("")}
          </div>
          ${reportRows("Vendas por mês", groupSum(filteredSales, "MES", "VALOR").sort((a, b) => monthIndex(a.label) - monthIndex(b.label)), formatMoney)}
          ${reportRows("Serviços ativos", groupCount(filteredBase, "SERVICO"), String)}
          ${reportRows("Cancelamentos por motivo", groupSum(filteredCancels, "MOTIVO", "VALOR"), formatMoney)}
          ${reportRows("Receita por sistema", groupSum(filteredBase, "SISTEMA", "VALOR"), formatMoney)}
          ${reportRows("Top clientes por valor", topClients, formatMoney)}
        </body>
      </html>
    `;
    const date = new Date().toISOString().slice(0, 10);
    downloadFile(`dados-conectel-dashboard-${date}.doc`, html, "application/msword;charset=utf-8");
  }

  function detectCsvDelimiter(text) {
    const firstLine = String(text || "").split(/\r?\n/).find((line) => line.trim()) || "";
    const candidates = [";", ",", "\t"];
    return candidates
      .map((delimiter) => ({ delimiter, count: firstLine.split(delimiter).length }))
      .sort((a, b) => b.count - a.count)[0]?.delimiter || ",";
  }

  function parseCsv(text) {
    const delimiter = detectCsvDelimiter(text);
    const cleanText = String(text || "").replace(/^\ufeff/, "");
    const rows = [];
    let row = [];
    let cell = "";
    let quoted = false;
    for (let i = 0; i < cleanText.length; i++) {
      const char = cleanText[i];
      const next = cleanText[i + 1];
      if (char === '"' && quoted && next === '"') {
        cell += '"';
        i++;
      } else if (char === '"') {
        quoted = !quoted;
      } else if (char === delimiter && !quoted) {
        row.push(cell.trim());
        cell = "";
      } else if ((char === "\n" || char === "\r") && !quoted) {
        if (char === "\r" && next === "\n") i++;
        row.push(cell.trim());
        if (row.some((value) => String(value).trim())) rows.push(row);
        row = [];
        cell = "";
      } else {
        cell += char;
      }
    }
    row.push(cell.trim());
    if (row.some((value) => String(value).trim())) rows.push(row);
    return rows;
  }

  function decodeFileText(result) {
    if (typeof result === "string") return result;
    const buffer = result instanceof ArrayBuffer ? result : new Uint8Array(result || []).buffer;
    const utf8 = new TextDecoder("utf-8").decode(buffer);
    if (!/�/.test(utf8)) return utf8;
    try {
      return new TextDecoder("windows-1252").decode(buffer);
    } catch (error) {
      return utf8;
    }
  }
  function importedRowsFromText(text, fileName) {
    if (/\.json$/i.test(fileName)) {
      const data = JSON.parse(text);
      return Array.isArray(data) ? data : [];
    }
    const rows = parseCsv(text);
    const headers = rows.shift()?.map((header) => toKey(header)) || [];
    return rows.map((row) => {
      const item = {};
      headers.forEach((header, index) => {
        item[header] = row[index] || "";
      });
      return item;
    });
  }


  function rowsFromMatrix(rows, sheetName = "") {
    const cleanRows = (rows || []).filter((row) => (row || []).some((value) => String(value ?? "").trim()));
    const headers = cleanRows.shift()?.map((header, index) => toKey(header || `COLUNA_${index + 1}`, index)) || [];
    return cleanRows.map((row) => {
      const item = {};
      headers.forEach((header, index) => {
        item[header] = row[index] ?? "";
      });
      item._sheet = sheetName;
      return item;
    });
  }

  function importedRowsFromWorkbook(buffer) {
    if (!window.XLSX) {
      throw new Error("Biblioteca de Excel indisponível. Verifique sua conexão e recarregue a página.");
    }
    const workbook = window.XLSX.read(buffer, { type: "array", cellDates: false });
    return workbook.SheetNames.flatMap((sheetName) => {
      const matrix = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
      return rowsFromMatrix(matrix, sheetName);
    });
  }

  function importedRowsFromFile(result, fileName) {
    if (/\.xlsx?$|\.xls$/i.test(fileName)) return importedRowsFromWorkbook(result);
    return importedRowsFromText(decodeFileText(result), fileName);
  }
  function normalizeImportedClient(item) {
    const client = {
      EMPRESA: getField(item, ["EMPRESA", "NOME", "CLIENTE"]),
      FILIAL: getField(item, ["FILIAL"]),
      SISTEMA: getField(item, ["SISTEMA"]),
      SERVICO: getField(item, ["SERVICO", "SERVI_O", "VENDA"]),
      VALOR: getField(item, ["VALOR"]),
      VENCIMENTO: getField(item, ["VENCIMENTO", "DATA"]),
      STATUS: getField(item, ["STATUS"]) || "ATIVO",
      FINANCEIRO: getField(item, ["FINANCEIRO", "COBRANCA", "STATUS_FINANCEIRO"]) || "RECEBER",
      MES_REFERENCIA: getField(item, ["MES_REFERENCIA", "MES", "M_S", "MES_"]) || state.month || latestBaseMonth() || "JULHO",
      _manual: true
    };
    return normalizeManualClient(client);
  }

  function normalizeImportedSale(item) {
    return {
      CONSULTOR: getField(item, ["CONSULTOR"]) || "Não informado",
      CONSULTOR_CANONICO: canonicalConsultant(getField(item, ["CONSULTOR"])),
      MES: getField(item, ["MES", "M_S", "MES_"]) || state.salesMonth || state.month || latestBaseMonth() || "JULHO",
      VALOR: getField(item, ["VALOR"]),
      EMPRESA: getField(item, ["EMPRESA", "CLIENTE", "NOME"]),
      VENDA: getField(item, ["VENDA", "SERVICO", "SERVI_O", "TIPO"]),
      TIPO_REGISTRO: getField(item, ["TIPO_REGISTRO"]) || "Venda/aditivo",
      STATUS_FINANCEIRO: "Venda",
      ORIGEM: "Importado",
      _manualSale: true,
      _saleKey: item._saleKey || `sale-${Date.now()}-${Math.random().toString(16).slice(2)}`
    };
  }

  function normalizeImportedReceivable(item) {
    const receivable = {
      CONSULTOR: getField(item, ["CONSULTOR"]) || "Não informado",
      CONSULTOR_CANONICO: canonicalConsultant(getField(item, ["CONSULTOR"])),
      MES: getField(item, ["MES", "M_S", "MES_"]) || state.salesMonth || state.month || latestBaseMonth() || "JULHO",
      VALOR: getField(item, ["VALOR"]),
      EMPRESA: getField(item, ["EMPRESA", "CLIENTE", "NOME"]),
      VENDA: getField(item, ["VENDA", "SERVICO", "SERVI_O", "TIPO"]),
      TIPO_REGISTRO: getField(item, ["TIPO_REGISTRO"]) || "Mensalidade",
      STATUS_FINANCEIRO: getField(item, ["STATUS_FINANCEIRO", "FINANCEIRO", "STATUS"]) || "Pendente",
      ORIGEM: "Importado",
      _manualReceivable: true,
      _receivableKey: `receivable-${Date.now()}-${Math.random().toString(16).slice(2)}`
    };
    return receivable;
  }

  function normalizeImportedCancel(item) {
    return {
      EMPRESA: getField(item, ["EMPRESA", "CLIENTE", "NOME"]),
      MES: getField(item, ["MES", "M_S", "MES_"]) || state.month || state.salesMonth || latestBaseMonth() || "JULHO",
      STATUS: getField(item, ["STATUS"]) || "CANCELADO",
      VALOR: getField(item, ["VALOR"]),
      MOTIVO: getField(item, ["MOTIVO", "RECORRENCIA", "TIPO"]) || "Importado",
      CONSULTOR: getField(item, ["CONSULTOR"]) || "POS",
      CONSULTOR_CANONICO: canonicalConsultant(getField(item, ["CONSULTOR"]) || "POS"),
      ORIGEM: "Importado",
      _manualCancel: true,
      _cancelKey: `cancel-${Date.now()}-${Math.random().toString(16).slice(2)}`
    };
  }

  function importCurrentViewFile(file) {
    if (!canManage() || !file) return;
    const reader = new FileReader();
    reader.addEventListener("load", async () => {
      try {
        const rows = importedRowsFromFile(reader.result, file.name);
        let imported = [];
        let importResult = null;
        if (state.view === "clientes" || state.view === "financeiro") {
          imported = rows.map(normalizeImportedClient).filter((item) => getField(item, ["EMPRESA"]));
          state.manual = imported.concat(state.manual);
          saveManual();

          refreshBaseRows();
          importResult = await saveSupabaseClientsBatch(imported);
        } else if (state.view === "vendas" && state.salesMode === "pagos") {
          const sourceRows = rows.some((item) => normalizeText(item._sheet).includes("PAGOS")) ? rows.filter((item) => normalizeText(item._sheet).includes("PAGOS")) : rows;
          imported = sourceRows.map(normalizeImportedReceivable).filter((item) => getField(item, ["EMPRESA"]));
          state.manualReceivables = imported.concat(state.manualReceivables);
          imported.forEach((item) => paidMonthlyRows.unshift(item));
          saveManualReceivables();
          importResult = await saveSupabaseBatch(imported, saveSupabaseReceivable);
        } else if (state.view === "vendas") {
          const sourceRows = rows.filter((item) => !normalizeText(item._sheet).includes("PAGOS"));
          imported = sourceRows.map(normalizeImportedSale).filter((item) => getField(item, ["EMPRESA"]));
          state.manualSales = imported.concat(state.manualSales);
          imported.forEach((item) => {
            salesMonthlyRows.unshift(item);
            salesRows.unshift(item);
          });
          saveManualSales();
          importResult = await saveSupabaseBatch(imported, saveSupabaseSale);
        } else if (state.view === "cancelamentos") {
          imported = rows.map(normalizeImportedCancel).filter((item) => getField(item, ["EMPRESA"]));
          imported.forEach((item) => delete state.cancelEdits[item._cancelKey]);
          state.manualCancels = imported.concat(state.manualCancels.map(normalizeManualCancel));
          imported.forEach((item) => cancelRows.unshift(item));
          persistManualCancels();
          saveEdits();
          refreshBaseRows();
          importResult = await saveSupabaseBatch(imported, saveSupabaseCancel);
        }
        if (supabaseClient) await loadSupabaseData();
        setupMonths();
        setupSelectOptions();
        renderAll();
        alert(importResultMessage(imported.length, importResult));
      } catch (error) {
        alert(error.message || "Não foi possível importar esse arquivo. Use CSV, JSON ou Excel (.xlsx/.xls).");
      }
    });
    reader.readAsArrayBuffer(file);
  }
  function getField(item, fields) {
    for (const field of fields) {
      if (item[field] !== undefined && item[field] !== null && item[field] !== "") return item[field];
    }
    return "";
  }

  function canonicalConsultant(value) {
    const text = String(value || "").trim();
    const compact = normalizeText(text).replace(/[^A-Z0-9]/g, "");
    if (compact === "POS" || compact === "PAS") return "PÓS";
    return text;
  }

  const monthNumber = {
    JANEIRO: 1,
    FEVEREIRO: 2,
    MARÇO: 3,
    ABRIL: 4,
    MAIO: 5,
    JUNHO: 6,
    JULHO: 7,
    AGOSTO: 8,
    SETEMBRO: 9,
    OUTUBRO: 10,
    NOVEMBRO: 11,
    DEZEMBRO: 12
  };

  function appYear() {
    return new Date().getFullYear();
  }

  function toIsoDate(value) {
    const text = String(value || "").trim();
    if (!text) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
    const serial = Number(text);
    if (Number.isFinite(serial) && serial > 20000) {
      return new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString().slice(0, 10);
    }
    const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
    return null;
  }

  function competenciaFromMonth(value) {
    const month = canonicalMonth(value) || latestBaseMonth() || "JULHO";
    const number = monthNumber[month] || (new Date().getMonth() + 1);
    return `${appYear()}-${String(number).padStart(2, "0")}-01`;
  }

  function monthFromCompetencia(value) {
    const match = String(value || "").match(/^\d{4}-(\d{2})/);
    if (!match) return canonicalMonth(value);
    return monthOrder[Number(match[1]) - 1] || "";
  }

  function dbClientStatus(value) {
    const text = normalizeText(value);
    if (text.includes("BLOQUE")) return "bloqueado";
    if (text.includes("CANCEL")) return "cancelado";
    return "ativo";
  }

  function appClientStatus(value) {
    const text = normalizeText(value);
    if (text.includes("BLOQUE")) return "BLOQUEIO";
    if (text.includes("CANCEL")) return "CANCELADO";
    return "ATIVO";
  }

  function dbFinanceStatus(value) {
    const text = normalizeText(value);
    if (text.includes("PAGO")) return "pago";
    if (text.includes("ATRASO")) return "atraso";
    if (text.includes("RECEBER")) return "receber";
    return "pendente";
  }

  function appFinanceStatus(value) {
    const text = normalizeText(value);
    if (text.includes("PAGO")) return "Pago";
    if (text.includes("ATRASO")) return "Atraso";
    if (text.includes("RECEBER")) return "Receber";
    return "Pendente";
  }

  function dbCancelType(value) {
    const text = normalizeText(value);
    return text.includes("DOWNGRADE") ? "downgrade" : "cancelamento";
  }

  function clientIdentityPayload(item) {
    return {
      empresa: getField(item, ["EMPRESA"]),
      filial: getField(item, ["FILIAL"]) || null,
      sistema: getField(item, ["SISTEMA"]) || null,
      servico: getField(item, ["SERVICO", "SERVI_O", "VENDA"]) || null,
      entrada: toIsoDate(getField(item, ["ENTRADA", "VENCIMENTO"])),
      vencimento_dia: (() => {
        const date = toIsoDate(getField(item, ["VENCIMENTO"]));
        return date ? Number(date.slice(8, 10)) : null;
      })(),
      status: dbClientStatus(getField(item, ["STATUS"])),
      valor_mensal: numberValue(getField(item, ["VALOR"])),
      observacao: getField(item, ["OBSERVACAO", "OBSERVAÇÃO"]) || null
    };
  }

  async function ensureSupabaseClient(item) {
    if (!supabaseClient) return null;
    if (item._supabaseClientId) return item._supabaseClientId;
    const payload = clientIdentityPayload(item);
    if (!payload.empresa) throw new Error("Cliente sem empresa para salvar no banco.");
    let query = supabaseClient
      .from("clientes")
      .select("id")
      .eq("empresa", payload.empresa);
    ["filial", "sistema", "servico"].forEach((field) => {
      query = payload[field] == null ? query.is(field, null) : query.eq(field, payload[field]);
    });
    const found = await query.limit(1).maybeSingle();
    if (found.error) throw found.error;
    if (found.data?.id) {
      const { error } = await supabaseClient.from("clientes").update(payload).eq("id", found.data.id);
      if (error) throw error;
      return found.data.id;
    }
    const { data, error } = await supabaseClient.from("clientes").insert(payload).select("id").single();
    if (error) throw error;
    return data.id;
  }

  function dbMonthlyPayload(item, clienteId) {
    return {
      cliente_id: clienteId,
      competencia: competenciaFromMonth(getField(item, ["MES_REFERENCIA", "MES", "ENTRADA"])),
      empresa: getField(item, ["EMPRESA"]),
      filial: getField(item, ["FILIAL"]) || null,
      sistema: getField(item, ["SISTEMA"]) || null,
      servico: getField(item, ["SERVICO", "SERVI_O", "VENDA"]) || null,
      status: dbClientStatus(getField(item, ["STATUS"])),
      financeiro: dbFinanceStatus(getField(item, ["FINANCEIRO", "STATUS_FINANCEIRO"])),
      valor_mensal: numberValue(getField(item, ["VALOR"])),
      vencimento: toIsoDate(getField(item, ["VENCIMENTO"])),
      entrada: toIsoDate(getField(item, ["ENTRADA", "VENCIMENTO"])),
      observacao: getField(item, ["OBSERVACAO", "OBSERVAÇÃO"]) || null
    };
  }

  async function saveSupabaseClient(item) {
    if (!supabaseClient || !canManage()) return null;
    const clienteId = await ensureSupabaseClient(item);
    const payload = dbMonthlyPayload(item, clienteId);
    const target = supabaseClient.from("clientes_mensais");
    const request = item._supabaseMonthlyId
      ? target.update(payload).eq("id", item._supabaseMonthlyId).select("id").single()
      : target.upsert(payload, { onConflict: "cliente_id,competencia" }).select("id").single();
    const { data, error } = await request;
    if (error) throw error;
    item._supabaseClientId = clienteId;
    item._supabaseMonthlyId = data?.id || item._supabaseMonthlyId;
    item._rowKey = item._supabaseMonthlyId ? `client-${item._supabaseMonthlyId}` : item._rowKey;
    return clienteId;
  }

  async function saveSupabaseSale(item) {
    if (!supabaseClient || !canManage()) return;
    const payload = {
      cliente_id: item._supabaseClientId || null,
      competencia: competenciaFromMonth(getField(item, ["MES"])),
      data_venda: toIsoDate(getField(item, ["DATA", "DATA_VENDA"])) || null,
      tipo: normalizeText(getField(item, ["TIPO_REGISTRO", "VENDA"])).includes("ADITIVO") ? "aditivo" : "venda",
      empresa: getField(item, ["EMPRESA"]),
      sistema: getField(item, ["SISTEMA"]) || null,
      servico: getField(item, ["VENDA", "SERVICO", "SERVI_O"]) || null,
      consultor: getField(item, ["CONSULTOR", "CONSULTOR_CANONICO"]) || null,
      valor: numberValue(getField(item, ["VALOR"])),
      observacao: getField(item, ["OBSERVACAO", "OBSERVAÇÃO"]) || null
    };
    const table = supabaseClient.from("vendas");
    const { data, error } = item._supabaseId
      ? await table.update(payload).eq("id", item._supabaseId).select("id").single()
      : await table.insert(payload).select("id").single();
    if (error) throw error;
    item._supabaseId = data.id;
  }

  async function saveSupabaseReceivable(item) {
    if (!supabaseClient || !canManage()) return;
    const clienteId = await ensureSupabaseClient({
      EMPRESA: getField(item, ["EMPRESA"]),
      SISTEMA: getField(item, ["SISTEMA"]),
      SERVICO: getField(item, ["VENDA", "SERVICO", "SERVI_O"]),
      VALOR: getField(item, ["VALOR"]),
      STATUS: "ATIVO",
      FINANCEIRO: getField(item, ["STATUS_FINANCEIRO"])
    });
    const payload = {
      cliente_id: clienteId,
      competencia: competenciaFromMonth(getField(item, ["MES"])),
      empresa: getField(item, ["EMPRESA"]),
      sistema: getField(item, ["SISTEMA"]) || null,
      servico: getField(item, ["VENDA", "SERVICO", "SERVI_O"]) || null,
      valor_original: numberValue(getField(item, ["VALOR"])),
      valor_downgrade: 0,
      status: dbFinanceStatus(getField(item, ["STATUS_FINANCEIRO"])),
      vencimento: toIsoDate(getField(item, ["VENCIMENTO"])),
      observacao: getField(item, ["OBSERVACAO", "OBSERVAÇÃO"]) || null
    };
    const table = supabaseClient.from("pagos_mensais");
    const { data, error } = item._supabaseId
      ? await table.update(payload).eq("id", item._supabaseId).select("id").single()
      : await table.upsert(payload, { onConflict: "cliente_id,competencia" }).select("id").single();
    if (error) throw error;
    item._supabaseId = data.id;
    item._supabaseClientId = clienteId;
  }

  async function saveSupabaseCancel(item) {
    if (!supabaseClient || !canManage()) return;
    const payload = {
      cliente_id: item._supabaseClientId || null,
      competencia: competenciaFromMonth(getField(item, ["MES", "MES_REFERENCIA"])),
      data_evento: toIsoDate(getField(item, ["DATA", "DATA_EVENTO"])) || null,
      tipo: dbCancelType(getField(item, ["STATUS", "MOTIVO"])),
      empresa: getField(item, ["EMPRESA"]),
      sistema: getField(item, ["SISTEMA"]) || null,
      servico: getField(item, ["SERVICO", "SERVI_O", "VENDA"]) || null,
      consultor: getField(item, ["CONSULTOR", "CONSULTOR_CANONICO"]) || null,
      motivo: getField(item, ["MOTIVO", "RECORRENCIA"]) || null,
      valor: numberValue(getField(item, ["VALOR"])),
      observacao: getField(item, ["OBSERVACAO", "OBSERVAÇÃO"]) || null
    };
    const table = supabaseClient.from("cancelamentos");
    const { data, error } = item._supabaseId
      ? await table.update(payload).eq("id", item._supabaseId).select("id").single()
      : await table.insert(payload).select("id").single();
    if (error) throw error;
    item._supabaseId = data.id;
  }
  function chunkRows(rows, size = 100) {
    const chunks = [];
    for (let index = 0; index < rows.length; index += size) chunks.push(rows.slice(index, index + size));
    return chunks;
  }

  function identityKeyFromPayload(item) {
    return JSON.stringify([item.empresa ?? null, item.filial ?? null, item.sistema ?? null, item.servico ?? null]);
  }

  async function saveSupabaseClientsBatch(rows) {
    if (!supabaseClient || !canManage()) return { saved: rows.length, failed: [] };
    const failed = [];
    let saved = 0;

    for (const chunk of chunkRows(rows, 100)) {
      try {
        const clientPayloadMap = new Map();
        chunk.map(clientIdentityPayload).filter((item) => item.empresa).forEach((item) => {
          clientPayloadMap.set(identityKeyFromPayload(item), item);
        });
        const clientPayloads = [...clientPayloadMap.values()];
        const { data: clients, error: clientError } = await supabaseClient
          .from("clientes")
          .upsert(clientPayloads, { onConflict: "empresa,filial,sistema,servico" })
          .select("id, empresa, filial, sistema, servico");
        if (clientError) throw clientError;

        const clientMap = new Map((clients || []).map((item) => [identityKeyFromPayload(item), item.id]));
        const monthlyPayloads = chunk.map((item, index) => {
          const identity = clientIdentityPayload(item);
          const clientId = clientMap.get(identityKeyFromPayload(identity));
          if (!clientId) {
            failed.push({ index: saved + index + 1, message: "Cliente não retornou ID do banco." });
            return null;
          }
          item._supabaseClientId = clientId;
          return dbMonthlyPayload(item, clientId);
        }).filter(Boolean);

        const uniqueMonthlyPayloads = [...new Map(
          monthlyPayloads.map((item) => [JSON.stringify([item.cliente_id, item.competencia]), item])
        ).values()];
        if (uniqueMonthlyPayloads.length) {
          const { error: monthlyError } = await supabaseClient
            .from("clientes_mensais")
            .upsert(uniqueMonthlyPayloads, { onConflict: "cliente_id,competencia" });
          if (monthlyError) throw monthlyError;
        }
        saved += uniqueMonthlyPayloads.length;
      } catch (error) {
        chunk.forEach((_, index) => failed.push({ index: saved + index + 1, message: error.message || String(error) }));
      }
    }

    return { saved, failed };
  }

  async function saveSupabaseBatch(rows, saveFn) {
    if (!supabaseClient || !canManage()) return { saved: rows.length, failed: [] };
    const failed = [];
    let saved = 0;
    for (let index = 0; index < rows.length; index++) {
      try {
        await saveFn(rows[index]);
        saved++;
      } catch (error) {
        failed.push({ index: index + 1, message: error.message || String(error) });
      }
    }
    return { saved, failed };
  }

  function importResultMessage(total, result) {
    if (!result || !result.failed?.length) return `${total} registros importados.`;
    const sample = result.failed.slice(0, 3).map((item) => `linha ${item.index}: ${item.message}`).join("\n");
    return `${result.saved} de ${total} registros importados.\n${result.failed.length} falharam.\n${sample}`;
  }

  function getItemMonth(item) {
    const month = getField(item, ["MES_REFERENCIA", "MES", "M_S", "MES_", "ENTRADA"]);
    if (month) return canonicalMonth(month);
    const dateValue = getField(item, ["VENCIMENTO", "BASE"]);
    const serial = Number(dateValue);
    if (Number.isFinite(serial) && serial > 20000) {
      const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
      return canonicalMonth(monthOrder[date.getUTCMonth()] || "");
    }
    const dateText = String(dateValue || "");
    const dateMatch = dateText.match(/(?:\d{1,2}\/)?(\d{1,2})\/\d{4}/);
    if (dateMatch) return canonicalMonth(monthOrder[Number(dateMatch[1]) - 1] || "");
    return "";
  }

  function matchesFilters(item) {
    const haystack = normalizeText(Object.values(item).join(" "));
    const statusText = normalizeText(getField(item, ["STATUS", "FINANCEIRO", "MOTIVO"]));
    const monthText = getItemMonth(item);
    const searchOk = !state.search || haystack.includes(normalizeText(state.search));
    const statusOk = !state.status || statusText.includes(normalizeText(state.status)) || haystack.includes(normalizeText(state.status));
    const monthOk = !state.month || monthText === normalizeText(state.month);
    const systemOk = matchesOptionalField(item, ["SISTEMA"], state.system);
    const serviceOk = matchesOptionalField(item, ["SERVICO", "SERVI_O"], state.service);
    const financeOk = matchesOptionalField(item, ["FINANCEIRO", "FINANCEIRO_2"], state.finance);
    const consultantOk = matchesOptionalField(item, ["CONSULTOR_CANONICO", "CONSULTOR"], state.consultant);
    const reasonOk = matchesOptionalField(item, ["MOTIVO", "RECORRENCIA"], state.reason);
    return searchOk && statusOk && monthOk && systemOk && serviceOk && financeOk && consultantOk && reasonOk;
  }

  function matchesOptionalField(item, fields, selectedValue) {
    if (!selectedValue) return true;
    const value = getField(item, fields);
    if (!value) return true;
    return normalizeText(value).includes(normalizeText(selectedValue));
  }

  function uniqueCount(rows, field) {
    return new Set(rows.map((item) => normalizeText(getField(item, [field]))).filter(Boolean)).size;
  }

  function sumBy(rows, field) {
    return rows.reduce((total, item) => total + numberValue(getField(item, [field])), 0);
  }

  function groupSum(rows, labelField, valueField) {
    const map = new Map();
    rows.forEach((item) => {
      const label = String(getField(item, [labelField]) || "Não informado").trim();
      map.set(label, (map.get(label) || 0) + numberValue(getField(item, [valueField])));
    });
    return [...map.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
  }

  function groupCount(rows, labelField) {
    const map = new Map();
    rows.forEach((item) => {
      const label = String(getField(item, [labelField]) || "Não informado").trim();
      map.set(label, (map.get(label) || 0) + 1);
    });
    return [...map.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
  }

  function renderBars(targetId, data, formatter = (value) => value, titleFormatter = formatter) {
    const target = document.getElementById(targetId);
    const top = data.filter((item) => item.label).slice(0, state.ranking);
    const max = Math.max(...top.map((item) => item.value), 1);
    target.innerHTML = top.length ? top.map((item) => `
      <div class="bar-row">
        <span class="bar-label" title="${escapeHtml(item.label)}">${escapeHtml(item.label)}</span>
        <span class="bar-track"><span class="bar-fill" style="width:${Math.max(4, (item.value / max) * 100)}%"></span></span>
        <span class="bar-value" title="${escapeHtml(titleFormatter(item.value))}">${formatter(item.value)}</span>
      </div>
    `).join("") : `<div class="empty">Sem dados para exibir.</div>`;
  }

  function statusPill(value) {
    const text = String(value || "Não informado");
    const normalized = normalizeText(text);
    const variant = normalized.includes("CANCEL") || normalized.includes("BLOQUEIO")
      ? "danger"
      : normalized.includes("ATRASO") || normalized.includes("DOWNGRADE")
        ? "warn"
        : "";
    return `<span class="pill ${variant}">${escapeHtml(text)}</span>`;
  }

  function renderTable(targetId, rows, columns, limit = 180, options = {}) {
    const target = document.getElementById(targetId);
    const filtered = rows.filter(matchesFilters);
    const visible = filtered.slice(0, limit);
    const totalHtml = options.totalField ? renderTotalTableRow(options.totalRows || filtered, columns, options) : "";
    target.innerHTML = visible.length ? visible.map((row) => `
      <tr class="${options.rowClass ? options.rowClass(row) : ""}">${columns.map((column) => {
        const value = column.render ? column.render(row) : escapeHtml(getField(row, column.fields));
        const rawValue = column.title ? column.title(row) : getField(row, column.fields);
        return `<td class="${column.className || ""}" title="${escapeHtml(rawValue)}">${value}</td>`;
      }).join("")}</tr>
    `).join("") + totalHtml : `<tr><td colspan="${columns.length}" class="empty">Nenhum registro encontrado.</td></tr>`;
  }

  function renderTotalTableRow(rows, columns, options) {
    const total = sumBy(rows, options.totalField);
    const totalColumnIndex = options.totalColumnIndex ?? columns.length - 1;
    return `<tr class="total-row">${columns.map((column, index) => {
      if (index === 0) return `<td>Total filtrado (${rows.length})</td>`;
      if (index === totalColumnIndex) return `<td class="num-cell">${formatMoney(total)}</td>`;
      return "<td></td>";
    }).join("")}</tr>`;
  }

  function cancelActions(row) {
    if (!canManage()) return "";
    return `
      <div class="row-actions">
        <button class="ghost-btn table-action" type="button" data-cancel-edit="${escapeHtml(row._cancelKey)}">Editar</button>
        <button class="ghost-btn table-action danger" type="button" data-cancel-delete="${escapeHtml(row._cancelKey)}">Excluir</button>
      </div>
    `;
  }

  function openCancelModal(title = "Novo cancelamento") {
    if (!canManage()) return;
    const modal = document.getElementById("cancelModal");
    const titleEl = document.getElementById("cancelModalTitle");
    if (titleEl) titleEl.textContent = title;
    if (!modal) return;
    modal.hidden = false;
    document.body.classList.add("modal-open");
    setTimeout(() => modal.querySelector("input, select, button")?.focus(), 0);
  }

  function closeCancelModal(reset = true) {
    const modal = document.getElementById("cancelModal");
    const form = document.getElementById("cancelForm");
    if (reset && form) {
      form.reset();
      form.querySelector("button[type='submit']").textContent = "Salvar cancelamento";
      state.editingCancelKey = "";
    }
    if (modal) modal.hidden = true;
    document.body.classList.remove("modal-open");
  }
  function fillCancelForm(row) {
    const form = document.getElementById("cancelForm");
    if (!form) return;
    form.elements.empresa.value = getField(row, ["EMPRESA"]);
    form.elements.mes.value = getField(row, ["MES", "MES_REFERENCIA"]) || getItemMonth(row);
    form.elements.status.value = getField(row, ["STATUS"]) || "CANCELADO";
    form.elements.motivo.value = getField(row, ["MOTIVO", "RECORRENCIA"]);
    form.elements.consultor.value = getField(row, ["CONSULTOR_CANONICO", "CONSULTOR"]);
    form.elements.valor.value = numberValue(getField(row, ["VALOR"])) || "";
    form.querySelector("button[type='submit']").textContent = "Salvar alteração";
    openCancelModal('Editar cancelamento');
  }

  function cancelPayloadFromForm(form) {
    const consultant = form.get("consultor") || "POS";
    return normalizeManualCancel({
      EMPRESA: form.get("empresa"),
      MES: form.get("mes"),
      STATUS: form.get("status") || "CANCELADO",
      MOTIVO: form.get("motivo") || "Manual",
      CONSULTOR: consultant,
      CONSULTOR_CANONICO: canonicalConsultant(consultant),
      VALOR: form.get("valor"),
      ORIGEM: "Manual",
      _manualCancel: true,
      _cancelKey: state.editingCancelKey || `cancel-${Date.now()}-${Math.random().toString(16).slice(2)}`
    });
  }

  function editCancel(key) {
    if (!canManage()) return;
    const row = allCancelRows().find((item) => (item._cancelKey || cancelKey(item)) === key);
    if (!row) return;
    state.editingCancelKey = key;
    fillCancelForm(row);
  }

  async function deleteCancel(key) {
    if (!canManage()) return;
    if (!confirm("Excluir este cancelamento?")) return;
    const row = allCancelRows().find((item) => (item._cancelKey || cancelKey(item)) === key);
    if (supabaseClient && row?._supabaseId) {
      const { error } = await supabaseClient.from("cancelamentos").delete().eq("id", row._supabaseId);
      if (error) return alert(error.message || "Não foi possível excluir no banco.");
    }
    state.cancelEdits[key] = { deleted: true };
    state.manualCancels = state.manualCancels.map(normalizeManualCancel).filter((item) => item._cancelKey !== key);
    saveManualCancels();
    saveEdits();
    renderAll();
  }

  function clearManualCancels() {
    if (!confirm("Limpar todos os cancelamentos importados/manuais e ocultar os da base?")) return;
    state.manualCancels = [];
    allCancelRows().forEach((item) => {
      const key = item._cancelKey || cancelKey(item);
      state.cancelEdits[key] = { deleted: true };
    });
    state.editingCancelKey = "";
    saveManualCancels();
    saveEdits();
    renderAll();
  }

  function clearAllData() {
    if (!confirm("Limpar todo o cache do sistema? Isso removerá clientes, vendas, recebíveis, cancelamentos, sobrescritas e sua sessão. Continuar?")) return;
    localStorage.removeItem(manualKey);
    localStorage.removeItem(manualSalesKey);
    localStorage.removeItem(manualReceivablesKey);
    localStorage.removeItem(manualCancelsKey);
    localStorage.removeItem(receivableStatusKey);
    localStorage.removeItem(salesEditsKey);
    localStorage.removeItem(receivableEditsKey);
    localStorage.removeItem(cancelEditsKey);
    localStorage.removeItem(overrideKey);
    localStorage.removeItem(authKey);
    sessionStorage.removeItem(authKey);
    state.manual = [];
    state.manualSales = [];
    state.manualReceivables = [];
    state.manualCancels = [];
    state.overrides = {};
    state.receivableStatus = {};
    state.salesEdits = {};
    state.receivableEdits = {};
    state.cancelEdits = {};
    state.session = null;
    saveManual();
    saveManualSales();
    saveManualReceivables();
    saveManualCancels();
    saveReceivableStatus();
    saveOverrides();
    saveSession();
    refreshBaseRows();
    renderAll();
  }
  function clientActions(row) {
    if (!canManage()) return "";
    const key = row._rowKey || clientKey(row);
    return `
      <div class="row-actions">
        <button class="ghost-btn table-action" type="button" data-client-edit="${escapeHtml(key)}">Editar</button>
        <button class="ghost-btn table-action danger" type="button" data-client-delete="${escapeHtml(key)}">Excluir</button>
      </div>
    `;
  }

  function saleActions(row) {
    if (!canManage()) return "";
    const key = row._saleKey || saleKey(row);
    return `
      <div class="row-actions">
        <button class="ghost-btn table-action" type="button" data-sale-edit="${escapeHtml(key)}">Editar</button>
        <button class="ghost-btn table-action danger" type="button" data-sale-delete="${escapeHtml(key)}">Excluir</button>
      </div>
    `;
  }

  function paidActions(row) {
    if (!canManage()) return "";
    const key = row._receivableKey || receivableKey(row);
    const currentStatus = normalizeText(getField(row, ["STATUS_FINANCEIRO"]));
    const statuses = ["Pago", "Pendente", "Atraso"];
    return `
      <div class="row-actions">
        <select class="table-status-select" data-receivable-status="${escapeHtml(key)}" aria-label="Status financeiro">
          ${statuses.map((status) => `<option value="${escapeHtml(status)}" ${normalizeText(status) === currentStatus ? "selected" : ""}>${escapeHtml(status)}</option>`).join("")}
        </select>
        <button class="ghost-btn table-action" type="button" data-receivable-edit="${escapeHtml(key)}">Editar</button>
        <button class="ghost-btn table-action danger" type="button" data-receivable-delete="${escapeHtml(key)}">Excluir</button>
      </div>
    `;
  }

  function salePayloadFromForm(form, type = "sale") {
    const payload = {
      CONSULTOR: form.get("consultor"),
      CONSULTOR_CANONICO: canonicalConsultant(form.get("consultor")),
      MES: form.get("mes"),
      VALOR: form.get("valor"),
      EMPRESA: form.get("empresa"),
      VENDA: form.get("venda")
    };
    if (type === "receivable") {
      return {
        ...payload,
        TIPO_REGISTRO: "Mensalidade",
        STATUS_FINANCEIRO: "Pendente",
        _manualReceivable: true,
        _receivableKey: state.editingReceivableKey || `receivable-${Date.now()}-${Math.random().toString(16).slice(2)}`
      };
    }
    return {
      ...payload,
      TIPO_REGISTRO: "Venda/aditivo",
      STATUS_FINANCEIRO: "Venda",
      _manualSale: true,
      _saleKey: state.editingSaleKey || `sale-${Date.now()}-${Math.random().toString(16).slice(2)}`
    };
  }




  function activeFiltersCount() {
    return [state.search, state.status, state.month, state.system, state.service, state.finance, state.consultant, state.reason]
      .filter((value) => String(value || "").trim()).length;
  }

  function updateFiltersButton() {
    const button = document.getElementById("filtersToggleBtn");
    const counter = document.getElementById("filtersActiveCount");
    const panel = document.getElementById("advancedFilters");
    if (!button) return;
    const count = activeFiltersCount();
    const isOpen = panel ? !panel.hidden : false;
    const icon = button.querySelector(".filter-toggle-icon");
    if (icon) icon.textContent = isOpen ? "^" : "v";
    if (counter) {
      counter.textContent = String(count);
      counter.hidden = count === 0;
    }
    button.classList.toggle("has-active-filters", count > 0);
  }

  function setAdvancedFiltersOpen(open) {
    const panel = document.getElementById("advancedFilters");
    const button = document.getElementById("filtersToggleBtn");
    if (!panel || !button) return;
    panel.hidden = !open;
    button.setAttribute("aria-expanded", String(open));
    button.classList.toggle("active", open);
    updateFiltersButton();
  }

  function ensureSalesModal() {
    const form = document.getElementById("saleForm");
    if (!form) return;
    let modal = document.getElementById("saleModal");
    if (!modal) {
      modal = document.createElement("div");
      modal.className = "modal-backdrop";
      modal.id = "saleModal";
      modal.hidden = true;
      modal.innerHTML = `
        <section class="modal-card" role="dialog" aria-modal="true" aria-labelledby="saleModalTitle">
          <div class="panel-head modal-head">
            <div>
              <p class="eyebrow">Comercial</p>
              <h2 id="saleModalTitle">Nova venda ou aditivo</h2>
            </div>
            <button class="icon-btn" id="saleModalXBtn" type="button" aria-label="Fechar modal">×</button>
          </div>
        </section>
      `;
      document.body.appendChild(modal);
    }
    const card = modal.querySelector(".modal-card");
    if (card && form.closest("#saleModal") !== modal) {
      form.classList.add("compact-form");
      let actions = form.querySelector(".modal-actions");
      if (!actions) {
        const submit = form.querySelector("button[type='submit']");
        actions = document.createElement("div");
        actions.className = "modal-actions full";
        actions.innerHTML = '<button class="ghost-btn" id="saleModalCloseBtn" type="button">Cancelar</button>';
        if (submit) actions.appendChild(submit);
        form.appendChild(actions);
      }
      card.appendChild(form);
    }
    document.querySelectorAll("#vendas .form-panel").forEach((panel) => {
      if (!panel.querySelector("form")) panel.remove();
    });
    const salesHead = document.querySelector("#vendas .sales-workspace .panel-head");
    if (salesHead && !document.getElementById("newSaleBtn")) {
      const segmented = salesHead.querySelector(".segmented");
      const actions = document.createElement("div");
      actions.className = "section-actions";
      actions.innerHTML = '<button class="primary-btn admin-only" id="newSaleBtn" type="button">Nova venda/aditivo</button>';
      if (segmented) actions.appendChild(segmented);
      salesHead.appendChild(actions);
      applyAccessState();
    }
  }  
  function openSaleModal(title = "Nova venda ou aditivo") {
    if (!canManage()) return;
    const modal = document.getElementById("saleModal");
    const titleEl = document.getElementById("saleModalTitle");
    if (titleEl) titleEl.textContent = title;
    if (!modal) return;
    modal.hidden = false;
    document.body.classList.add("modal-open");
    setTimeout(() => modal.querySelector("input, select, button")?.focus(), 0);
  }

  function closeSaleModal(reset = true) {
    const modal = document.getElementById("saleModal");
    const form = document.getElementById("saleForm");
    if (reset && form) {
      form.reset();
      form.querySelector("button[type='submit']").textContent = "Salvar venda";
      state.editingSaleKey = "";
      state.editingReceivableKey = "";
    }
    if (modal) modal.hidden = true;
    document.body.classList.remove("modal-open");
  }
  function fillSaleForm(row) {
    const form = document.getElementById("saleForm");
    if (!form) return;
    form.elements.empresa.value = getField(row, ["EMPRESA"]);
    form.elements.consultor.value = getField(row, ["CONSULTOR_CANONICO", "CONSULTOR"]) || "PÓS";
    form.elements.mes.value = monthLabel(getItemMonth(row));
    form.elements.valor.value = numberValue(getField(row, ["VALOR"])) || "";
    form.elements.venda.value = getField(row, ["VENDA", "SERVICO", "SERVI_O"]);
    form.querySelector("button[type='submit']").textContent = "Salvar alteração";
  }

  function editSale(key) {
    const row = getSalesWorkspaceRows().find((item) => (item._saleKey || saleKey(item)) === key);
    if (!row) return;
    state.editingSaleKey = key;
    state.editingReceivableKey = "";
    state.salesMode = "vendas";
    setView("vendas");
    fillSaleForm(row);
    openSaleModal('Editar venda/aditivo');
  }

  function editReceivable(key) {
    const row = getSalesWorkspaceRows().find((item) => (item._receivableKey || receivableKey(item)) === key);
    if (!row) return;
    state.editingReceivableKey = key;
    state.editingSaleKey = "";
    state.salesMode = "pagos";
    setView("vendas");
    fillSaleForm(row);
    openSaleModal('Editar pago mensal');
  }

  async function deleteSale(key) {
    if (!confirm("Excluir esta venda/aditivo?")) return;
    const row = getSalesWorkspaceRows().find((item) => (item._saleKey || saleKey(item)) === key);
    if (supabaseClient && row?._supabaseId) {
      const { error } = await supabaseClient.from("vendas").delete().eq("id", row._supabaseId);
      if (error) return alert(error.message || "Não foi possível excluir no banco.");
    }
    state.salesEdits[key] = { deleted: true };
    state.manualSales = state.manualSales.filter((item) => (item._saleKey || saleKey(item)) !== key);
    saveManualSales();
    saveEdits();
    renderAll();
  }

  async function deleteReceivable(key) {
    if (!confirm("Excluir este pagamento mensal?")) return;
    const row = getSalesWorkspaceRows().find((item) => (item._receivableKey || receivableKey(item)) === key);
    if (supabaseClient && row?._supabaseId) {
      const { error } = await supabaseClient.from("pagos_mensais").delete().eq("id", row._supabaseId);
      if (error) return alert(error.message || "Não foi possível excluir no banco.");
    }
    state.receivableEdits[key] = { deleted: true };
    state.manualReceivables = state.manualReceivables.filter((item) => (item._receivableKey || receivableKey(item)) !== key);
    saveManualReceivables();
    saveEdits();
    renderAll();
  }
  function receivableActions(row) {
    if (!canManage()) return "";
    const key = receivableKey(row);
    const currentStatus = normalizeText(getField(row, ["STATUS_FINANCEIRO"]));
    const statuses = ["Pago", "Pendente", "Atraso"];
    return `
      <div class="row-actions">
        <select class="table-status-select" data-receivable-status="${escapeHtml(key)}" aria-label="Status financeiro">
          ${statuses.map((status) => `<option value="${escapeHtml(status)}" ${normalizeText(status) === currentStatus ? "selected" : ""}>${escapeHtml(status)}</option>`).join("")}
        </select>
        <button class="ghost-btn table-action" type="button" data-receivable-edit="${escapeHtml(key)}">Editar</button>
        <button class="ghost-btn table-action danger" type="button" data-receivable-delete="${escapeHtml(key)}">Excluir</button>
      </div>
    `;
  }

  async function setReceivableStatus(key, status) {
    if (!canManage()) return;
    state.receivableStatus[key] = status;
    state.manualReceivables = state.manualReceivables.map((item) => (
      item._receivableKey === key ? { ...item, STATUS_FINANCEIRO: status } : item
    ));
    const row = paidMonthlyRows.find((item) => item._receivableKey === key);
    if (row) {
      row.STATUS_FINANCEIRO = status;
      if (supabaseClient && row._supabaseId) {
        const { error } = await supabaseClient.from("pagos_mensais").update({ status: dbFinanceStatus(status) }).eq("id", row._supabaseId);
        if (error) return alert(error.message || "Não foi possível atualizar no banco.");
      }
    }
    saveReceivableStatus();
    saveManualReceivables();
    renderAll();
  }

  function toInputDate(value) {
    const text = String(value || "");
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
    const serial = Number(value);
    if (Number.isFinite(serial) && serial > 20000) {
      const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
      return date.toISOString().slice(0, 10);
    }
    const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!match) return "";
    return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
  }

  function clientPayloadFromForm(form) {
    return {
      EMPRESA: form.get("empresa"),
      FILIAL: form.get("filial"),
      SISTEMA: form.get("sistema"),
      SERVICO: form.get("servico"),
      VALOR: form.get("valor"),
      VENCIMENTO: form.get("vencimento") || new Date().toLocaleDateString("pt-BR"),
      STATUS: form.get("status"),
      FINANCEIRO: form.get("financeiro"),
      MES_REFERENCIA: state.month || latestBaseMonth() || "JULHO"
    };
  }


  function openClientModal(title = "Novo cliente") {
    if (!canManage()) return;
    const modal = document.getElementById("clientModal");
    const titleEl = document.getElementById("clientModalTitle");
    if (titleEl) titleEl.textContent = title;
    if (!modal) return;
    modal.hidden = false;
    document.body.classList.add("modal-open");
    setTimeout(() => modal.querySelector("input, select, button")?.focus(), 0);
  }

  function closeClientModal(reset = true) {
    const modal = document.getElementById("clientModal");
    const form = document.getElementById("clientForm");
    if (reset && form) {
      form.reset();
      form.querySelector("button[type='submit']").textContent = "Salvar cliente";
      state.editingKey = "";
    }
    if (modal) modal.hidden = true;
    document.body.classList.remove("modal-open");
  }
  function fillClientForm(row) {
    const form = document.getElementById("clientForm");
    form.elements.empresa.value = getField(row, ["EMPRESA"]);
    form.elements.filial.value = getField(row, ["FILIAL"]);
    form.elements.sistema.value = getField(row, ["SISTEMA"]);
    form.elements.servico.value = getField(row, ["SERVICO", "SERVI_O"]);
    form.elements.valor.value = numberValue(getField(row, ["VALOR"])) || "";
    form.elements.vencimento.value = toInputDate(getField(row, ["VENCIMENTO"]));
    form.elements.status.value = getField(row, ["STATUS"]) || "ATIVO";
    form.elements.financeiro.value = getField(row, ["FINANCEIRO", "FINANCEIRO_2"]) || "RECEBER";
  }

  function editClient(key) {
    if (!canManage()) return;
    const row = baseRows.find((item) => (item._rowKey || clientKey(item)) === key);
    if (!row) return;
    state.editingKey = key;
    fillClientForm(row);
    setView("clientes");
    document.querySelector("#clientForm button[type='submit']").textContent = "Salvar alterações";
    openClientModal('Editar cliente');
  }

  async function deleteClient(key) {
    if (!canManage()) return;
    const row = baseRows.find((item) => (item._rowKey || clientKey(item)) === key);
    if (!row) return;
    if (!confirm(`Excluir ${getField(row, ["EMPRESA"]) || "este cliente"}?`)) return;
    if (supabaseClient && row._supabaseClientId) {
      const { error } = await supabaseClient.from("clientes").delete().eq("id", row._supabaseClientId);
      if (error) return alert(error.message || "Não foi possível excluir no banco.");
    }
    if (row._manual) {
      state.manual = state.manual.filter((item) => (item._rowKey || clientKey(item)) !== key);
      saveManual();
    } else {
      state.overrides[key] = { ...(state.overrides[key] || {}), deleted: true };
      saveOverrides();
    }
    refreshBaseRows();
    renderAll();
  }

  function renderTopClients(rows) {
    const target = document.getElementById("topClientsList");
    const top = rows
      .map((item) => ({
        empresa: getField(item, ["EMPRESA"]) || "Não informado",
        sistema: getField(item, ["SISTEMA"]) || "Não informado",
        servico: getField(item, ["SERVICO", "SERVI_O"]) || "Não informado",
        valor: numberValue(getField(item, ["VALOR"]))
      }))
      .filter((item) => item.valor > 0)
      .sort((a, b) => b.valor - a.valor)
      .slice(0, state.ranking);

    target.innerHTML = top.length ? top.map((item, index) => `
      <article class="ranking-item">
        <span class="rank-position">${String(index + 1).padStart(2, "0")}</span>
        <div>
          <strong title="${escapeHtml(item.empresa)}">${escapeHtml(item.empresa)}</strong>
          <small>${escapeHtml(item.sistema)} - ${escapeHtml(item.servico)}</small>
        </div>
        <b>${formatMoney(item.valor)}</b>
      </article>
    `).join("") : `<div class="empty">Sem clientes para esse filtro.</div>`;
  }

  function renderSalesCancelCompare(salesRowsFiltered, cancelRowsFiltered) {
    const target = document.getElementById("salesCancelBars");
    const months = [...new Set(salesRows.concat(allCancelRows()).map((item) => getItemMonth(item)).filter(Boolean))]
      .sort((a, b) => monthIndex(a) - monthIndex(b));
    const salesMap = new Map(groupSum(salesRowsFiltered, "MES", "VALOR").map((item) => [normalizeText(item.label), item.value]));
    const cancelMap = new Map(groupSum(cancelRowsFiltered, "MES", "VALOR").map((item) => [normalizeText(item.label), item.value]));
    const data = months
      .filter((month) => !state.month || normalizeText(month) === normalizeText(state.month))
      .map((month) => ({
        month,
        sales: salesMap.get(normalizeText(month)) || 0,
        cancel: cancelMap.get(normalizeText(month)) || 0
      }));
    const max = Math.max(...data.flatMap((item) => [item.sales, item.cancel]), 1);
    target.innerHTML = data.length ? data.map((item) => `
      <div class="compare-row">
        <span class="compare-month">${escapeHtml(item.month)}</span>
        <div class="compare-stack">
          <span class="compare-line sales" style="width:${Math.max(3, (item.sales / max) * 100)}%"></span>
          <span class="compare-line cancel" style="width:${Math.max(3, (item.cancel / max) * 100)}%"></span>
        </div>
        <span class="compare-values">
          <b>${formatCompactMoney(item.sales)}</b>
          <small>${formatCompactMoney(item.cancel)}</small>
        </span>
      </div>
    `).join("") : `<div class="empty">Sem dados para comparar.</div>`;
  }

  function getPaidMonthlyWorkspaceRows() {
    const monthlyRows = baseMonthlyReceivableRows
      .concat(buildProjectedMonthlyReceivableRows())
      .concat(individualSalesRows(paidMonthlyRows))
      .filter((item) => !isCanceledClient(item) && !hasCancelRecord(item));
    return normalizeFinanceRows(applyEdits(applySavedReceivableStatus(applyMonthlyAdjustments(monthlyRows)), receivableKey, state.receivableEdits, "_receivableKey"))
      .filter((item) => numberValue(getField(item, ["VALOR"])) > 0);
  }

  function getSalesWorkspaceRows() {
    if (state.salesMode !== "pagos") return individualSalesRows(applyEdits(salesMonthlyRows, saleKey, state.salesEdits, "_saleKey"));
    return getPaidMonthlyWorkspaceRows();
  }

  function monthLabel(month) {
    return monthOrder.find((item) => normalizeText(item) === normalizeText(month)) || month;
  }

  function ensureSalesMonth(rows) {
    if (state.salesMonth && rows.some((item) => getItemMonth(item) === state.salesMonth)) return;
    const firstMonth = monthOrder.find((month) => rows.some((item) => getItemMonth(item) === normalizeText(month)));
    state.salesMonth = normalizeText(firstMonth || monthOrder[0]);
  }

  function renderSalesMonthCards(rows) {
    const target = document.getElementById("salesMonthCards");
    const rowsForTotals = rows;
    const totals = new Map();
    const paidTotals = new Map();
    const lateTotals = new Map();
    const pendingTotals = new Map();
    rowsForTotals.forEach((item) => {
      const month = getItemMonth(item);
      if (!month) return;
      const value = numberValue(getField(item, ["VALOR"]));
      const status = normalizeText(getField(item, ["STATUS_FINANCEIRO"]));
      totals.set(month, (totals.get(month) || 0) + value);
      if (status.includes("ATRASO")) {
        lateTotals.set(month, (lateTotals.get(month) || 0) + value);
      } else if (status.includes("PENDENTE") || status.includes("RECEBER")) {
        pendingTotals.set(month, (pendingTotals.get(month) || 0) + value);
      } else if (state.salesMode === "pagos") {
        paidTotals.set(month, (paidTotals.get(month) || 0) + value);
      }
    });

    target.innerHTML = monthOrder.map((month) => {
      const key = normalizeText(month);
      const total = totals.get(key) || 0;
      const active = state.salesMonth === key ? "active" : "";
      const detail = state.salesMode === "pagos"
        ? `<small>Pago ${formatCompactMoney(paidTotals.get(key) || 0)} / Pendente ${formatCompactMoney(pendingTotals.get(key) || 0)} / Atraso ${formatCompactMoney(lateTotals.get(key) || 0)}</small>`
        : `<small>${rows.filter((item) => getItemMonth(item) === key).length} registros</small>`;
      return `
        <button class="month-card ${active}" type="button" data-sales-month="${key}">
          <span>${escapeHtml(month)}</span>
          <strong>${formatCompactMoney(total)}</strong>
          ${detail}
        </button>
      `;
    }).join("");
  }


  function ensureCancelMonth(rows) {
    if (state.cancelMonth && rows.some((item) => getItemMonth(item) === state.cancelMonth)) return;
    const firstMonth = monthOrder.find((month) => rows.some((item) => getItemMonth(item) === normalizeText(month)));
    state.cancelMonth = normalizeText(firstMonth || monthOrder[0]);
  }

  function renderCancelMonthCards(rows) {
    const target = document.getElementById("cancelMonthCards");
    if (!target) return;
    const totals = new Map();
    const counts = new Map();
    rows.forEach((item) => {
      const month = getItemMonth(item);
      if (!month) return;
      totals.set(month, (totals.get(month) || 0) + numberValue(getField(item, ["VALOR"])));
      counts.set(month, (counts.get(month) || 0) + 1);
    });
    target.innerHTML = monthOrder.map((month) => {
      const key = normalizeText(month);
      const active = state.cancelMonth === key ? "active" : "";
      return `
        <button class="month-card ${active}" type="button" data-cancel-month="${key}">
          <span>${escapeHtml(month)}</span>
          <strong>${formatCompactMoney(totals.get(key) || 0)}</strong>
          <small>${counts.get(key) || 0} registros</small>
        </button>
      `;
    }).join("");
  }

  function renderCancelWorkspace(cancelColumns) {
    const rows = allCancelRows().filter(matchesFilters);
    ensureCancelMonth(rows);
    renderCancelMonthCards(rows);
    const selectedRows = rows.filter((item) => getItemMonth(item) === state.cancelMonth);
    const title = document.getElementById("cancelTableTitle");
    const eyebrow = document.getElementById("cancelTableEyebrow");
    if (title) title.textContent = `Cancelamentos - ${monthLabel(state.cancelMonth)}`;
    if (eyebrow) eyebrow.textContent = `${selectedRows.length} registros no mês`;
    renderTable("cancelTable", selectedRows, cancelColumns, 180, { totalField: "VALOR", totalColumnIndex: 5, totalRows: selectedRows });
  }
  function renderSalesWorkspace() {
    const rows = getSalesWorkspaceRows().filter(matchesFilters);
    ensureSalesMonth(rows);
    renderSalesMonthCards(rows);
    const selectedRows = rows.filter((item) => getItemMonth(item) === state.salesMonth);
    const tableTitle = state.salesMode === "pagos" ? "Pagos mensais" : "Vendas mensais";
    const summaryTitle = state.salesMode === "pagos" ? "Pagos x atrasos" : "Vendas por consultor";
    document.getElementById("salesTableTitle").textContent = `${tableTitle} - ${monthLabel(state.salesMonth)}`;
    document.getElementById("salesTableEyebrow").textContent = `${selectedRows.length} registros no mês`;
    document.getElementById("salesSummaryTitle").textContent = summaryTitle;

    if (state.salesMode === "pagos") {
      renderBars("consultantBars", groupSum(normalizeFinanceRows(selectedRows), "STATUS_FINANCEIRO", "VALOR"), formatCompactMoney, formatMoney);
    } else {
      renderBars("consultantBars", groupSum(individualSalesRows(selectedRows), "CONSULTOR_CANONICO", "VALOR"), formatCompactMoney, formatMoney);
    }

    const salesColumns = [
      { fields: ["TIPO_REGISTRO"], render: (row) => escapeHtml(getField(row, ["TIPO_REGISTRO"]) || "Venda/aditivo") },
      { fields: ["CONSULTOR_CANONICO", "CONSULTOR"], render: (row) => escapeHtml(getField(row, ["CONSULTOR_CANONICO", "CONSULTOR"]) || "-") },
      { fields: ["MES"], render: (row) => escapeHtml(monthLabel(getItemMonth(row))) },
      { fields: ["STATUS_FINANCEIRO", "FINANCEIRO"], render: (row) => statusPill(getField(row, ["STATUS_FINANCEIRO", "FINANCEIRO", "FINANCEIRO_2"]) || "Venda") },
      { fields: ["VALOR"], className: "num-cell", render: (row) => formatMoney(getField(row, ["VALOR"])) },
      { fields: ["EMPRESA"] },
      { fields: ["VENDA", "SERVICO", "SERVI_O"] }
    ];
    if (canManage()) {
      salesColumns.push({ fields: [state.salesMode === "pagos" ? "_receivableKey" : "_saleKey"], render: state.salesMode === "pagos" ? paidActions : saleActions });
    }
    renderTable("salesTable", selectedRows, salesColumns, 220, {
      totalField: "VALOR",
      totalColumnIndex: 4,
      totalRows: state.salesMode === "vendas" ? totalSalesRows(selectedRows) : selectedRows,
      rowClass: () => ""
    });
  }


  function matchesSummaryFilters(item, includeMonth = true) {
    const haystack = normalizeText(Object.values(item).join(" "));
    const statusText = normalizeText(getField(item, ["STATUS", "FINANCEIRO", "MOTIVO"]));
    const monthText = getItemMonth(item);
    const searchOk = !state.search || haystack.includes(normalizeText(state.search));
    const statusOk = !state.status || statusText.includes(normalizeText(state.status)) || haystack.includes(normalizeText(state.status));
    const monthOk = !includeMonth || !state.month || monthText === normalizeText(state.month);
    const systemOk = matchesOptionalField(item, ["SISTEMA"], state.system);
    const serviceOk = matchesOptionalField(item, ["SERVICO", "SERVI_O"], state.service);
    const financeOk = matchesOptionalField(item, ["FINANCEIRO", "FINANCEIRO_2"], state.finance);
    const consultantOk = matchesOptionalField(item, ["CONSULTOR_CANONICO", "CONSULTOR"], state.consultant);
    const reasonOk = matchesOptionalField(item, ["MOTIVO", "RECORRENCIA"], state.reason);
    return searchOk && statusOk && monthOk && systemOk && serviceOk && financeOk && consultantOk && reasonOk;
  }

  function buildMonthlySummaryRows(includeMonth = true) {
    const clients = dedupeClientsByMonth(baseRows).filter((item) => matchesSummaryFilters(item, includeMonth));
    const sales = totalSalesRows(individualSalesRows(applyEdits(salesMonthlyRows, saleKey, state.salesEdits, "_saleKey"))).filter((item) => matchesSummaryFilters(item, includeMonth));
    const paid = getPaidMonthlyWorkspaceRows().filter((item) => matchesSummaryFilters(item, includeMonth));
    const cancels = allCancelRows().filter((item) => matchesSummaryFilters(item, includeMonth));

    return monthOrder
      .map((month) => {
        const key = normalizeText(month);
        const clientsInMonth = clients.filter((item) => getItemMonth(item) === key);
        const salesInMonth = sales.filter((item) => getItemMonth(item) === key);
        const paidInMonth = paid.filter((item) => getItemMonth(item) === key);
        const cancelsInMonth = cancels.filter((item) => getItemMonth(item) === key);
        const paidRows = paidInMonth.filter((item) => normalizeText(getField(item, ["STATUS_FINANCEIRO", "FINANCEIRO", "FINANCEIRO_2"])).includes("PAGO"));
        const openRows = paidInMonth.filter((item) => !normalizeText(getField(item, ["STATUS_FINANCEIRO", "FINANCEIRO", "FINANCEIRO_2"])).includes("PAGO"));
        const cancelRows = cancelsInMonth.filter(isCancellationRecord);
        const downgradeRows = cancelsInMonth.filter(isDowngradeRecord);
        const vendas = sumBy(salesInMonth, "VALOR");
        const cancelamentos = sumBy(cancelRows, "VALOR");
        const downgrades = sumBy(downgradeRows, "VALOR");
        return {
          MES: month,
          CLIENTES: uniqueCount(clientsInMonth, "EMPRESA"),
          RECEITA_CARTEIRA: sumBy(clientsInMonth, "VALOR"),
          VENDAS: vendas,
          PAGOS: sumBy(paidRows, "VALOR"),
          ABERTO: sumBy(openRows, "VALOR"),
          CANCELAMENTOS: cancelamentos,
          DOWNGRADES: downgrades,
          SALDO: vendas - cancelamentos - downgrades,
          REGISTROS: clientsInMonth.length + salesInMonth.length + paidInMonth.length + cancelsInMonth.length
        };
      })
      .filter((item) => !state.month || normalizeText(item.MES) === normalizeText(state.month));
  }

  function renderSummaryMonthCards(rows) {
    const target = document.getElementById("summaryMonthCards");
    if (!target) return;
    const rowsByMonth = new Map(rows.map((item) => [normalizeText(item.MES), item]));
    target.innerHTML = monthOrder.map((month) => {
      const key = normalizeText(month);
      const item = rowsByMonth.get(key) || { VENDAS: 0, CANCELAMENTOS: 0, DOWNGRADES: 0, SALDO: 0, CLIENTES: 0 };
      const active = state.month && normalizeText(state.month) === key ? "active" : "";
      return `
        <button class="month-card ${active}" type="button" data-summary-month="${key}">
          <span>${escapeHtml(month)}</span>
          <strong>${formatCompactMoney(item.SALDO)}</strong>
          <small>${item.CLIENTES || 0} clientes / vendas ${formatCompactMoney(item.VENDAS)}</small>
        </button>
      `;
    }).join("");
  }

  function renderMonthlySummary() {
    const rows = buildMonthlySummaryRows();
    renderSummaryMonthCards(buildMonthlySummaryRows(false));
    const count = document.getElementById("summaryCount");
    if (count) count.textContent = `${rows.length} meses`;
    renderTable("summaryTable", rows, [
      { fields: ["MES"], render: (row) => escapeHtml(monthLabel(getField(row, ["MES"]))) },
      { fields: ["CLIENTES"], className: "num-cell" },
      { fields: ["RECEITA_CARTEIRA"], className: "num-cell", render: (row) => formatMoney(getField(row, ["RECEITA_CARTEIRA"])) },
      { fields: ["VENDAS"], className: "num-cell", render: (row) => formatMoney(getField(row, ["VENDAS"])) },
      { fields: ["PAGOS"], className: "num-cell", render: (row) => formatMoney(getField(row, ["PAGOS"])) },
      { fields: ["ABERTO"], className: "num-cell", render: (row) => formatMoney(getField(row, ["ABERTO"])) },
      { fields: ["CANCELAMENTOS"], className: "num-cell", render: (row) => formatMoney(getField(row, ["CANCELAMENTOS"])) },
      { fields: ["DOWNGRADES"], className: "num-cell", render: (row) => formatMoney(getField(row, ["DOWNGRADES"])) },
      { fields: ["SALDO"], className: "num-cell", render: (row) => formatMoney(getField(row, ["SALDO"])) }
    ], 12);
  }
  function renderStatusDonut(rows) {
    const total = Math.max(rows.length, 1);
    const active = rows.filter((item) => normalizeText(getField(item, ["STATUS"])).includes("ATIVO")).length;
    const late = rows.filter((item) => normalizeText(getField(item, ["FINANCEIRO", "FINANCEIRO_2"])).includes("ATRASO")).length;
    const blocked = blockRows.filter(matchesFilters).length;
    const other = Math.max(total - active, 0);
    const slices = [
      { label: "Ativos", value: active, color: "#64c8ff" },
      { label: "Em atraso", value: late, color: "#f2b35f" },
      { label: "Bloqueios", value: blocked, color: "#f07474" },
      { label: "Outros", value: other, color: "#8a5cf6" }
    ].filter((item) => item.value > 0);
    let current = 0;
    const gradient = slices.map((item) => {
      const start = current;
      const end = current + (item.value / total) * 100;
      current = end;
      return `${item.color} ${start}% ${end}%`;
    }).join(", ");
    const donut = document.getElementById("statusDonut");
    const legend = document.getElementById("statusLegend");
    if (donut) donut.style.background = `conic-gradient(${gradient || "rgba(255,255,255,0.08) 0 100%"})`;
    if (legend) {
      legend.innerHTML = slices.map((item) => `
        <div class="legend-item">
          <span class="legend-label"><span class="legend-dot" style="background:${item.color}"></span>${item.label}</span>
          <strong>${item.value}</strong>
        </div>
      `).join("");
    }
  }

  function renderMetrics() {
    const filteredBase = scopedBaseRows().filter(matchesFilters);
    const filteredCancels = allCancelRows().filter(matchesFilters);
    const filteredSales = individualSalesRows(salesMonthlyRows).filter(matchesFilters);
    const filteredSalesTotals = totalSalesRows(filteredSales);
    const overdue = filteredBase.filter((item) => normalizeText(getField(item, ["FINANCEIRO", "FINANCEIRO_2"])).includes("ATRASO"));
    const blocked = blockRows.filter(matchesFilters);
    const metrics = [
      ["Clientes na base", uniqueCount(filteredBase, "EMPRESA"), `${filteredBase.length} registros importados`],
      ["Receita ativa", formatMoney(sumBy(filteredBase, "VALOR")), "Soma da base principal"],
      ["Em atraso", formatMoney(sumBy(overdue, "VALOR")), `${overdue.length} clientes com alerta`],
      ["Bloqueios", blocked.length, "Parcial ou total"],
      ["Vendas", formatMoney(sumBy(filteredSalesTotals, "VALOR")), `${filteredCancels.length} cancelamentos/downgrades`]
    ];
    document.getElementById("metricGrid").innerHTML = metrics.map(([label, value, helper]) => `
      <article class="metric">
        <span>${label}</span>
        <strong>${value}</strong>
        <small>${helper}</small>
      </article>
    `).join("");
  }

  function renderAll() {
    const currentBaseRows = scopedBaseRows();
    const filteredBase = currentBaseRows.filter(matchesFilters);
    const filteredSales = individualSalesRows(salesMonthlyRows).filter(matchesFilters);
    const filteredSalesTotals = totalSalesRows(filteredSales);
    const filteredSalesIndividuals = individualSalesRows(filteredSales);
    const filteredCancels = allCancelRows().filter(matchesFilters);
    renderMetrics();
    renderStatusDonut(filteredBase);
    renderBars("salesBars", groupSum(filteredSalesTotals, "MES", "VALOR").sort((a, b) => monthIndex(a.label) - monthIndex(b.label)), formatCompactMoney, formatMoney);
    renderBars("serviceBars", groupCount(filteredBase, "SERVICO"), String);
    renderBars("cancelReasonBars", groupSum(filteredCancels, "MOTIVO", "VALOR"), formatCompactMoney, formatMoney);
    renderBars("systemRevenueBars", groupSum(filteredBase, "SISTEMA", "VALOR"), formatCompactMoney, formatMoney);
    renderTopClients(filteredBase);
    renderSalesCancelCompare(filteredSalesTotals, filteredCancels);
    renderBars("financeBars", groupSum(filteredBase, "FINANCEIRO", "VALOR"), formatCompactMoney, formatMoney);
    renderBars("cancelConsultantBars", groupSum(filteredCancels, "CONSULTOR_CANONICO", "VALOR"), formatCompactMoney, formatMoney);
    renderSalesWorkspace();

    document.getElementById("clientCount").textContent = filteredBase.length;
    const clientColumns = [
      { fields: ["EMPRESA"] },
      { fields: ["FILIAL"] },
      { fields: ["SISTEMA"] },
      { fields: ["SERVICO"], render: (row) => escapeHtml(getField(row, ["SERVICO", "SERVI_O"])) },
      { fields: ["STATUS"], render: (row) => statusPill(getField(row, ["STATUS"])) },
      { fields: ["VALOR"], className: "num-cell", render: (row) => formatMoney(getField(row, ["VALOR"])) },
      { fields: ["VENCIMENTO"], className: "nowrap", render: (row) => escapeHtml(excelDate(getField(row, ["VENCIMENTO"]))) }
    ];
    if (canManage()) {
      clientColumns.push({ fields: ["_rowKey"], render: clientActions });
    }
    renderTable("clientsTable", currentBaseRows, clientColumns, 180, { totalField: "VALOR", totalColumnIndex: 5 });
    renderTable("financeTable", currentBaseRows.concat(blockRows), [
      { fields: ["EMPRESA"] },
      { fields: ["FINANCEIRO"], render: (row) => statusPill(getField(row, ["FINANCEIRO", "FINANCEIRO_2"])) },
      { fields: ["STATUS"], render: (row) => statusPill(getField(row, ["STATUS"])) },
      { fields: ["VALOR"], className: "num-cell", render: (row) => formatMoney(getField(row, ["VALOR"])) },
      { fields: ["VENCIMENTO"], className: "nowrap", render: (row) => escapeHtml(excelDate(getField(row, ["VENCIMENTO"]))) },
      { fields: ["SISTEMA"] }
    ], 180, { totalField: "VALOR", totalColumnIndex: 3 });
    const cancelColumns = [
      { fields: ["EMPRESA"] },
      { fields: ["MES"], render: (row) => escapeHtml(getField(row, ["MES"]) || excelDate(getField(row, ["VENCIMENTO"]))) },
      { fields: ["STATUS"], render: (row) => statusPill(getField(row, ["STATUS"])) },
      { fields: ["MOTIVO", "RECORRENCIA"] },
      { fields: ["CONSULTOR_CANONICO", "CONSULTOR"], render: (row) => escapeHtml(getField(row, ["CONSULTOR_CANONICO", "CONSULTOR"])) },
      { fields: ["VALOR"], className: "num-cell", render: (row) => formatMoney(getField(row, ["VALOR"])) }
    ];
    if (canManage()) cancelColumns.push({ fields: ["_cancelKey"], render: cancelActions });
    renderCancelWorkspace(cancelColumns);
    renderMonthlySummary();
  }

  function monthIndex(label) {
    const index = monthOrder.findIndex((month) => normalizeText(month) === normalizeText(label));
    return index === -1 ? 99 : index;
  }

  function setupMonths() {
    const months = new Set();
    baseRows.concat(blockRows, salesRows, allCancelRows()).forEach((item) => {
      const month = getItemMonth(item);
      if (month) months.add(String(month).trim());
    });
    const select = document.getElementById("monthFilter");
    if (!select) return;
    select.querySelectorAll("option:not(:first-child)").forEach((option) => option.remove());
    [...months].sort((a, b) => monthIndex(a) - monthIndex(b)).forEach((month) => {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = monthOrder.find((item) => normalizeText(item) === normalizeText(month)) || month;
      select.appendChild(option);
    });
  }

  function setupSelectOptions() {
    fillSelect("systemFilter", uniqueValues(baseRows, ["SISTEMA"]));
    fillSelect("serviceFilter", uniqueValues(baseRows, ["SERVICO", "SERVI_O"]));
    fillSelect("financeFilter", uniqueValues(baseRows.concat(blockRows), ["FINANCEIRO", "FINANCEIRO_2"]));
    fillSelect("consultantFilter", uniqueValues(individualSalesRows(salesRows).concat(allCancelRows()), ["CONSULTOR_CANONICO", "CONSULTOR"]));
    fillSelect("reasonFilter", uniqueValues(allCancelRows(), ["MOTIVO", "RECORRENCIA"]));
  }

  function uniqueValues(rows, fields) {
    const values = new Map();
    rows.forEach((item) => {
      const value = String(getField(item, fields) || "").trim();
      if (!value) return;
      values.set(normalizeText(value), value);
    });
    return [...values.values()].sort((a, b) => normalizeText(a).localeCompare(normalizeText(b)));
  }

  function fillSelect(id, values) {
    const select = document.getElementById(id);
    if (!select) return;
    select.querySelectorAll("option:not(:first-child)").forEach((option) => option.remove());
    values.forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });
  }

  function setView(view) {
    state.view = view;
    document.querySelectorAll(".view").forEach((item) => item.classList.toggle("active", item.id === view));
    document.querySelectorAll(".nav-btn").forEach((item) => item.classList.toggle("active", item.dataset.view === view));
    document.getElementById("actionsMenu")?.classList.remove("open");
    document.getElementById("actionsMenuBtn")?.setAttribute("aria-expanded", "false");
    applyAccessState();
  }

  function applyAccessState() {
    document.body.classList.toggle("logged-out", !state.session);
    document.body.classList.toggle("is-admin", canManage());
    document.querySelectorAll(".admin-only").forEach((item) => {
      item.hidden = !canManage();
    });
    const roleBadge = document.getElementById("roleBadge");
    if (roleBadge) roleBadge.textContent = state.session?.label || "Sem acesso";
    const importBtn = document.getElementById("importBtn");
    if (importBtn) {
      const canImportView = ["clientes", "financeiro", "vendas", "cancelamentos"].includes(state.view);
      importBtn.hidden = !(canManage() && canImportView);
      importBtn.textContent = state.view === "vendas" && state.salesMode === "pagos"
        ? "Importar pagos"
        : state.view === "vendas"
          ? "Importar vendas"
          : state.view === "cancelamentos"
            ? "Importar cancelamentos"
            : "Importar clientes";
    }
  }

  async function loadSupabaseProfile(authUser) {
    if (!supabaseClient || !authUser) return null;
    const { data, error } = await supabaseClient
      .from("usuarios_sistema")
      .select("nome, email, perfil, ativo")
      .eq("user_id", authUser.id)
      .eq("ativo", true)
      .single();

    if (error || !data) throw new Error("Perfil sem acesso ao sistema.");

    const isAdmin = data.perfil === "admin";
    return {
      userId: authUser.id,
      user: data.email || authUser.email,
      email: data.email || authUser.email,
      role: isAdmin ? "admin" : "viewer",
      label: isAdmin ? "Admin" : "Visualizador",
      name: data.nome || authUser.email
    };
  }

  function dbMonthlyToApp(row) {
    return normalizeManualClient({
      EMPRESA: row.empresa,
      FILIAL: row.filial || "",
      SISTEMA: row.sistema || "",
      SERVICO: row.servico || "",
      VALOR: row.valor_mensal || 0,
      VENCIMENTO: row.vencimento || "",
      ENTRADA: row.entrada || "",
      STATUS: appClientStatus(row.status),
      FINANCEIRO: appFinanceStatus(row.financeiro),
      MES_REFERENCIA: monthFromCompetencia(row.competencia),
      ORIGEM: "Supabase",
      _manual: true,
      _supabaseClientId: row.cliente_id,
      _supabaseMonthlyId: row.id,
      _rowKey: `client-${row.id}`
    });
  }

  function dbSaleToApp(row) {
    return normalizeManualSale({
      EMPRESA: row.empresa,
      SISTEMA: row.sistema || "",
      VENDA: row.servico || row.tipo || "Venda",
      CONSULTOR: row.consultor || "Não informado",
      CONSULTOR_CANONICO: canonicalConsultant(row.consultor || "Não informado"),
      MES: monthFromCompetencia(row.competencia),
      VALOR: row.valor || 0,
      TIPO_REGISTRO: row.tipo === "aditivo" ? "Aditivo" : "Venda/aditivo",
      STATUS_FINANCEIRO: "Venda",
      ORIGEM: "Supabase",
      _manualSale: true,
      _supabaseId: row.id,
      _supabaseClientId: row.cliente_id,
      _saleKey: `sale-${row.id}`
    });
  }

  function dbReceivableToApp(row) {
    return normalizeManualReceivable({
      EMPRESA: row.empresa,
      SISTEMA: row.sistema || "",
      VENDA: row.servico || "Mensalidade",
      CONSULTOR: "",
      CONSULTOR_CANONICO: "",
      MES: monthFromCompetencia(row.competencia),
      VALOR: row.valor_final ?? row.valor_original ?? 0,
      VENCIMENTO: row.vencimento || "",
      TIPO_REGISTRO: "Mensalidade",
      STATUS_FINANCEIRO: appFinanceStatus(row.status),
      ORIGEM: "Supabase",
      _manualReceivable: true,
      _supabaseId: row.id,
      _supabaseClientId: row.cliente_id,
      _receivableKey: `receivable-${row.id}`
    });
  }

  function dbCancelToApp(row) {
    return normalizeManualCancel({
      EMPRESA: row.empresa,
      SISTEMA: row.sistema || "",
      SERVICO: row.servico || "",
      MES: monthFromCompetencia(row.competencia),
      STATUS: row.tipo === "downgrade" ? "DOWNGRADE" : "CANCELADO",
      MOTIVO: row.motivo || (row.tipo === "downgrade" ? "Downgrade" : "Cancelamento"),
      CONSULTOR: row.consultor || "POS",
      CONSULTOR_CANONICO: canonicalConsultant(row.consultor || "POS"),
      VALOR: row.valor || 0,
      ORIGEM: "Supabase",
      _manualCancel: true,
      _supabaseId: row.id,
      _supabaseClientId: row.cliente_id,
      _cancelKey: `cancel-${row.id}`
    });
  }

  async function loadSupabaseData() {
    if (!supabaseClient || !state.session) return;
    const [monthly, sales, receivables, cancels] = await Promise.all([
      supabaseClient.from("clientes_mensais").select("*").order("competencia", { ascending: true }),
      supabaseClient.from("vendas").select("*").order("competencia", { ascending: true }),
      supabaseClient.from("pagos_mensais").select("*").order("competencia", { ascending: true }),
      supabaseClient.from("cancelamentos").select("*").order("competencia", { ascending: true })
    ]);
    const error = monthly.error || sales.error || receivables.error || cancels.error;
    if (error) throw error;

    state.manual = (monthly.data || []).map(dbMonthlyToApp);
    state.manualSales = (sales.data || []).map(dbSaleToApp);
    state.manualReceivables = (receivables.data || []).map(dbReceivableToApp);
    state.manualCancels = (cancels.data || []).map(dbCancelToApp);

    salesMonthlyRows.length = 0;
    salesMonthlyRows.push(...state.manualSales);
    salesRows.length = 0;
    salesRows.push(...state.manualSales);
    paidMonthlyRows.length = 0;
    paidMonthlyRows.push(...state.manualReceivables);
    cancelRows.length = 0;
    cancelRows.push(...state.manualCancels);

    state.overrides = {};
    state.salesEdits = {};
    state.receivableEdits = {};
    state.cancelEdits = {};
    state.receivableStatus = {};
    refreshBaseRows();
    setupMonths();
    setupSelectOptions();
  }

  async function restoreSupabaseSession() {
    if (!supabaseClient) return;
    const { data } = await supabaseClient.auth.getSession();
    if (!data.session?.user) return;
    try {
      state.session = await loadSupabaseProfile(data.session.user);
      saveSession();
      await loadSupabaseData();
    } catch (error) {
      state.session = null;
      saveSession();
      await supabaseClient.auth.signOut();
    }
  }

  async function login(user, password) {
    if (supabaseClient) {
      const { data, error } = await supabaseClient.auth.signInWithPassword({
        email: String(user || "").trim(),
        password: String(password || "")
      });
      if (error || !data.user) return false;
      state.session = await loadSupabaseProfile(data.user);
      saveSession();
      await loadSupabaseData();
      applyAccessState();
      closeCancelModal(false);
      renderAll();
      return true;
    }

    const account = users[normalizeText(user).toLowerCase()];
    if (!account || account.password !== password) return false;
    state.session = { user: normalizeText(user).toLowerCase(), role: account.role, label: account.label };
    saveSession();
    applyAccessState();
    renderAll();
    return true;
  }

  async function logout() {
    if (supabaseClient) await supabaseClient.auth.signOut();
    state.session = null;
    saveSession();
    applyAccessState();
  }

  function bindEvents() {
    setAdvancedFiltersOpen(false);
    ensureSalesModal();
    document.getElementById("sidebarToggle")?.addEventListener("click", () => {
      document.body.classList.toggle("sidebar-collapsed");
    });
    document.querySelectorAll("[data-view], [data-view-target]").forEach((button) => {
      button.addEventListener("click", () => setView(button.dataset.view || button.dataset.viewTarget));
    });
    document.getElementById("loginForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const formElement = event.currentTarget;
      const form = new FormData(formElement);
      const feedback = document.getElementById("loginFeedback");
      const submitButton = formElement.querySelector("button[type='submit']");
      feedback.textContent = "Entrando...";
      submitButton.disabled = true;
      try {
        if (await login(form.get("user"), form.get("password"))) {
          formElement.reset();
          feedback.textContent = "";
        } else {
          feedback.textContent = "E-mail ou senha inválidos.";
        }
      } catch (error) {
        feedback.textContent = error.message || "Não foi possível entrar.";
      } finally {
        submitButton.disabled = false;
      }
    });
    document.getElementById("logoutBtn").addEventListener("click", logout);
    document.getElementById("filtersToggleBtn")?.addEventListener("click", () => {
      const panel = document.getElementById("advancedFilters");
      setAdvancedFiltersOpen(panel ? panel.hidden : true);
    });
    document.getElementById("actionsMenuBtn").addEventListener("click", (event) => {
      event.stopPropagation();
      const menu = document.getElementById("actionsMenu");
      const isOpen = menu.classList.toggle("open");
      event.currentTarget.setAttribute("aria-expanded", String(isOpen));
    });
    document.addEventListener("click", (event) => {
      const menu = document.getElementById("actionsMenu");
      if (!menu.contains(event.target)) {
        menu.classList.remove("open");
        document.getElementById("actionsMenuBtn").setAttribute("aria-expanded", "false");
      }
    });
    document.getElementById("exportBtn").addEventListener("click", exportCurrentView);
    document.getElementById("importBtn").addEventListener("click", () => {
      document.getElementById("importFile").click();
    });
    document.getElementById("importFile").addEventListener("change", (event) => {
      importCurrentViewFile(event.target.files[0]);
      event.target.value = "";
    });
    document.getElementById("clearAllBtn")?.addEventListener("click", clearAllData);
    document.getElementById("cancelForm")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!canManage()) return;
      const payload = cancelPayloadFromForm(new FormData(event.currentTarget));
      const editingCancelRow = state.editingCancelKey ? allCancelRows().find((item) => (item._cancelKey || cancelKey(item)) === state.editingCancelKey) : null;
      if (editingCancelRow) {
        payload._supabaseId = editingCancelRow._supabaseId;
        payload._supabaseClientId = editingCancelRow._supabaseClientId;
      }
      const normalized = state.manualCancels.map(normalizeManualCancel);
      if (state.editingCancelKey) {
        const existsManual = normalized.some((item) => item._cancelKey === state.editingCancelKey);
        if (existsManual) {
          state.manualCancels = normalized.map((item) => item._cancelKey === state.editingCancelKey ? payload : item);
          persistManualCancels();
        } else {
          state.cancelEdits[state.editingCancelKey] = { data: payload };
          saveEdits();
        }
      } else {
        state.manualCancels = [payload].concat(normalized);
        persistManualCancels();
      }
      try {
        await saveSupabaseCancel(payload);
      } catch (error) {
        alert(error.message || "Não foi possível salvar o cancelamento no banco.");
      }
      state.editingCancelKey = "";
      event.currentTarget.reset();
      event.currentTarget.querySelector("button[type='submit']").textContent = "Salvar cancelamento";
      renderAll();
    });
    document.getElementById("clearCancelsBtn")?.addEventListener("click", clearManualCancels);
    document.getElementById("newCancelBtn")?.addEventListener("click", () => {
      if (!canManage()) return;
      state.editingCancelKey = "";
      document.getElementById("cancelForm")?.reset();
      document.querySelector("#cancelForm button[type='submit']").textContent = "Salvar cancelamento";
      openCancelModal("Novo cancelamento");
    });
    document.getElementById("cancelModalCloseBtn")?.addEventListener("click", () => closeCancelModal());
    document.getElementById("cancelModalXBtn")?.addEventListener("click", () => closeCancelModal());
    document.getElementById("cancelModal")?.addEventListener("click", (event) => {
      if (event.target.id === "cancelModal") closeCancelModal();
    });
    document.getElementById("cancelTable")?.addEventListener("click", (event) => {
      const editButton = event.target.closest("[data-cancel-edit]");
      const deleteButton = event.target.closest("[data-cancel-delete]");
      if (editButton) editCancel(editButton.dataset.cancelEdit);
      if (deleteButton) deleteCancel(deleteButton.dataset.cancelDelete);
    });
    document.getElementById("newClientBtn")?.addEventListener("click", () => {
      if (!canManage()) return;
      state.editingKey = "";
      document.getElementById("clientForm")?.reset();
      document.querySelector("#clientForm button[type='submit']").textContent = "Salvar cliente";
      openClientModal("Novo cliente");
    });
    document.getElementById("clientModalCloseBtn")?.addEventListener("click", () => closeClientModal());
    document.getElementById("clientModalXBtn")?.addEventListener("click", () => closeClientModal());
    document.getElementById("clientModal")?.addEventListener("click", (event) => {
      if (event.target.id === "clientModal") closeClientModal();
    });
    document.getElementById("clientsTable").addEventListener("click", (event) => {
      const editButton = event.target.closest("[data-client-edit]");
      const deleteButton = event.target.closest("[data-client-delete]");
      if (editButton) editClient(editButton.dataset.clientEdit);
      if (deleteButton) deleteClient(deleteButton.dataset.clientDelete);
    });
    document.getElementById("newSaleBtn")?.addEventListener("click", () => {
      if (!canManage()) return;
      state.editingSaleKey = "";
      state.editingReceivableKey = "";
      document.getElementById("saleForm")?.reset();
      document.querySelector("#saleForm button[type='submit']").textContent = "Salvar venda";
      openSaleModal('Nova venda ou aditivo');
    });
    document.getElementById("saleModalCloseBtn")?.addEventListener("click", () => closeSaleModal());
    document.getElementById("saleModalXBtn")?.addEventListener("click", () => closeSaleModal());
    document.getElementById("saleModal")?.addEventListener("click", (event) => {
      if (event.target.id === "saleModal") closeSaleModal();
    });
    document.getElementById("salesTable").addEventListener("change", (event) => {
      const statusSelect = event.target.closest("[data-receivable-status]");
      if (!statusSelect) return;
      setReceivableStatus(statusSelect.dataset.receivableStatus, statusSelect.value);
    });
    document.getElementById("salesTable").addEventListener("click", (event) => {
      const saleEdit = event.target.closest("[data-sale-edit]");
      const saleDelete = event.target.closest("[data-sale-delete]");
      const receivableEdit = event.target.closest("[data-receivable-edit]");
      const receivableDelete = event.target.closest("[data-receivable-delete]");
      if (saleEdit) editSale(saleEdit.dataset.saleEdit);
      if (saleDelete) deleteSale(saleDelete.dataset.saleDelete);
      if (receivableEdit) editReceivable(receivableEdit.dataset.receivableEdit);
      if (receivableDelete) deleteReceivable(receivableDelete.dataset.receivableDelete);
    });
    document.querySelectorAll("[data-sales-mode]").forEach((button) => {
      button.addEventListener("click", () => {
        state.salesMode = button.dataset.salesMode;
        document.querySelectorAll("[data-sales-mode]").forEach((item) => item.classList.toggle("active", item === button));
        renderAll();
      });
    });
    document.getElementById("salesMonthCards").addEventListener("click", (event) => {
      const card = event.target.closest("[data-sales-month]");
      if (!card) return;
      state.salesMonth = card.dataset.salesMonth;
      renderAll();
    });
    document.getElementById("summaryMonthCards")?.addEventListener("click", (event) => {
      const card = event.target.closest("[data-summary-month]");
      if (!card) return;
      state.month = state.month === card.dataset.summaryMonth ? "" : card.dataset.summaryMonth;
      document.getElementById("monthFilter").value = state.month;
      state.salesMonth = state.month;
      state.cancelMonth = state.month;
      renderAll();
    });
    document.getElementById("cancelMonthCards")?.addEventListener("click", (event) => {
      const card = event.target.closest("[data-cancel-month]");
      if (!card) return;
      state.cancelMonth = card.dataset.cancelMonth;
      renderAll();
    });
    document.getElementById("searchInput").addEventListener("input", (event) => {
      state.search = event.target.value;
      renderAll();
    });
    document.getElementById("statusFilter").addEventListener("change", (event) => {
      state.status = event.target.value;
      renderAll();
    });
    document.getElementById("monthFilter").addEventListener("change", (event) => {
      state.month = event.target.value;
      state.salesMonth = event.target.value ? normalizeText(event.target.value) : "";
      state.cancelMonth = event.target.value ? normalizeText(event.target.value) : "";
      renderAll();
    });
    const showDupEl = document.getElementById("showDuplicates");
    if (showDupEl) {
      showDupEl.checked = state.showDuplicates;
      showDupEl.addEventListener("change", (e) => {
        state.showDuplicates = e.target.checked;
        renderAll();
      });
    }
    document.getElementById("systemFilter").addEventListener("change", (event) => {
      state.system = event.target.value;
      renderAll();
    });
    document.getElementById("serviceFilter").addEventListener("change", (event) => {
      state.service = event.target.value;
      renderAll();
    });
    document.getElementById("financeFilter").addEventListener("change", (event) => {
      state.finance = event.target.value;
      renderAll();
    });
    document.getElementById("consultantFilter").addEventListener("change", (event) => {
      state.consultant = event.target.value;
      renderAll();
    });
    document.getElementById("reasonFilter").addEventListener("change", (event) => {
      state.reason = event.target.value;
      renderAll();
    });
    document.getElementById("rankingFilter").addEventListener("change", (event) => {
      state.ranking = Number(event.target.value) || 10;
      renderAll();
    });
    document.getElementById("clearFiltersBtn").addEventListener("click", () => {
      state.search = "";
      state.status = "";
      state.month = "";
      state.system = "";
      state.service = "";
      state.finance = "";
      state.consultant = "";
      state.reason = "";
      state.ranking = 10;
      state.salesMonth = "";
      state.cancelMonth = "";
      document.getElementById("searchInput").value = "";
      document.getElementById("statusFilter").value = "";
      document.getElementById("monthFilter").value = "";
      document.getElementById("systemFilter").value = "";
      document.getElementById("serviceFilter").value = "";
      document.getElementById("financeFilter").value = "";
      const consultantFilter = document.getElementById("consultantFilter");
      if (consultantFilter) consultantFilter.value = "";
      document.getElementById("reasonFilter").value = "";
      document.getElementById("rankingFilter").value = "10";
      renderAll();
    });
    document.getElementById("clientForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!canManage()) return;
      const form = new FormData(event.currentTarget);
      const payload = clientPayloadFromForm(form);
      let savedClient = null;
      if (state.editingKey) {
        const row = baseRows.find((item) => (item._rowKey || clientKey(item)) === state.editingKey);
        if (row?._manual) {
          savedClient = { ...row, ...payload, _manual: true, _rowKey: state.editingKey };
          state.manual = state.manual.map((item) => (item._rowKey || clientKey(item)) === state.editingKey ? savedClient : item);
          saveManual();
        } else {
          savedClient = { ...row, ...payload };
          state.overrides[state.editingKey] = { ...(state.overrides[state.editingKey] || {}), data: payload };
          saveOverrides();
        }
        state.editingKey = "";
        event.currentTarget.querySelector("button[type='submit']").textContent = "Salvar cliente";
      } else {
        const client = normalizeManualClient({ ...payload, _manual: true });
        savedClient = client;
        state.manual.unshift(client);
        saveManual();
      }
      try {
        if (savedClient) await saveSupabaseClient(savedClient);
      } catch (error) {
        alert(error.message || "Não foi possível salvar o cliente no banco.");
      }
      refreshBaseRows();
      event.currentTarget.reset();
      closeClientModal(false);
      renderAll();
    });
    document.getElementById("saleForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!canManage()) return;
      const form = new FormData(event.currentTarget);
      let savedSale = null;
      let savedReceivable = null;
      if (state.editingSaleKey) {
        const payload = salePayloadFromForm(form, "sale");
        const editingRow = getSalesWorkspaceRows().find((item) => (item._saleKey || saleKey(item)) === state.editingSaleKey);
        savedSale = { ...editingRow, ...payload };
        state.salesEdits[state.editingSaleKey] = { data: savedSale };
        state.manualSales = state.manualSales.map((item) => (item._saleKey || saleKey(item)) === state.editingSaleKey ? savedSale : item);
        saveManualSales();
        saveEdits();
        state.salesMonth = normalizeText(payload.MES);
        state.editingSaleKey = "";
      } else if (state.editingReceivableKey) {
        const payload = salePayloadFromForm(form, "receivable");
        const editingRow = getSalesWorkspaceRows().find((item) => (item._receivableKey || receivableKey(item)) === state.editingReceivableKey);
        savedReceivable = { ...editingRow, ...payload };
        state.receivableEdits[state.editingReceivableKey] = { data: savedReceivable };
        state.manualReceivables = state.manualReceivables.map((item) => (item._receivableKey || receivableKey(item)) === state.editingReceivableKey ? savedReceivable : item);
        saveManualReceivables();
        saveEdits();
        state.salesMonth = normalizeText(payload.MES);
        state.editingReceivableKey = "";
      } else {
        const sale = salePayloadFromForm(form, "sale");
        const receivable = {
          ...sale,
          TIPO_REGISTRO: "Mensalidade",
          STATUS_FINANCEIRO: "Pendente",
          VENDA: `${form.get("venda")} - pendente`,
          _manualReceivable: true,
          _receivableKey: `receivable-${Date.now()}-${Math.random().toString(16).slice(2)}`
        };
        savedSale = sale;
        savedReceivable = receivable;
        state.manualSales.unshift(sale);
        state.manualReceivables.unshift(receivable);
        salesMonthlyRows.unshift(sale);
        salesRows.unshift(sale);
        paidMonthlyRows.unshift(receivable);
        saveManualSales();
        saveManualReceivables();
        state.salesMode = "vendas";
        state.salesMonth = normalizeText(sale.MES);
      }
      try {
        if (savedSale) await saveSupabaseSale(savedSale);
        if (savedReceivable) await saveSupabaseReceivable(savedReceivable);
      } catch (error) {
        alert(error.message || "Não foi possível salvar vendas/pagos no banco.");
      }
      document.querySelectorAll("[data-sales-mode]").forEach((item) => item.classList.toggle("active", item.dataset.salesMode === state.salesMode));
      event.currentTarget.querySelector("button[type='submit']").textContent = "Salvar venda";
      event.currentTarget.reset();
      closeSaleModal(false);
      renderAll();
    });
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (char) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "\"": "&quot;",
      "'": "&#039;"
    }[char]));
  }

  setupMonths();
  setupSelectOptions();
  async function init() {
    bindEvents();
    document.body.classList.add("loading");
    try {
      await restoreSupabaseSession();
    } catch (error) {
      console.error(error);
      state.session = null;
      saveSession();
    } finally {
      document.body.classList.remove("loading");
      applyAccessState();
      renderAll();
    }
  }

  init();
})();



















































