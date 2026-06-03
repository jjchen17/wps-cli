/* WPS CLI Bridge — JS Addon
 *
 * 运行在 WPS 内嵌浏览器中，通过 WebSocket 接收 Python bridge 发来的命令，
 * 调用 WPS JS API 执行操作，返回结果。
 *
 * 协议：
 *   连接成功后发送 {"type":"hello", "app":"writer|calc|impress"}
 *   接收 {"id":"...", "method":"writer.info", "params":{...}}
 *   返回 {"id":"...", "ok":true, "data":{...}} 或 {"id":"...", "ok":false, "error":"..."}
 */

(function () {
  "use strict";

  var BRIDGE_HOST = "127.0.0.1";
  var BRIDGE_PORT = 3890;
  var ws = null;
  var appType = "unknown";
  var reconnectTimer = null;

  // ── App Detection ──────────────────────────────────────────

  function detectApp() {
    try {
      if (
        typeof Application.ActiveDocument !== "undefined" &&
        typeof Application.Documents !== "undefined"
      ) {
        return "writer";
      }
    } catch (e) {}
    try {
      if (
        typeof Application.ActiveWorkbook !== "undefined" ||
        typeof Application.Workbooks !== "undefined"
      ) {
        return "calc";
      }
    } catch (e) {}
    try {
      if (typeof Application.ActivePresentation !== "undefined") {
        return "impress";
      }
    } catch (e) {}
    return "unknown";
  }

  // ── WebSocket ──────────────────────────────────────────────

  function connect() {
    if (ws && ws.readyState === WebSocket.OPEN) return;
    try {
      ws = new WebSocket("ws://" + BRIDGE_HOST + ":" + BRIDGE_PORT + "/ws");
    } catch (e) {
      scheduleReconnect();
      return;
    }

    ws.onopen = function () {
      appType = detectApp();
      ws.send(JSON.stringify({ type: "hello", app: appType }));
      if (reconnectTimer) {
        clearTimeout(reconnectTimer);
        reconnectTimer = null;
      }
    };

    ws.onmessage = function (event) {
      var cmd;
      try {
        cmd = JSON.parse(event.data);
      } catch (e) {
        return;
      }
      // 忽略非命令消息（如握手响应）
      if (!cmd.method) return;

      var result = dispatch(cmd);
      try {
        ws.send(JSON.stringify(result));
      } catch (e) {
        // 连接已断开
      }
    };

    ws.onclose = function () {
      scheduleReconnect();
    };

    ws.onerror = function () {
      // onclose 会紧随其后触发
    };
  }

  function scheduleReconnect() {
    if (reconnectTimer) return;
    reconnectTimer = setTimeout(function () {
      reconnectTimer = null;
      connect();
    }, 2000);
  }

  // ── Command Dispatcher ─────────────────────────────────────

  function dispatch(cmd) {
    var id = cmd.id;
    var method = cmd.method;
    var params = cmd.params || {};

    try {
      var data = execute(method, params);
      return { id: id, ok: true, data: data };
    } catch (e) {
      return { id: id, ok: false, error: e.message || String(e) };
    }
  }

  function execute(method, params) {
    var parts = method.split(".");
    var app = parts[0];
    var action = parts.slice(1).join(".");

    if (app === "writer") return writerAction(action, params);
    if (app === "calc") return calcAction(action, params);
    if (app === "impress") return impressAction(action, params);
    throw new Error("Unknown app: " + app);
  }

  // ── Helpers ────────────────────────────────────────────────

  function toPath(p) {
    // JS API expects paths without file:// prefix
    return String(p).replace(/^file:\/\//, "");
  }

  function safeStr(v) {
    if (v === null || v === undefined) return "";
    return String(v);
  }

  // ── Writer Actions ─────────────────────────────────────────

  function writerAction(action, p) {
    switch (action) {
      case "info":
        return writerInfo(p.path);
      case "replace":
        return writerReplace(p.path, p.old, p["new"], p.wildcard, p.case);
      case "count":
        return writerCount(p.path);
      case "table_insert":
        return writerTableInsert(p.path, p.rows, p.cols, p.data);
      case "table_get":
        return writerTableGet(p.path, p.index);
      case "image_insert":
        return writerImageInsert(p.path, p.image, p.width, p.height);
      case "page_setup":
        return writerPageSetup(
          p.path,
          p.width,
          p.height,
          p.margin_top,
          p.margin_bottom,
          p.margin_left,
          p.margin_right
        );
      case "export_pdf":
        return writerExportPdf(p.path, p.output);
      default:
        throw new Error("Unknown writer action: " + action);
    }
  }

  function _openDoc(path, readonly) {
    var fp = toPath(path);
    if (readonly) {
      return Application.Documents.Open(fp, undefined, true);
    }
    return Application.Documents.Open(fp);
  }

  function writerInfo(path) {
    var doc = _openDoc(path, true);
    try {
      var props = doc.BuiltInDocumentProperties;
      return {
        path: String(doc.FullName),
        pages: doc.ComputeStatistics(2), // wdStatisticPages
        words: doc.ComputeStatistics(0), // wdStatisticWords
        characters: doc.ComputeStatistics(3), // wdStatisticCharacters
        paragraphs: doc.Paragraphs.Count,
        author: safeStr(props("Author").Value),
        created: safeStr(props("Creation Date").Value),
        modified: safeStr(props("Last Save Time").Value),
      };
    } finally {
      doc.Close(0); // wdDoNotSaveChanges
    }
  }

  function writerReplace(path, oldText, newText, wildcard, matchCase) {
    var doc = _openDoc(path, false);
    try {
      var find = doc.Content.Find;
      find.ClearFormatting();
      find.Text = oldText;
      find.Replacement.Text = newText;
      find.MatchWildcards = !!wildcard;
      find.MatchCase = !!matchCase;
      find.Forward = true;
      find.Wrap = 0; // wdFindStop

      var count = 0;
      if (!wildcard) {
        var scan = doc.Content.Find;
        scan.ClearFormatting();
        scan.Text = oldText;
        scan.MatchCase = !!matchCase;
        scan.Forward = true;
        scan.Wrap = 0;
        /* eslint-disable no-empty */
        while (scan.Execute(undefined, undefined, undefined, undefined, undefined,
          undefined, undefined, undefined, undefined, undefined, 0)) {
          count++;
        }
      }
      find.Execute(
        undefined, undefined, undefined, undefined, undefined, undefined,
        undefined, undefined, undefined, undefined, 2 // wdReplaceAll
      );
      doc.Save();
      return { replaced: wildcard ? -1 : count };
    } finally {
      doc.Close(0);
    }
  }

  function writerCount(path) {
    var doc = _openDoc(path, true);
    try {
      return {
        words: doc.ComputeStatistics(0),
        characters: doc.ComputeStatistics(3),
        paragraphs: doc.Paragraphs.Count,
        pages: doc.ComputeStatistics(2),
      };
    } finally {
      doc.Close(0);
    }
  }

  function writerTableInsert(path, rows, cols, data) {
    var doc = _openDoc(path, false);
    try {
      var rng = doc.Content;
      rng.Collapse(0); // wdCollapseEnd
      var table = doc.Tables.Add(rng, rows, cols);
      try {
        table.Borders.Enable = true;
      } catch (e) {}
      if (data) {
        for (var i = 0; i < data.length; i++) {
          for (var j = 0; j < (data[i] || []).length; j++) {
            if (i < rows && j < cols) {
              table.Cell(i + 1, j + 1).Range.Text = String(data[i][j]);
            }
          }
        }
      }
      doc.Save();
      return { table_index: table.Index || 1 };
    } finally {
      doc.Close(0);
    }
  }

  function writerTableGet(path, index) {
    var doc = _openDoc(path, true);
    try {
      var table = doc.Tables(index || 1);
      var result = [];
      for (var i = 1; i <= table.Rows.Count; i++) {
        var row = [];
        for (var j = 1; j <= table.Columns.Count; j++) {
          row.push(String(table.Cell(i, j).Range.Text).replace(/\r?\n?$/, ""));
        }
        result.push(row);
      }
      return { data: result };
    } finally {
      doc.Close(0);
    }
  }

  function writerImageInsert(path, imagePath, width, height) {
    var doc = _openDoc(path, false);
    try {
      var sel = doc.ActiveWindow.Selection;
      var shape = sel.InlineShapes.AddPicture(toPath(imagePath));
      if (width) shape.Width = width;
      if (height) shape.Height = height;
      doc.Save();
      return { image: String(imagePath) };
    } finally {
      doc.Close(0);
    }
  }

  function writerPageSetup(path, width, height, mt, mb, ml, mr) {
    var doc = _openDoc(path, false);
    try {
      var page = doc.PageSetup;
      var mm2pt = 2.835;
      if (width !== undefined) page.PageWidth = width * mm2pt;
      if (height !== undefined) page.PageHeight = height * mm2pt;
      if (mt !== undefined) page.TopMargin = mt * mm2pt;
      if (mb !== undefined) page.BottomMargin = mb * mm2pt;
      if (ml !== undefined) page.LeftMargin = ml * mm2pt;
      if (mr !== undefined) page.RightMargin = mr * mm2pt;
      doc.Save();
      return {};
    } finally {
      doc.Close(0);
    }
  }

  function writerExportPdf(path, output) {
    var doc = _openDoc(path, true);
    try {
      doc.ExportAsFixedFormat(toPath(output), 17); // wdFormatPDF = 17
      return { path: String(output) };
    } finally {
      doc.Close(0);
    }
  }

  // ── Calc Actions ───────────────────────────────────────────

  function calcAction(action, p) {
    switch (action) {
      case "info":
        return calcInfo(p.path);
      case "cell_get":
        return calcCellGet(p.path, p.ref, p.sheet);
      case "cell_set":
        return calcCellSet(p.path, p.ref, p.value, p.sheet);
      case "cell_formula":
        return calcCellFormula(p.path, p.ref, p.formula, p.sheet);
      case "range_get":
        return calcRangeGet(p.path, p.ref, p.sheet);
      case "range_set":
        return calcRangeSet(p.path, p.ref, p.data, p.sheet);
      case "sheet_list":
        return calcSheetList();
      case "sheet_add":
        return calcSheetAdd(p.name);
      case "sheet_delete":
        return calcSheetDelete(p.name);
      case "sheet_rename":
        return calcSheetRename(p.old, p["new"]);
      case "chart_create":
        return calcChartCreate(p.path, p.data_range, p.chart_type, p.title);
      default:
        throw new Error("Unknown calc action: " + action);
    }
  }

  function _openWorkbook(path, readonly) {
    var fp = toPath(path);
    return Application.Workbooks.Open(fp, 0, readonly || false);
  }

  function _getSheet(wb, sheetName) {
    if (sheetName) return wb.Sheets(sheetName);
    return wb.ActiveSheet;
  }

  function calcInfo(path) {
    var wb = _openWorkbook(path, true);
    try {
      var sheetNames = [];
      for (var i = 1; i <= wb.Sheets.Count; i++) {
        sheetNames.push(String(wb.Sheets(i).Name));
      }
      var author = "";
      try {
        author = String(wb.BuiltInDocumentProperties("Author").Value);
      } catch (e) {}
      return {
        path: String(wb.FullName),
        sheets: wb.Sheets.Count,
        sheet_names: sheetNames,
        author: author,
      };
    } finally {
      wb.Close(0);
    }
  }

  function calcCellGet(path, ref, sheet) {
    var wb = _openWorkbook(path, true);
    try {
      var ws = _getSheet(wb, sheet);
      var val = ws.Range(ref).Value;
      return { value: val };
    } finally {
      wb.Close(0);
    }
  }

  function calcCellSet(path, ref, value, sheet) {
    var wb = _openWorkbook(path, false);
    try {
      var ws = _getSheet(wb, sheet);
      ws.Range(ref).Value = value;
      wb.Save();
      return {};
    } finally {
      wb.Close(0);
    }
  }

  function calcCellFormula(path, ref, formula, sheet) {
    var wb = _openWorkbook(path, false);
    try {
      var ws = _getSheet(wb, sheet);
      ws.Range(ref).Formula = formula;
      wb.Save();
      return {};
    } finally {
      wb.Close(0);
    }
  }

  function calcRangeGet(path, ref, sheet) {
    var wb = _openWorkbook(path, true);
    try {
      var ws = _getSheet(wb, sheet);
      var values = ws.Range(ref).Value;
      if (values === null || values === undefined) return { data: [] };
      if (!Array.isArray(values)) return { data: [[values]] };
      // May be a flat array or 2D
      if (values.length > 0 && !Array.isArray(values[0])) {
        return { data: [values] };
      }
      return { data: values };
    } finally {
      wb.Close(0);
    }
  }

  function calcRangeSet(path, ref, data, sheet) {
    var wb = _openWorkbook(path, false);
    try {
      var ws = _getSheet(wb, sheet);
      ws.Range(ref).Value = data;
      wb.Save();
      return {};
    } finally {
      wb.Close(0);
    }
  }

  function calcSheetList() {
    var wb = Application.ActiveWorkbook;
    if (!wb) throw new Error("No active workbook");
    var sheets = [];
    for (var i = 1; i <= wb.Sheets.Count; i++) {
      sheets.push({ index: i, name: String(wb.Sheets(i).Name) });
    }
    return { sheets: sheets };
  }

  function calcSheetAdd(name) {
    var wb = Application.ActiveWorkbook;
    if (!wb) throw new Error("No active workbook");
    var ws = wb.Sheets.Add(undefined, wb.Sheets(wb.Sheets.Count));
    if (name) ws.Name = name;
    wb.Save();
    return { name: String(ws.Name) };
  }

  function calcSheetDelete(name) {
    var wb = Application.ActiveWorkbook;
    if (!wb) throw new Error("No active workbook");
    wb.Sheets(name).Delete();
    wb.Save();
    return {};
  }

  function calcSheetRename(oldName, newName) {
    var wb = Application.ActiveWorkbook;
    if (!wb) throw new Error("No active workbook");
    wb.Sheets(oldName).Name = newName;
    wb.Save();
    return {};
  }

  function calcChartCreate(path, dataRange, chartType, title) {
    var wb = _openWorkbook(path, false);
    try {
      var ws = wb.ActiveSheet;
      var data = ws.Range(dataRange);
      var chartObj = ws.ChartObjects().Add(300, 50, 400, 250);
      var chart = chartObj.Chart;
      chart.SetSourceData(data);

      var typeMap = {
        bar: 51, // xlColumnClustered
        line: 4, // xlLine
        pie: 5, // xlPie
        scatter: -4169, // xlXYScatter
        area: 1, // xlArea
      };
      chart.ChartType = typeMap[chartType] || 51;
      if (title) {
        chart.HasTitle = true;
        chart.ChartTitle.Text = title;
      }
      wb.Save();
      return { index: chartObj.Index || 1 };
    } finally {
      wb.Close(0);
    }
  }

  // ── Impress Actions (stub) ─────────────────────────────────

  function impressAction(action, p) {
    throw new Error("Impress actions not yet implemented");
  }

  // ── Start ──────────────────────────────────────────────────

  connect();
})();
