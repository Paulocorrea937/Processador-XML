const FOLDER_ID = "1geD-SwS98xzSWnxYcF3zCTrlz21ZG5WL"; // pasta dos XML no Drive

function doGet(e) {
  var folder;
  try {
    folder = DriveApp.getFolderById(FOLDER_ID);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      "<html><body style='font-family:sans-serif;background:#020617;color:#e5e7eb;padding:20px;'>" +
      "<h2>Erro</h2><p>ID da pasta inválido ou sem permissão.</p></body></html>"
    );
  }

  var it = folder.getFiles();
  var files = [];

  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    if (name && name.toLowerCase().endsWith(".xml")) {
      files.push({
        id: f.getId(),
        name: name,
        date: f.getLastUpdated()
      });
    }
  }

  // mais novos primeiro
  files.sort(function (a, b) {
    return b.date - a.date;
  });

  var html = [];
  html.push("<html><head><meta charset='UTF-8'><title>Selecionar XML</title></head>");
  html.push("<body style='background:#020617;color:#e5e7eb;font-family:system-ui,Segoe UI,sans-serif;padding:20px;'>");
  html.push("<h2 style='margin-top:0'>Selecione os XML que deseja baixar</h2>");
  html.push("<p style='font-size:13px;color:#9ca3af'>Marque os arquivos desejados e clique em <b>Baixar arquivos selecionados (ZIP)</b>.</p>");

  if (files.length === 0) {
    html.push("<p>Nenhum arquivo XML encontrado na pasta configurada.</p>");
  } else {
    html.push("<form method='post'>");
    html.push("<table style='width:100%;max-width:980px;border-collapse:collapse;font-size:13px;'>");
    html.push("<thead><tr>");
    html.push("<th style='text-align:left;padding:6px;border-bottom:1px solid #334155;'>Sel</th>");
    html.push("<th style='text-align:left;padding:6px;border-bottom:1px solid #334155;'>Nome do arquivo</th>");
    html.push("<th style='text-align:left;padding:6px;border-bottom:1px solid #334155;'>Última modificação</th>");
    html.push("</tr></thead><tbody>");

    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      var dataStr = Utilities.formatDate(
        f.date,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      );
      html.push("<tr>");
      html.push("<td style='padding:4px 6px;border-bottom:1px solid #1f2937;'>");
      html.push("<input type='checkbox' name='fileId' value='" + f.id + "'>");
      html.push("</td>");
      html.push("<td style='padding:4px 6px;border-bottom:1px solid #1f2937;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:420px;'>" +
                f.name + "</td>");
      html.push("<td style='padding:4px 6px;border-bottom:1px solid #1f2937;'>" + dataStr + "</td>");
      html.push("</tr>");
    }

    html.push("</tbody></table>");
    html.push("<div style='margin-top:16px;'>");
    html.push("<button type='submit' " +
      "style='padding:8px 16px;border-radius:999px;border:none;cursor:pointer;" +
      "background:linear-gradient(90deg,#0ea5e9,#22c55e);color:#020617;font-weight:600;" +
      "box-shadow:0 12px 30px rgba(34,197,94,0.35);'>" +
      "Baixar arquivos selecionados (ZIP)" +
      "</button>");
    html.push("</div>");
    html.push("</form>");
  }

  html.push("<p style='margin-top:18px;font-size:12px;color:#9ca3af;'>");
  html.push("Arquivos que chegam depois da meia-noite aparecem aqui com a data correta. ");
  html.push("É só marcar os que você quiser, independente do dia.");
  html.push("</p>");

  html.push("</body></html>");

  return HtmlService
    .createHtmlOutput(html.join(""))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var ids = e.parameter.fileId;

  if (!ids) {
    return HtmlService.createHtmlOutput(
      "<html><body style='font-family:sans-serif;background:#020617;color:#e5e7eb;padding:20px;'>" +
      "<h2>Nenhum arquivo selecionado</h2>" +
      "<p>Volte e marque pelo menos um XML.</p></body></html>"
    );
  }

  var list = Array.isArray(ids) ? ids : [ids];
  var blobs = [];

  for (var i = 0; i < list.length; i++) {
    var file = DriveApp.getFileById(list[i]);
    var nome = file.getName() || ("arquivo_" + (i + 1) + ".xml");
    // garante nomes únicos dentro do ZIP
    blobs.push(file.getBlob().setName((i + 1) + "_" + nome));
  }

  var zip = Utilities.zip(blobs, "xml_selecionados.zip");

  return ContentService
    .createBlob(zip.getBytes())
    .setName("xml_selecionados.zip")
    .setMimeType("application/zip");
}
