function log(msg) {
    var logEl = document.getElementById("log");
    logEl.textContent += msg + "\n";
}

function limparLog() {
    document.getElementById("log").textContent = "";
}

function getText(parent, tagName) {
    if (!parent) return "";
    var els = parent.getElementsByTagName(tagName);
    if (!els.length) return "";
    return (els[0].textContent || "").trim();
}

function adicionarAoResumo(resumo, codProd, nomeProd, unidade, qtd) {
    var chave = codProd + "|" + unidade;
    if (resumo[chave]) {
        resumo[chave].quantidade += qtd;
    } else {
        resumo[chave] = {
            codProd: codProd,
            nomeProd: nomeProd,
            unidade: unidade,
            quantidade: qtd
        };
    }
}

function processarXml(xmlTexto, nomeArquivo, relatorio, resumo, totais) {
    var parser = new DOMParser();
    var doc = parser.parseFromString(xmlTexto, "application/xml");

    var inf = doc.getElementsByTagName("infNFe")[0];
    if (!inf) {
        log("infNFe não encontrado em " + nomeArquivo);
        return;
    }

    var ide = inf.getElementsByTagName("ide")[0];
    var numero = getText(ide, "nNF");

    var volumes = 0;
    var peso = 0;

    var transp = inf.getElementsByTagName("transp")[0];
    if (transp) {
        var vol = transp.getElementsByTagName("vol")[0];
        if (vol) {
            volumes = parseFloat(getText(vol, "qVol").replace(",", ".")) || 0;
            peso = parseFloat(getText(vol, "pesoB").replace(",", ".")) || 0;
        }
    }

    var dest = inf.getElementsByTagName("dest")[0];
    var nomeDest = dest ? getText(dest, "xNome") : "";
    var endereco = "";
    var cidade = "";

    if (dest) {
        var end = dest.getElementsByTagName("enderDest")[0];
        if (end) {
            endereco = getText(end, "xLgr") + ", " + getText(end, "nro");
            cidade = getText(end, "xMun");
        }
    }

    relatorio.push([numero, volumes, peso, nomeDest, endereco, cidade]);

    totais.totalNotas++;
    totais.totalVolumes += volumes;
    totais.totalPeso += peso;

    var dets = inf.getElementsByTagName("det");
    for (var i = 0; i < dets.length; i++) {
        var prod = dets[i].getElementsByTagName("prod")[0];
        if (!prod) continue;

        var cod = getText(prod, "cProd");
        var nomeProd = getText(prod, "xProd");
        var un = getText(prod, "uCom");
        var q = parseFloat(getText(prod, "qCom").replace(",", ".")) || 0;

        adicionarAoResumo(resumo, cod, nomeProd, un, q);
    }

    log("OK: " + nomeArquivo);
}

function gerarXlsx(relatorio, resumo, totais) {
    relatorio.push([]);
    relatorio.push(["TOTAL DE NOTAS:", totais.totalNotas]);
    relatorio.push(["TOTAL DE VOLUMES:", totais.totalVolumes]);
    relatorio.push(["TOTAL DE PESO:", totais.totalPeso]);

    var resumoArr = [["Código", "Produto", "Unidade", "Quantidade"]];

    for (var chave in resumo) {
        var item = resumo[chave];
        resumoArr.push([item.codProd, item.nomeProd, item.unidade, item.quantidade]);
    }

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(relatorio), "Relatorio");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumoArr), "ResumoProduto");

    XLSX.writeFile(wb, "Relatorio_NFe.xlsx");
}

window.onload = function () {
    var btn = document.getElementById("btnProcessar");
    var input = document.getElementById("xmlFiles");

    btn.onclick = function () {
        limparLog();
        log("Botão clicado.");

        var files = input.files;
        if (!files.length) {
            alert("Nenhum arquivo selecionado.");
            return;
        }

        log("Arquivos selecionados: " + files.length);

        var relatorio = [["Nota", "Volumes", "Peso", "Destinatário", "Endereço", "Cidade"]];
        var resumo = {};
        var totais = { totalNotas: 0, totalVolumes: 0, totalPeso: 0 };

        var contador = 0;

        for (let i = 0; i < files.length; i++) {
            let file = files[i];
            let reader = new FileReader();

            reader.onload = function (e) {
                processarXml(e.target.result, file.name, relatorio, resumo, totais);
                contador++;

                if (contador === files.length) {
                    log("Gerando XLSX...");
                    gerarXlsx(relatorio, resumo, totais);
                    log("Arquivo gerado!");
                    alert("Relatório gerado com sucesso!");
                }
            };

            reader.readAsText(file);
        }
    };
};
