function log(msg) {
    var logEl = document.getElementById("log");
    if (!logEl) return;
    logEl.textContent += msg + "\n";
}

function limparLog() {
    var logEl = document.getElementById("log");
    if (logEl) logEl.textContent = "";
}

function getText(parent, tagName) {
    if (!parent) return "";
    var els = parent.getElementsByTagName(tagName);
    if (!els || !els.length) return "";
    var t = els[0].textContent || "";
    return t.replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
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
            var qVolStr = getText(vol, "qVol").replace(",", ".");
            var pesoStr = getText(vol, "pesoB").replace(",", ".");
            volumes = qVolStr ? parseFloat(qVolStr) : 0;
            peso = pesoStr ? parseFloat(pesoStr) : 0;
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

    totais.totalNotas += 1;
    totais.totalVolumes += volumes;
    totais.totalPeso += peso;

    var dets = inf.getElementsByTagName("det");
    for (var i = 0; i < dets.length; i++) {
        var prod = dets[i].getElementsByTagName("prod")[0];
        if (!prod) continue;

        var cod = getText(prod, "cProd");
        var nomeProd = getText(prod, "xProd");
        var un = getText(prod, "uCom");
        var qStr = getText(prod, "qCom").replace(",", ".");
        var q = qStr ? parseFloat(qStr) : 0;

        adicionarAoResumo(resumo, cod, nomeProd, un, q);
    }

    log("OK: " + nomeArquivo);
}

function gerarXlsx(relatorio, resumo, totais) {
    // linhas de total
    relatorio.push([]);
    relatorio.push(["TOTAL DE NOTAS:", totais.totalNotas]);
    relatorio.push(["TOTAL DE VOLUMES:", totais.totalVolumes]);
    relatorio.push(["TOTAL DE PESO:", totais.totalPeso]);

    // resumo
    var resumoArr = [["Código", "Produto", "Unidade", "Quantidade"]];
    for (var chave in resumo) {
        if (!resumo.hasOwnProperty(chave)) continue;
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

    if (!btn || !input) {
        console.error("Elementos não encontrados no DOM.");
        return;
    }

    // Botão que processa arquivos locais
    btn.onclick = function () {
        limparLog();
        log("Botão clicado.");

        if (typeof XLSX === "undefined") {
            log("Erro: XLSX não carregou.");
            alert("Erro: biblioteca XLSX não carregou. Verifique a conexão e recarregue a página.");
            return;
        }

        var files = input.files;
        if (!files || !files.length) {
            log("Nenhum XML selecionado.");
            alert("Selecione pelo menos um arquivo XML.");
            return;
        }

        log("Arquivos selecionados: " + files.length);

        var relatorio = [["Nota", "Volumes", "Peso", "Destinatário", "Endereço", "Cidade"]];
        var resumo = {};
        var totais = { totalNotas: 0, totalVolumes: 0, totalPeso: 0 };

        var totalArquivos = files.length;
        var processados = 0;

        for (var i = 0; i < files.length; i++) {
            (function (file) {
                var reader = new FileReader();

                reader.onload = function (e) {
                    try {
                        processarXml(e.target.result, file.name, relatorio, resumo, totais);
                    } catch (erro) {
                        log("Erro ao processar " + file.name + ": " + erro);
                    }
                    processados++;

                    if (processados === totalArquivos) {
                        log("Todos os arquivos processados. Gerando XLSX...");
                        try {
                            gerarXlsx(relatorio, resumo, totais);
                            log("Arquivo XLSX gerado com sucesso.");
                            alert("Relatório gerado! Verifique o arquivo Relatorio_NFe.xlsx na pasta de downloads.");
                        } catch (erro2) {
                            log("Erro ao gerar XLSX: " + erro2);
                            alert("Erro ao gerar XLSX. Veja o log na tela.");
                        }
                    }
                };

                reader.onerror = function () {
                    log("Erro de leitura no arquivo " + file.name);
                    processados++;
                };

                reader.readAsText(file);
            })(files[i]);
        }
    };

    // Botão que abre o painel para baixar XML do Google Drive
    var btnDrive = document.getElementById("btnDrive");
    if (btnDrive) {
        btnDrive.onclick = function () {
            var webAppUrl = "https://script.google.com/macros/s/AKfycbzC31h7Hf0nt8sjeetmhOkVm1IE_35FbnNPp1IFpIgliabWiOttInkHNTZtzmw-N6Fm/exec";
            window.open(webAppUrl, "_blank");
        };
    } else {
        console.warn("Botão btnDrive não encontrado no HTML.");
    }
};
