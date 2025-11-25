function log(msg) {
    const logEl = document.getElementById("log");
    logEl.textContent += msg + "\n";
}

function limparLog() {
    document.getElementById("log").textContent = "";
}

function getText(parent, tag) {
    if (!parent) return "";
    const el = parent.getElementsByTagName(tag)[0];
    return el ? el.textContent.trim() : "";
}

function adicionarAoResumo(resumo, cod, nome, un, qtd) {
    const key = cod + "|" + un;
    if (!resumo[key]) resumo[key] = { cod, nome, un, qtd: 0 };
    resumo[key].qtd += qtd;
}

function processarXml(xml, nome, relatorio, resumo, tot) {
    try {
        const doc = new DOMParser().parseFromString(xml, "text/xml");
        const inf = doc.getElementsByTagName("infNFe")[0];
        if (!inf) return log("Erro: infNFe não encontrado em " + nome);

        const ide = inf.getElementsByTagName("ide")[0];
        const num = getText(ide, "nNF");

        const dest = inf.getElementsByTagName("dest")[0];
        const nomeDest = dest ? getText(dest, "xNome") : "";

        const end = dest ? dest.getElementsByTagName("enderDest")[0] : null;
        const endereco = end ? `${getText(end, "xLgr")}, ${getText(end, "nro")}` : "";
        const cidade = end ? getText(end, "xMun") : "";

        let volumes = 0, peso = 0;
        const transp = inf.getElementsByTagName("transp")[0];
        if (transp) {
            const vol = transp.getElementsByTagName("vol")[0];
            if (vol) {
                volumes = parseFloat(getText(vol, "qVol").replace(",", ".")) || 0;
                peso = parseFloat(getText(vol, "pesoB").replace(",", ".")) || 0;
            }
        }

        relatorio.push([num, volumes, peso, nomeDest, endereco, cidade]);

        tot.notas++;
        tot.volumes += volumes;
        tot.peso += peso;

        const dets = inf.getElementsByTagName("det");
        for (let d of dets) {
            const prod = d.getElementsByTagName("prod")[0];
            if (!prod) continue;
            adicionarAoResumo(
                resumo,
                getText(prod, "cProd"),
                getText(prod, "xProd"),
                getText(prod, "uCom"),
                parseFloat(getText(prod, "qCom").replace(",", ".")) || 0
            );
        }

        log("OK: " + nome);
    } catch (e) {
        log("Erro processando " + nome + ": " + e);
    }
}

function gerar(relatorio, resumo, tot) {
    relatorio.push([]);
    relatorio.push(["TOTAL NOTAS", tot.notas]);
    relatorio.push(["TOTAL VOLUMES", tot.volumes]);
    relatorio.push(["TOTAL PESO", tot.peso]);

    const resumoArr = [["Código", "Produto", "Un", "Quantidade"]];
    for (let k in resumo) resumoArr.push(Object.values(resumo[k]));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(relatorio), "Relatório");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumoArr), "Produtos");

    XLSX.writeFile(wb, "Relatorio_NFe.xlsx");
}

async function lerZIP(file, callback) {
    const JSZip = await import("https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js");
    log("Lendo ZIP...");
    const zip = await JSZip.loadAsync(file);

    for (let filename in zip.files) {
        if (filename.endsWith(".xml")) {
            const xml = await zip.files[filename].async("text");
            callback(xml, filename);
        }
    }
}

window.onload = () => {
    const btn = document.getElementById("btnProcessar");
    const input = document.getElementById("xmlFiles");

    btn.onclick = async () => {
        limparLog();
        if (!input.files.length) return alert("Selecione arquivos XML ou ZIP!");

        log("Processando arquivos...");

        const relatorio = [["Nota", "Volumes", "Peso", "Destinatário", "Endereço", "Cidade"]];
        const resumo = {};
        const tot = { notas: 0, volumes: 0, peso: 0 };

        let count = 0, total = input.files.length;

        for (let file of input.files) {
            if (file.name.endsWith(".zip")) {
                await lerZIP(file, (xml, nome) => {
                    processarXml(xml, nome, relatorio, resumo, tot);
                });
                count++;
                continue;
            }

            const txt = await file.text();
            processarXml(txt, file.name, relatorio, resumo, tot);
            count++;
        }

        if (count === total) {
            log("Gerando XLSX...");
            gerar(relatorio, resumo, tot);
        }
    };

    document.getElementById("btnDrive").onclick = () => {
        const url = "https://drive.google.com/drive/folders/1geD-SwS98xzSWnxYcF3zCTrlz21ZG5WL";
        log("Abrindo pasta 'Coletor XML' no Drive...");
        window.open(url, "_blank");
    };
};
