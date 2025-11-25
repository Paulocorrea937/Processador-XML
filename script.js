// ------------------ ESTADO GLOBAL ------------------
const state = {
    relatorio: [],
    resumoProdutos: {},          // chave: cod|un
    totais: { notas: 0, volumes: 0, peso: 0 },
    cidades: {},                 // cidade -> { notas, volumes }
    clientes: new Set(),
    produtosUnicos: new Set(),
    processado: false
};

// ------------------ HELPER LOG ------------------
function log(msg) {
    const logEl = document.getElementById("log");
    if (!logEl) return;
    if (logEl.textContent.trim() === "Aguardando processamento...") {
        logEl.textContent = "";
    }
    logEl.textContent += msg + "\n";
    logEl.scrollTop = logEl.scrollHeight;
}

function limparLog() {
    const logEl = document.getElementById("log");
    if (logEl) logEl.textContent = "▶ pronto para processar arquivos...\n";
}

// ------------------ XML HELPERS ------------------
function getText(parent, tag) {
    if (!parent) return "";
    const el = parent.getElementsByTagName(tag)[0];
    return el ? (el.textContent || "").trim() : "";
}

function adicionarAoResumoProduto(cod, nome, un, qtd) {
    const chave = cod + "|" + un;
    if (!state.resumoProdutos[chave]) {
        state.resumoProdutos[chave] = { cod, nome, un, qtd: 0 };
    }
    state.resumoProdutos[chave].qtd += qtd;
    state.produtosUnicos.add(chave);
}

// ------------------ PROCESSAMENTO XML ------------------
function processarXml(xmlTexto, nomeArquivo) {
    try {
        const doc = new DOMParser().parseFromString(xmlTexto, "application/xml");

        const inf = doc.getElementsByTagName("infNFe")[0];
        if (!inf) {
            log("ERRO: infNFe não encontrado em " + nomeArquivo);
            return;
        }

        const ide = inf.getElementsByTagName("ide")[0];
        const numero = getText(ide, "nNF") || "(sem número)";

        let volumes = 0;
        let peso = 0;
        const transp = inf.getElementsByTagName("transp")[0];
        if (transp) {
            const vol = transp.getElementsByTagName("vol")[0];
            if (vol) {
                const qVolStr = getText(vol, "qVol").replace(",", ".");
                const pesoStr = getText(vol, "pesoB").replace(",", ".");
                volumes = qVolStr ? parseFloat(qVolStr) : 0;
                peso = pesoStr ? parseFloat(pesoStr) : 0;
            }
        }

        const dest = inf.getElementsByTagName("dest")[0];
        const nomeDest = dest ? getText(dest, "xNome") : "";
        let endereco = "";
        let cidade = "";

        if (dest) {
            const end = dest.getElementsByTagName("enderDest")[0];
            if (end) {
                endereco = `${getText(end, "xLgr")}, ${getText(end, "nro")}`.trim();
                cidade = getText(end, "xMun") || "";
            }
        }

        // Linha no relatório
        state.relatorio.push([numero, volumes, peso, nomeDest, endereco, cidade]);

        // Totais
        state.totais.notas += 1;
        state.totais.volumes += volumes;
        state.totais.peso += peso;

        if (nomeDest) state.clientes.add(nomeDest);
        if (cidade) {
            if (!state.cidades[cidade]) state.cidades[cidade] = { notas: 0, volumes: 0 };
            state.cidades[cidade].notas += 1;
            state.cidades[cidade].volumes += volumes;
        }

        // Produtos
        const dets = inf.getElementsByTagName("det");
        for (let i = 0; i < dets.length; i++) {
            const prod = dets[i].getElementsByTagName("prod")[0];
            if (!prod) continue;
            const cod = getText(prod, "cProd");
            const nomeProd = getText(prod, "xProd");
            const un = getText(prod, "uCom");
            const qStr = getText(prod, "qCom").replace(",", ".");
            const q = qStr ? parseFloat(qStr) : 0;

            adicionarAoResumoProduto(cod, nomeProd, un, q);
        }

        log("OK: " + nomeArquivo);
    } catch (err) {
        log("ERRO ao processar " + nomeArquivo + ": " + err);
    }
}

// ------------------ GERAÇÃO XLSX ------------------
function gerarXlsx() {
    const relatorio = [["Nota", "Volumes", "Peso", "Destinatário", "Endereço", "Cidade"], ...state.relatorio];

    // Totais
    relatorio.push([]);
    relatorio.push(["TOTAL DE NOTAS:", state.totais.notas]);
    relatorio.push(["TOTAL DE VOLUMES:", state.totais.volumes]);
    relatorio.push(["TOTAL DE PESO:", state.totais.peso]);

    const resumoArr = [["Código", "Produto", "Unidade", "Quantidade"]];
    for (const chave in state.resumoProdutos) {
        const item = state.resumoProdutos[chave];
        resumoArr.push([item.cod, item.nome, item.un, item.qtd]);
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(relatorio), "Relatorio");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumoArr), "ResumoProdutos");

    XLSX.writeFile(wb, "Relatorio_NFe.xlsx");
}

// ------------------ DASHBOARD ------------------
function atualizarDashboard() {
    const info = document.getElementById("dashInfo");
    if (!state.processado) {
        info.textContent = "Importe arquivos na aba Processador para atualizar os dados.";
    } else {
        info.textContent = "Dados atualizados a partir do último processamento.";
    }

    document.getElementById("dashNotas").textContent = state.totais.notas;
    document.getElementById("dashVolumes").textContent = state.totais.volumes.toFixed(2);
    document.getElementById("dashPeso").textContent = state.totais.peso.toFixed(2);
    document.getElementById("dashClientes").textContent = state.clientes.size;
    document.getElementById("dashCidades").textContent = Object.keys(state.cidades).length;
    document.getElementById("dashProdutos").textContent = state.produtosUnicos.size;

    // Top cidades
    const cidadesArray = Object.entries(state.cidades).map(([nome, obj]) => ({
        cidade: nome,
        notas: obj.notas,
        volumes: obj.volumes
    }));
    cidadesArray.sort((a, b) => b.volumes - a.volumes);
    const tbodyCid = document.querySelector("#tblCidades tbody");
    tbodyCid.innerHTML = "";
    cidadesArray.slice(0, 6).forEach(c => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${c.cidade}</td><td>${c.notas}</td><td>${c.volumes.toFixed(2)}</td>`;
        tbodyCid.appendChild(tr);
    });
    if (tbodyCid.children.length === 0) {
        tbodyCid.innerHTML = `<tr><td colspan="3" class="muted">Sem dados ainda.</td></tr>`;
    }

    // Top produtos
    const produtosArray = Object.values(state.resumoProdutos);
    produtosArray.sort((a, b) => b.qtd - a.qtd);
    const tbodyProd = document.querySelector("#tblProdutos tbody");
    tbodyProd.innerHTML = "";
    produtosArray.slice(0, 6).forEach(p => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${p.nome}</td><td>${p.un}</td><td>${p.qtd.toFixed(2)}</td>`;
        tbodyProd.appendChild(tr);
    });
    if (tbodyProd.children.length === 0) {
        tbodyProd.innerHTML = `<tr><td colspan="3" class="muted">Sem dados ainda.</td></tr>`;
    }
}

// ------------------ PROCESSAMENTO DE ARQUIVOS ------------------
async function processarArquivos(files) {
    if (!files || !files.length) {
        alert("Selecione pelo menos um arquivo XML ou ZIP.");
        return;
    }

    if (typeof XLSX === "undefined") {
        alert("Biblioteca XLSX não carregou. Verifique a conexão e recarregue a página.");
        return;
    }

    // reset estado
    state.relatorio = [];
    state.resumoProdutos = {};
    state.totais = { notas: 0, volumes: 0, peso: 0 };
    state.cidades = {};
    state.clientes = new Set();
    state.produtosUnicos = new Set();
    state.processado = false;
    limparLog();
    log("Iniciando processamento...");
    log("Arquivos selecionados: " + files.length);

    for (const file of files) {
        const lower = file.name.toLowerCase();
        if (lower.endsWith(".zip")) {
            log("Lendo ZIP: " + file.name);
            try {
                const zip = await JSZip.loadAsync(file);
                const entries = Object.keys(zip.files).filter(name => name.toLowerCase().endsWith(".xml"));
                log("  XML encontrados no ZIP: " + entries.length);
                for (const entryName of entries) {
                    const xml = await zip.files[entryName].async("text");
                    processarXml(xml, entryName);
                }
            } catch (err) {
                log("ERRO ao ler ZIP " + file.name + ": " + err);
            }
        } else if (lower.endsWith(".xml")) {
            const text = await file.text();
            processarXml(text, file.name);
        } else {
            log("Ignorando arquivo (não é XML nem ZIP): " + file.name);
        }
    }

    state.processado = state.totais.notas > 0;
    atualizarDashboard();

    if (!state.processado) {
        log("Nenhuma NF-e válida encontrada. Verifique os arquivos.");
        alert("Nenhuma NF-e válida encontrada. Confira os arquivos selecionados.");
        return;
    }

    log("Todos os arquivos processados. Gerando XLSX...");
    try {
        gerarXlsx();
        log("Arquivo Relatorio_NFe.xlsx gerado com sucesso.");
        alert("Relatório gerado! Verifique o arquivo Relatorio_NFe.xlsx na pasta de downloads.");
    } catch (err) {
        log("ERRO ao gerar XLSX: " + err);
        alert("Erro ao gerar XLSX. Veja o log na tela.");
    }
}

// ------------------ UI: TEMA & TABS ------------------
function initTema() {
    const root = document.documentElement;
    const toggleBtn = document.getElementById("themeToggle");
    const modeLabel = document.getElementById("modeLabel");

    function applyTheme(theme) {
        root.setAttribute("data-theme", theme);
        modeLabel.textContent = theme === "light" ? "Modo claro" : "Modo escuro";
    }

    let saved = localStorage.getItem("mblog-theme");
    if (!saved) saved = "dark";
    applyTheme(saved);

    toggleBtn.addEventListener("click", () => {
        const current = root.getAttribute("data-theme") === "light" ? "light" : "dark";
        const next = current === "light" ? "dark" : "light";
        applyTheme(next);
        localStorage.setItem("mblog-theme", next);
    });
}

function initTabs() {
    const tabs = document.querySelectorAll(".nav-tab");
    tabs.forEach(tab => {
        tab.addEventListener("click", () => {
            tabs.forEach(t => t.classList.remove("active"));
            tab.classList.add("active");

            const sectionId = tab.dataset.section;
            document.querySelectorAll(".section").forEach(sec => {
                sec.classList.remove("active");
            });
            document.getElementById("section-" + sectionId).classList.add("active");
        });
    });
}

// ------------------ INICIALIZAÇÃO ------------------
window.addEventListener("load", () => {
    initTema();
    initTabs();
    limparLog();
    atualizarDashboard();

    const btnProcessar = document.getElementById("btnProcessar");
    const inputArquivos = document.getElementById("xmlFiles");
    const btnDrive = document.getElementById("btnDrive");

    btnProcessar.addEventListener("click", () => {
        processarArquivos(inputArquivos.files);
    });

    btnDrive.addEventListener("click", () => {
        const driveFolderUrl = "https://drive.google.com/drive/folders/1geD-SwS98xzSWnxYcF3zCTrlz21ZG5WL";
        log("Abrindo pasta 'Coletor XML' no Google Drive...");
        window.open(driveFolderUrl, "_blank");
    });
});
