document.addEventListener("DOMContentLoaded", function() {
    const seletorArquivo = document.getElementById("arquivoExcel");

    if (seletorArquivo) {
        seletorArquivo.addEventListener("change", function(event) {
            const arquivo = event.target.files[0]; // Pega o primeiro arquivo selecionado
            if (!arquivo) return;

            const leitor = new FileReader();

            leitor.onload = function(e) {
                const dados = new Uint8Array(e.target.result);
                const workbook = XLSX.read(dados, { type: "array" });
                const planilha = workbook.Sheets["Ficha Pessoal"];

                if (!planilha) {
                    alert("A aba 'ficha pessoal' não foi encontrada!");
                    return;
                }

                // 1. Aqui definimos o objeto com as coordenadas das células
                const nomes = {
                    nome: { valorCelula: "F3" }
                };

                const atributos = {
                    forca: { valorCelula: "G12", modCelula: "K12" },
                    destreza: { valorCelula: "G14", modCelula: "K14" },
                    constituicao: { valorCelula: "G16", modCelula: "K16" }, 
                    inteligencia: { valorCelula: "T12", modCelula: "X12" },
                    sabedoria: { valorCelula: "T14", modCelula: "X14" },
                    carisma: { valorCelula: "T16", modCelula: "X16" },

                };



                console.log("--- ATUALIZANDO HTML DO PERSONAGEM ---");

                // 2. Este é o loop onde a "const atributo" nasce e roda para cada linha do objeto acima
                for (const atributo in atributos) {
                    const celulaV = atributos[atributo].valorCelula;
                    const celulaM = atributos[atributo].modCelula;

                    const objValor = planilha[celulaV];
                    const objModificador = planilha[celulaM];

                    const valorReal = objValor ? objValor.v : "—";
                    const modificadorReal = objModificador ? objModificador.v : "—";

                    // 3. Aqui o JS busca as tags no seu HTML usando os IDs dinâmicos
                    const elementoValor = document.getElementById(`valor-${atributo}`);
                    const elementoMod = document.getElementById(`mod-${atributo}`);

                    // 4. Se o JS encontrar os IDs no HTML, ele injeta os valores do Excel
                    if (elementoValor) elementoValor.innerText = valorReal;
                    if (elementoMod) elementoMod.innerText = modificadorReal;

                    console.log(`Atualizado: ${atributo.toUpperCase()} -> Valor: ${valorReal}, Mod: ${modificadorReal}`);
                }

               for (const nome in nomes) {
                    const celulaN = nomes[nome].valorCelula;
                    const objNome = planilha[celulaN];
                    const nomeReal = objNome ? objNome.v : "—";

                    const elementoNome = document.getElementById(`valor-${nome}`);
                    if (elementoNome) {
                        elementoNome.innerText = nomeReal;
                    }
                }

            };

            leitor.readAsArrayBuffer(arquivo);
        });
    }
});