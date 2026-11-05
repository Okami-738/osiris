document.addEventListener("DOMContentLoaded", function() {
    const seletorArquivo = document.getElementById("arquivoExcel");

    if (seletorArquivo) {
        seletorArquivo.addEventListener("change", function(event) {
            const arquivo = event.target.files[0]; // Pega o primeiro arquivo selecionado
            if (!arquivo) return;

            const leitor = new FileReader();

            // Quando o arquivo for carregado, processamos os dados
            leitor.onload = function(e) {
                const dados = new Uint8Array(e.target.result);
                const workbook = XLSX.read(dados, { type: "array" });
                const planilha = workbook.Sheets["Ficha Pessoal"];

                if (!planilha) {
                    alert("A aba 'ficha pessoal' não foi encontrada!");
                    return;
                }
                //#region definimos o objeto com as coordenadas das células
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

                const pericias = {
                    B22: { totalCelula: "L22", nomeCelula: "B22" },
                    B23: { totalCelula: "L23", nomeCelula: "B23" },
                    B24: { totalCelula: "L24", nomeCelula: "B24" },
                    B25: { totalCelula: "L25", nomeCelula: "B25" },
                    B26: { totalCelula: "L26", nomeCelula: "B26" },
                    B27: { totalCelula: "L27", nomeCelula: "B27" },
                    B28: { totalCelula: "L28", nomeCelula: "B28" },
                    B29: { totalCelula: "L29", nomeCelula: "B29" },
                    B30: { totalCelula: "L30", nomeCelula: "B30" },
                    B31: { totalCelula: "L31", nomeCelula: "B31" },
                    B32: { totalCelula: "L32", nomeCelula: "B32" },
                    B33: { totalCelula: "L33", nomeCelula: "B33" },
                    B34: { totalCelula: "L34", nomeCelula: "B34" },
                    B35: { totalCelula: "L35", nomeCelula: "B35" },
                    C22: { totalCelula: "Z22", nomeCelula: "P22" },
                    C23: { totalCelula: "Z23", nomeCelula: "P23" },
                    C24: { totalCelula: "Z24", nomeCelula: "P24" },
                    C25: { totalCelula: "Z25", nomeCelula: "P25" },
                    C26: { totalCelula: "Z26", nomeCelula: "P26" },
                    C27: { totalCelula: "Z27", nomeCelula: "P27" },
                    C28: { totalCelula: "Z28", nomeCelula: "P28" },
                    C29: { totalCelula: "Z29", nomeCelula: "P29" },
                    C30: { totalCelula: "Z30", nomeCelula: "P30" },
                    C31: { totalCelula: "Z31", nomeCelula: "P31" },
                    C32: { totalCelula: "Z32", nomeCelula: "P32" },
                    C33: { totalCelula: "Z33", nomeCelula: "P33" },
                    C34: { totalCelula: "Z34", nomeCelula: "P34" },
                    C35: { totalCelula: "Z35", nomeCelula: "P35" }

                };

                const status = {
                    vida: { atualCelula: "AC7", maximaCelula: "AF7" },
                    energia: { atualCelula: "AC15", maximaCelula: "AF15" },
                    integridade: { atualCelula: "AM7", maximaCelula: "AP7" }
                };

                //#endregion

                //#region processamento atributo, nome e status
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

                    //console.log(`Atualizado: ${atributo.toUpperCase()} -> Valor: ${valorReal}, Mod: ${modificadorReal}`);
                }

               for (const nome in nomes) {
                    const celulaN = nomes[nome].valorCelula;//pega a célula do nome
                    const objNome = planilha[celulaN];//pega o objeto da célula do nome
                    const nomeReal = objNome ? objNome.v : "—";//se o "objNome" for nulo, undefined ou vazio, exibe "—"

                    const elementoNome = document.getElementById(`valor-${nome}`);//
                    if (elementoNome) {
                        elementoNome.innerText = nomeReal;
                    }
                }


                //#endregion
                //#region processamento das perícias
                // --- PROCESSAMENTO DAS PERÍCIAS ---
                const containerPericias = document.getElementById("container-pericias");
                if (containerPericias) {
                    // Limpa perícias anteriores ao carregar novo arquivo
                    containerPericias.innerHTML = "";

                    for (const chave in pericias) {
                        const celulaTotal = pericias[chave].totalCelula;
                        const celulaNome = pericias[chave].nomeCelula;

                        const objTotal = planilha[celulaTotal];
                        const objNome = planilha[celulaNome];

                        const valorTotal = objTotal ? objTotal.v : null;
                        const valorNomePericia = objNome ? objNome.v : null;

                        // Verifica se 'totalCelula' possui um valor válido (diferente de null, undefined ou vazio)
                        if (valorTotal !== null && valorTotal !== undefined && String(valorTotal).trim() !== "") {


                            // FORMATAÇÃO DO VALOR (Adiciona + se for positivo, mantém o - se for negativo ou exibe normal)
                              let valorFormatado = valorTotal;
                             const numero = Number(valorTotal);
            
                                // Se for um número válido, aplica a regra de sinais
                                if (!isNaN(numero)) {
                                    if (numero > 0) {
                                        valorFormatado = `+${numero}`;
                                    } else {
                                        valorFormatado = `${numero}`; // Se for negativo já vem com "-", se for 0 fica "0"
                                    }
                                }


                            // LIMITA O NOME DA PERÍCIA A 8 CARACTERES:
                             let nomeFormatado = valorNomePericia ? String(valorNomePericia) : "Perícia";
                            if (nomeFormatado.length > 8) {
                                  nomeFormatado = nomeFormatado.substring(0, 8) + "..."; // Corta em 8 e adiciona "..."
                            }

                            
                            // Cria a tag do elemento da perícia parecido com a estrutura de atributos
                            const artigoPericia = document.createElement("article");
                            artigoPericia.className = "pericia";

                            artigoPericia.innerHTML = `
                             <h2>${nomeFormatado}</h2>
                             <p class="pericia_valor"><span>${valorFormatado}</span></p>
                             `;

                            containerPericias.appendChild(artigoPericia);
                            //console.log(`Perícia Exibida: ${valorNomePericia} -> Total: ${valorTotal}`);
                        }
                    }
                }
                // #endregion


            };

            leitor.readAsArrayBuffer(arquivo);
        });
    }
});