// Objeto global para armazenar os valores
const dadosStatus = {
    vida: { atual: 0, max: 40 },
    energia: { atual: 0, max: 0 },
    integridade: { atual: 0, max: 0 }
};

// Função chamada pelos botões (onclick) no HTML
function alterarStatus(tipo, quantidade) {
    if (!dadosStatus[tipo]) return;

    // Atualiza o valor garantindo que fica entre 0 e o Máximo
    dadosStatus[tipo].atual = Math.min(
        dadosStatus[tipo].max,
        Math.max(0, dadosStatus[tipo].atual + quantidade)
    );

    atualizarInterfaceStatus(tipo);
}

// Função para atualizar os textos e a largura visual da barra
function atualizarInterfaceStatus(tipo) {
    const status = dadosStatus[tipo];
    const elemAtual = document.getElementById(`valor-${tipo}-atual`);
    const elemMax = document.getElementById(`valor-${tipo}-max`);
    const elemBarra = document.getElementById(`barra-${tipo}`);

    if (elemAtual) elemAtual.innerText = status.atual;
    if (elemMax) elemMax.innerText = status.max;

    if (elemBarra) {
        const porcentagem = status.max > 0 ? (status.atual / status.max) * 100 : 0;
        elemBarra.style.width = `${porcentagem}%`;
    }
}

