$(document).ready(function() {
    let table = $('#tabelaDados').DataTable({
        language: { url: '//cdn.datatables.net/plug-ins/1.13.6/i18n/pt-BR.json' },
        pageLength: 15
    });

    let dadosParaExportar = []; // Variável global para armazenar os dados limpos

    document.getElementById('fileInput').addEventListener('change', function(e) {
        let file = e.target.files[0];
        if (!file) return;

        let reader = new FileReader();
        reader.onload = function(event) {
            try {
                let content = event.target.result;
                let jsonStartIndex = content.indexOf('{');
                let jsonOnly = content.substring(jsonStartIndex);
                let parsedData = JSON.parse(jsonOnly);

                if (parsedData.data) {
                    processarDados(parsedData.data);
                    document.getElementById('btnExportar').disabled = false; // Habilita o botão
                }
            } catch (err) {
                alert("Erro no formato do arquivo.");
            }
        };
        reader.readAsText(file);
    });

    function processarDados(beneficios) {
        table.clear();
        dadosParaExportar = []; // Reseta a lista de exportação

        beneficios.forEach(item => {
            const nomePrograma = item.benefit.name;
            
            item.benefit.products.forEach(prod => {
                // Adiciona na tabela visual
                table.row.add([
                    nomePrograma,
                    `<strong>${prod.name}</strong>`,
                    prod.presentation,
                    `<span class="ean-badge">${prod.ean}</span>`,
                    `R$ ${prod.salePrice.toLocaleString('pt-BR', {minimumFractionDigits: 2})}`,
                    `${prod.discountPercent}%`,
                    prod.commercialGrade
                ]);

                // Prepara o objeto para a planilha Excel (Dados Limpos)
                dadosParaExportar.push({
                    "Programa": nomePrograma,
                    "Produto": prod.name,
                    "Apresentação": prod.presentation,
                    "EAN": prod.ean,
                    "Preço Venda": prod.salePrice,
                    "Desconto %": prod.discountPercent,
                    "Categoria": prod.commercialGrade
                });
            });
        });
        table.draw();
    }

    // FUNÇÃO DE EXPORTAÇÃO
    document.getElementById('btnExportar').addEventListener('click', function() {
        if (dadosParaExportar.length === 0) return;

        // 1. Cria uma nova "folha" de trabalho
        const worksheet = XLSX.utils.json_to_sheet(dadosParaExportar);
        
        // 2. Cria um novo "livro" de Excel
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Carga Epharma");

        // 3. Gera o arquivo e inicia o download
        const dataAtual = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(workbook, `Carga_Epharma_${dataAtual}.xlsx`);
    });
});