var btnConverter = document.querySelector('#btnConverter');
var inputFile = document.querySelector('#inputFile');
var lblInput = document.querySelector('#lblInput');
var file;
var extension;
var dados = [];

btnConverter.onclick = function () {
    if (file) {
        if (extension === 'xlsx' || extension === 'xls') {
            fileRead();
        } else {
            alert('Apenas arquivos com a extensão XLS ou XLSX!');
        }
    } else {
        alert('Selecione um arquivo!');
    }
};

inputFile.onchange = function (event) {
    file = event.target.files[0];
    extension = file.name.split('.')[1];
    lblInput.innerHTML = '';
    lblInput.appendChild(document.createTextNode(file.name));
};

function fileRead() {
    var fileReader = new FileReader();
    fileReader.onload = function (event) {
        var data = event.target.result;

        var workbook = XLSX.read(data, {
            type: "binary"
        });
        workbook.SheetNames.forEach(sheet => {
            let rowObject = XLSX.utils.sheet_to_row_object_array(
                workbook.Sheets[sheet]
            ); 
            rowObject.map(function (item) {
                if (item['Título da Pergunta'] == 'Resumo:') {
                    dados.push({
                        Nome: `${item.Nome} ${item.Sobrenome}`,
                        Disciplina: item['Bônus?'],
                        Nota: (item['__EMPTY'] / item['__EMPTY_1']).toLocaleString(undefined,{style: 'percent', minimumFractionDigits:2}),
                        Acertos: item['__EMPTY'],
                        Quantidade: item['__EMPTY_1']

                    });
                    console.log(item.Nome + '');
                }
            });
            dados.sort((a, b) => (a.Nome > b.Nome) ? 1 : ((b.Nome > a.Nome) ? -1 : 0));
            $('#table').htmlson({
                data: dados
            });
        });
    };
    fileReader.readAsBinaryString(file);
}