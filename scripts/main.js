var btnConverter = document.querySelector('#btnConverter');
var inputFile = document.querySelector('#inputFile');
var lblInput = document.querySelector('#lblInput');
var drpDisciplinas = document.querySelector('#drpDisciplinas');
var table = document.querySelector('#table');
var file;
var extension;
var dados = [];
var filtro = []

drpDisciplinas.onchange = function (event) {
    var opcao = drpDisciplinas.options[drpDisciplinas.selectedIndex].value;
    if(opcao == 'todos'){
        renderTable(dados);
    }else{
        filtro = dados.filter(item => item.Disciplina == opcao);
        renderTable(filtro);
    }
};

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
                        Nota: (item['__EMPTY'] / item['__EMPTY_1']).toLocaleString(undefined, { style: 'percent', minimumFractionDigits: 2 }),
                        Acertos: item['__EMPTY'],
                        Quantidade: item['__EMPTY_1']

                    });
                }
            });
            renderTable(dados);
            createDropDown();
        });
    };
    fileReader.readAsBinaryString(file);
}

function createDropDown() {
    var disciplinas = dados.map(item => item.Disciplina)
        .filter(function (value, index, self) {
            return self.indexOf(value) === index;
        });
    disciplinas.forEach(function (value, item, array) {
        var disciplina = document.createElement('option');
        disciplina.appendChild(document.createTextNode(value));
        disciplina.value = value;
        drpDisciplinas.appendChild(disciplina);
    });
};

function renderTable(filtro) {
    table.innerHTML = '';
    dados.sort((a, b) => (a.Nome > b.Nome) ? 1 : ((b.Nome > a.Nome) ? -1 : 0));
    $('#table').htmlson({
        data: filtro
    });
}