document.getElementById('inputExcel').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const tbody = document.querySelector('#tabelaExcel tbody');
    tbody.innerHTML = ''; // Limpa tabela

    // pula o cabeçalho (linha 0)
    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i];
      const tr = document.createElement('tr');

      for (let j = 0; j < 6; j++) {
        const td = document.createElement('td');
        td.textContent = row[j] || '';
        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }
  };

  reader.readAsArrayBuffer(file);
}
