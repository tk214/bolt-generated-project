import * as XLSX from 'xlsx';

document.getElementById('excelFile').addEventListener('change', handleFile);

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const queries = [];

        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            const cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: 0 });
            const cell = worksheet[cellAddress];
            if (cell && cell.v) {
                queries.push(cell.v);
            }
        }

        displayQueries(queries);
    };
    reader.readAsArrayBuffer(file);
}

function displayQueries(queries) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = '';

    if (queries.length === 0) {
        resultsDiv.innerHTML = '<p>No queries found in the Excel file.</p>';
        return;
    }

    for (let i = 0; i < queries.length; i += 31) {
        const group = queries.slice(i, i + 31);
        const groupDiv = document.createElement('div');
        groupDiv.classList.add('query-group');
        groupDiv.innerHTML = `<h3>Queries ${i + 1} - ${Math.min(i + 31, queries.length)}
        <button onclick="openAllQueries(${i})">Open All</button>
        </h3>`;

        group.forEach(query => {
            const queryDiv = document.createElement('div');
            queryDiv.classList.add('query-item');
            const encodedQuery = encodeURIComponent(query);
            queryDiv.innerHTML = `
                <span>${query}</span>
                <button onclick="window.open('https://www.google.com/search?q=${encodedQuery}', '_blank')">Search</button>
            `;
            groupDiv.appendChild(queryDiv);
        });
        resultsDiv.appendChild(groupDiv);
    }
}

window.openAllQueries = function(startIndex) {
    const resultsDiv = document.getElementById('results');
    const queries = [];
    resultsDiv.querySelectorAll('.query-item span').forEach(span => {
        queries.push(span.textContent);
    });
    const group = queries.slice(startIndex, startIndex + 31);
    group.forEach(query => {
        const encodedQuery = encodeURIComponent(query);
        window.open(`https://www.google.com/search?q=${encodedQuery}`, '_blank');
    });
}
