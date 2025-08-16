document.addEventListener('DOMContentLoaded', () => {
  // Elementos de la UI
  const fileUpload   = document.getElementById('file-upload');
  const micButton    = document.getElementById('mic-button');
  const resetButton  = document.getElementById('reset-button');
  const exportButton = document.getElementById('export-button');
  const inventoryList= document.getElementById('inventory-list');
  const status       = document.getElementById('status');
  const textCommand  = document.getElementById('text-command');
  const textButton   = document.getElementById('text-button');

  // Estado
  let inventoryData = [];      // [{ Producto, UMB, Stock, ...original }]
  let fuse;                    // Fuse.js index
  let originalWorkbook;        // XLSX workbook
  let headers = [];            // Cabeceras de la hoja
  let stockColumnIndex = -1;   // Índice de columna 'Stock'

  // ========= CARGA DE EXCEL =========
  fileUpload.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      originalWorkbook = XLSX.read(data, { type: 'array', cellStyles: true });

      const sheetName = originalWorkbook.SheetNames[0];
      const ws = originalWorkbook.Sheets[sheetName];

      // Obtenemos cabeceras
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1 });
      headers = aoa[0] || [];

      // Localiza columna 'Stock' (si no existe, la crearemos a la exportación)
      stockColumnIndex = headers.indexOf('Stock');

      // Datos como objetos (usa la 1ª fila como header)
      const jsonData = XLSX.utils.sheet_to_json(ws);

      // Normaliza: arrancamos Stock a 0 para el recuento
      inventoryData = jsonData.map(item => ({
        ...item,
        Stock: 0,
      }));

      // Índice Fuse por nombre de producto
      fuse = new Fuse(inventoryData, {
        keys: ['Producto'],
        includeScore: true,
        threshold: 0.4,
        ignoreLocation: true,
      });

      // Habilita controles
      micButton.disabled  = false;
      resetButton.disabled= false;
      exportButton.disabled = false;
      textButton.disabled = false;

      status.textContent = `${inventoryData.length} productos cargados. Listo para inventariar.`;
      renderInventory();
      prepareSpeech();
    };
    reader.readAsArrayBuffer(file);
  });

  // ========= RENDER DE LISTA =========
  function renderInventory() {
    inventoryList.innerHTML = '';
    inventoryData.forEach(item => {
      const row = document.createElement('div');
      row.className = 'inventory-item';

      const name = document.createElement('span');
      name.textContent = `${item['Producto']} (${item['UMB'] ?? ''})`;

      const qty = document.createElement('span');
      qty.textContent = `${item['Stock']}`;
      row.appendChild(name);
      row.appendChild(qty);
      inventoryList.appendChild(row);
    });
  }

  // ========= PARSER DE COMANDOS =========
  // Acepta:
  //  - "Producto 3"
  //  - "3 Producto"
  //  - "x3 Producto" o "Producto x3"
  //  - Con/ sin unidades, mayúsculas/acentos.
  function parseCommand(raw) {
    const command = raw
      .replace(',', ' ')
      .replace(/\s+/g, ' ')
      .trim();

    // Busca cantidad: números enteros al inicio, al final o con "x"
    // Casos: "3 producto", "producto 3", "x3 producto", "producto x3"
    const patterns = [
      /^(\d+)\s+(.+)$/i,       // 3 producto
      /^x(\d+)\s+(.+)$/i,      // x3 producto
      /^(.+)\s+(\d+)$/i,       // producto 3
      /^(.+)\s+x(\d+)$/i,      // producto x3
    ];

    for (const re of patterns) {
      const m = command.match(re);
      if (m) {
        const left = m[1];
        const right = m[2];

        if (re === patterns[0]) { // "3 producto"
          return { qty: parseInt(left, 10), name: right.trim() };
        }
        if (re === patterns[1]) { // "x3 producto"
          return { qty: parseInt(left, 10), name: right.trim() };
        }
        if (re === patterns[2]) { // "producto 3"
          return { qty: parseInt(right, 10), name: left.trim() };
        }
        if (re === patterns[3]) { // "producto x3"
          return { qty: parseInt(right, 10), name: left.trim() };
        }
      }
    }

    // Si no encuentra patrón válido:
    return null;
  }

  function applyCommand(parsed) {
    const { name, qty } = parsed;
    if (!fuse) return;

    const results = fuse.search(name);
    if (!results.length) {
      status.textContent = `Producto no encontrado: "${name}".`;
      return;
    }

    const found = results[0].item;
    found.Stock += qty;

    status.textContent = `Añadido: ${qty} a ${found.Producto}. Total: ${found.Stock}`;
    renderInventory();
  }

  // Para reutilizar en voz y en texto
  function processCommandLine(line) {
    const parsed = parseCommand(line);
    if (!parsed) {
      status.textContent = `No se entendió: "${line}". Usa "Producto 3" o "3 Producto".`;
      return;
    }
    applyCommand(parsed);
  }

  // ========= ENTRADA DE TEXTO MULTILÍNEA =========
  textButton.addEventListener('click', () => {
    const lines = textCommand.value
      .split('\n')
      .map(l => l.trim())
      .filter(Boolean);

    if (!lines.length) return;

    lines.forEach(line => processCommandLine(line));
    textCommand.value = '';
  });

  // ========= RECONOCIMIENTO DE VOZ =========
  function prepareSpeech() {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      micButton.disabled = true;
      status.textContent = "El reconocimiento de voz no es compatible en este navegador. Puedes pegar texto.";
      return;
    }

    const recognition = new SpeechRecognition();
    recognition.lang = 'es-ES';
    recognition.interimResults = false;

    micButton.addEventListener('click', () => {
      if (micButton.classList.contains('is-listening')) {
        recognition.stop();
      } else {
        recognition.start();
      }
    });

    recognition.onstart = () => {
      micButton.classList.add('is-listening');
      status.textContent = 'Escuchando...';
    };

    recognition.onend = () => {
      micButton.classList.remove('is-listening');
      status.textContent = 'Micrófono desactivado. Vuelve a pulsar para hablar.';
    };

    recognition.onerror = (event) => {
      status.textContent = `Error de micrófono: ${event.error}`;
      micButton.classList.remove('is-listening');
    };

    recognition.onresult = (event) => {
      const transcript = event.results[0][0].transcript.trim();
      processCommandLine(transcript);
    };
  }

  // ========= RESETEAR =========
  resetButton.addEventListener('click', () => {
    if (!inventoryData.length) return;
    if (confirm('¿Resetear todas las cantidades a 0?')) {
      inventoryData.forEach(item => item.Stock = 0);
      renderInventory();
      status.textContent = 'Cantidades reseteadas a 0.';
    }
  });

  // ========= EXPORTAR =========
  exportButton.addEventListener('click', () => {
    if (!originalWorkbook) return;

    const sheetName = originalWorkbook.SheetNames[0];
    const ws = originalWorkbook.Sheets[sheetName];

    // Asegura cabeceras y columna Stock
    const headerRow = XLSX.utils.sheet_to_json(ws, { header: 1 })[0] || [];
    let localHeaders = headerRow.slice();
    let stockCol = localHeaders.indexOf('Stock');

    if (stockCol === -1) {
      // Añade columna 'Stock' al final
      localHeaders.push('Stock');
      XLSX.utils.sheet_add_aoa(ws, [localHeaders], { origin: { r: 0, c: 0 } });
      stockCol = localHeaders.length - 1;
    }

    // Escribe las cantidades de inventoryData en filas correspondientes
    // Asumimos que el orden de inventoryData coincide con el orden original (sheet_to_json)
    inventoryData.forEach((item, idx) => {
      const cellRef = XLSX.utils.encode_cell({ r: idx + 1, c: stockCol });
      XLSX.utils.sheet_add_aoa(ws, [[item.Stock]], { origin: cellRef });
    });

    // Nombre de archivo por mes actual
    const monthNames = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
    const currentMonth = monthNames[new Date().getMonth()];
    const fileName = `Inventario Mejorado ${currentMonth}.xlsx`;

    XLSX.writeFile(originalWorkbook, fileName);
  });
});