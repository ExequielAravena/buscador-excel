let datosExcel = [];

function obtenerElemento(id) {
  return document.getElementById(id);
}

function normalizarTexto(texto) {
  if (texto === null || texto === undefined) return "";
  return String(texto)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function formatearMiles(valor) {
  if (valor === null || valor === undefined || valor === "") return "";
  const numero = Number(valor);
  if (Number.isNaN(numero)) return String(valor);

  if (Number.isInteger(numero)) {
    return numero.toLocaleString("es-CL");
  }

  return numero.toLocaleString("es-CL", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function buscarFilaEncabezados(matriz) {
  for (let i = 0; i < matriz.length; i++) {
    const fila = (matriz[i] || []).map(x => String(x ?? "").trim());
    if (fila.includes("Codigo") && fila.includes("Producto")) {
      return i;
    }
  }
  return -1;
}

function manejarArchivoDirecto(event) {
  const nombreArchivo = obtenerElemento("nombreArchivo");
  const contador = obtenerElemento("contador");
  const archivo = event.target.files[0];

  if (!archivo) {
    nombreArchivo.textContent = "Ningún archivo seleccionado";
    datosExcel = [];
    renderizarFilas([]);
    contador.textContent = "Resultados encontrados: 0";
    return;
  }

  nombreArchivo.textContent = archivo.name;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const primeraHoja = workbook.SheetNames[0];
      const hoja = workbook.Sheets[primeraHoja];
      const matriz = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "" });

      if (!matriz || matriz.length === 0) {
        alert("El archivo está vacío o no tiene estructura válida.");
        datosExcel = [];
        renderizarFilas([]);
        contador.textContent = "Resultados encontrados: 0";
        return;
      }

      const idxHeader = buscarFilaEncabezados(matriz);

      if (idxHeader === -1) {
        alert("No se encontró la fila de encabezados (Codigo / Producto).");
        datosExcel = [];
        renderizarFilas([]);
        contador.textContent = "Resultados encontrados: 0";
        return;
      }

      const encabezados = matriz[idxHeader].map(x => String(x ?? "").trim());

      const idxCodigo = encabezados.indexOf("Codigo");
      const idxProducto = encabezados.indexOf("Producto");
      const idxBodega = encabezados.indexOf("Bodega");
      const idxDisponible = encabezados.indexOf("Disponible");
      const idxL1 = encabezados.indexOf("L1");
      const idxL2 = encabezados.indexOf("L2");
      const idxL3 = encabezados.indexOf("L3");

      const faltantes = [];
      if (idxCodigo === -1) faltantes.push("Codigo");
      if (idxProducto === -1) faltantes.push("Producto");
      if (idxBodega === -1) faltantes.push("Bodega");
      if (idxDisponible === -1) faltantes.push("Disponible");
      if (idxL1 === -1) faltantes.push("L1");
      if (idxL2 === -1) faltantes.push("L2");
      if (idxL3 === -1) faltantes.push("L3");

      if (faltantes.length > 0) {
        alert("Faltan columnas en el Excel: " + faltantes.join(", "));
        datosExcel = [];
        renderizarFilas([]);
        contador.textContent = "Resultados encontrados: 0";
        return;
      }

      datosExcel = matriz.slice(idxHeader + 1).map(fila => ({
        Codigo: fila[idxCodigo] ?? "",
        Producto: fila[idxProducto] ?? "",
        Bodega: fila[idxBodega] ?? "",
        Disponible: fila[idxDisponible] ?? "",
        L1: fila[idxL1] ?? "",
        L2: fila[idxL2] ?? "",
        L3: fila[idxL3] ?? ""
      }))
      .filter(item => {
        return String(item.Codigo).trim() !== "" || String(item.Producto).trim() !== "";
      });

      renderizarFilas(datosExcel);
      contador.textContent = `Resultados encontrados: ${datosExcel.length}`;
    } catch (error) {
      alert("Error al leer el archivo: " + error.message);
      datosExcel = [];
      renderizarFilas([]);
      contador.textContent = "Resultados encontrados: 0";
    }
  };

  reader.readAsArrayBuffer(archivo);
}

function renderizarBusquedaDirecta() {
  const contador = obtenerElemento("contador");
  const busquedaInput = obtenerElemento("busqueda");

  if (!datosExcel || datosExcel.length === 0) {
    renderizarFilas([]);
    contador.textContent = "Resultados encontrados: 0";
    return;
  }

  const texto = normalizarTexto(busquedaInput.value);

  if (!texto) {
    renderizarFilas(datosExcel);
    contador.textContent = `Resultados encontrados: ${datosExcel.length}`;
    return;
  }

  const palabras = texto.split(/\s+/).filter(Boolean);

  const filtrados = datosExcel.filter(item => {
    const base = normalizarTexto(item.Codigo + " " + item.Producto);
    return palabras.every(palabra => base.includes(palabra));
  });

  renderizarFilas(filtrados);
  contador.textContent = `Resultados encontrados: ${filtrados.length}`;
}

function limpiarBusquedaDirecta() {
  const busquedaInput = obtenerElemento("busqueda");
  const contador = obtenerElemento("contador");

  busquedaInput.value = "";

  if (datosExcel.length > 0) {
    renderizarFilas(datosExcel);
    contador.textContent = `Resultados encontrados: ${datosExcel.length}`;
  } else {
    renderizarFilas([]);
    contador.textContent = "Resultados encontrados: 0";
  }
}

function renderizarFilas(filas) {
  const tbodyResultados = obtenerElemento("tbodyResultados");
  tbodyResultados.innerHTML = "";

  if (!filas || filas.length === 0) {
    tbodyResultados.innerHTML = `
      <tr>
        <td colspan="8" class="sin-datos">No hay resultados para mostrar.</td>
      </tr>
    `;
    return;
  }

  filas.forEach(fila => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td class="codigo">${fila.Codigo}</td>
      <td class="producto">${fila.Producto}</td>
      <td class="bodega">${fila.Bodega}</td>
      <td class="numero">${formatearMiles(fila.Disponible)}</td>
      <td class="numero">${formatearMiles(fila.L1)}</td>
      <td class="numero">${formatearMiles(fila.L2)}</td>
      <td class="numero">${formatearMiles(fila.L3)}</td>
      <td><button class="copiar-btn" type="button">Copiar</button></td>
    `;

    const boton = tr.querySelector(".copiar-btn");
    boton.onclick = function() {
      copiarFila(fila);
    };

    tbodyResultados.appendChild(tr);
  });
}

function copiarFila(fila) {
  const texto = [
    fila.Codigo,
    fila.Producto,
    fila.Bodega,
    formatearMiles(fila.Disponible),
    formatearMiles(fila.L1),
    formatearMiles(fila.L2),
    formatearMiles(fila.L3)
  ].join(" | ");

  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(texto)
      .then(() => alert("Fila copiada"))
      .catch(() => copiarFallback(texto));
  } else {
    copiarFallback(texto);
  }
}

function copiarFallback(texto) {
  const area = document.createElement("textarea");
  area.value = texto;
  document.body.appendChild(area);
  area.select();
  try {
    document.execCommand("copy");
    alert("Fila copiada");
  } catch (e) {
    alert("No se pudo copiar");
  }
  document.body.removeChild(area);
}