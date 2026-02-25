import ExcelJS from "exceljs";
import formidable from "formidable";

export const config = {
  api: { bodyParser: false },
};

/* =========================
   NORMALIZADOR ROBUSTO
========================= */
function normalizarCodigo(valor) {
  return String(valor || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // quitar tildes
    .replace(/[^a-z0-9]/gi, "") // eliminar caracteres especiales
    .trim();
}

/* =========================
   DETECTOR INTELIGENTE
========================= */
function esColumnaCodigo(valor) {
  if (!valor) return false;

  const texto = String(valor)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();

  const claves = [
    "codigo",
    "cod",
    "codigo de barras",
    "codigos",
    "barcode",
    "bar code",
    "serial",
    "serial number",
    "id",
    "identificador",
    "identificacion",
    "identificacao",
    "sku",
    "ean",
    "upc",
  ];

  return claves.some(k => texto.includes(k));
}

/* =========================
   HANDLER
========================= */
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  const form = formidable({ multiples: true });

  form.parse(req, async (err, fields, files) => {
    if (err) return res.status(500).json({ error: err.message });

    const inventarioFile = files.inventario?.[0];
    const escaneoFile = files.escaneo?.[0];

    if (!inventarioFile || !escaneoFile) {
      return res.status(400).json({ error: "Faltan archivos Excel" });
    }

    try {
      const wbInventario = new ExcelJS.Workbook();
      const wbEscaneo = new ExcelJS.Workbook();

      await wbInventario.xlsx.readFile(inventarioFile.filepath);
      await wbEscaneo.xlsx.readFile(escaneoFile.filepath);

      const wsInventario = wbInventario.worksheets[0];
      const wsEscaneo = wbEscaneo.worksheets[0];

      /* =========================
         EXTRAER CÓDIGOS ESCANEADOS
      ========================= */
      const codigosEscaneo = new Set();
      let totalEscaneados = 0;

      wsEscaneo.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const valor = row.getCell(1).value;
        const codigo = normalizarCodigo(valor);

        if (codigo) {
          codigosEscaneo.add(codigo);
          totalEscaneados++;
        }
      });

      if (codigosEscaneo.size === 0) {
        return res.status(400).json({ error: "No se encontraron códigos válidos en el archivo de escaneo" });
      }

      /* =========================
         DETECTAR COLUMNA CÓDIGO
      ========================= */
      let colCodigo = null;

      wsInventario.getRow(1).eachCell((cell, col) => {
        if (esColumnaCodigo(cell.value)) {
          colCodigo = col;
        }
      });

      if (!colCodigo) {
        return res.status(400).json({ error: "No se encontró columna tipo código en inventario" });
      }

      /* =========================
         CRUCE
      ========================= */
      let coincidencias = 0;
      const noEncontrados = [];

      wsInventario.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const cell = row.getCell(colCodigo);
        const codigoInventario = normalizarCodigo(cell.value);

        if (!codigoInventario) return;

        if (codigosEscaneo.has(codigoInventario)) {
          // SOLO pinta la celda del código
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF00FF00" },
          };

          coincidencias++;
        } else {
          noEncontrados.push(codigoInventario);
        }
      });

      /* =========================
         AGREGAR NO ENCONTRADOS
      ========================= */
      if (noEncontrados.length > 0) {
        const filaInicio = wsInventario.rowCount + 2;

        wsInventario.getCell(`A${filaInicio}`).value = "CÓDIGOS NO ENCONTRADOS EN ESCANEO";

        noEncontrados.forEach((codigo, index) => {
          wsInventario.getCell(`A${filaInicio + index + 1}`).value = codigo;
        });
      }

      /* =========================
         GENERAR RESULTADO
      ========================= */
      const buffer = await wbInventario.xlsx.writeBuffer();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=inventario_cruzado.xlsx"
      );

      res.send(buffer);

      console.log(
        `De ${totalEscaneados} códigos escaneados se hallaron ${coincidencias} coincidencias`
      );

    } catch (error) {
      return res.status(500).json({
        error: "Error procesando archivos",
        detalle: error.message,
      });
    }
  });
}
