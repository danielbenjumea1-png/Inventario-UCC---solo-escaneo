import ExcelJS from "exceljs";
import formidable from "formidable";
import fs from "fs";

export const config = {
  api: { bodyParser: false },
};

// ===============================
// NORMALIZADOR ROBUSTO
// ===============================
function normalizarTexto(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9]/g, "")
    .trim();
}

// ===============================
// DETECTOR DE COLUMNA
// ===============================
function esColumnaCodigo(valor) {
  const texto = normalizarTexto(valor);
  const palabrasClave = [
    "codigo","cod","codigo de barras","barcode","bar code","id","identificador",
    "serial","serialnumber","itemcode","productcode","sku","ean","upc"
  ];
  return palabrasClave.some((p) => texto.includes(p));
}

function encontrarColumnaCodigo(row) {
  let columna = null;
  row.eachCell((cell, col) => {
    if (esColumnaCodigo(cell.value)) columna = col;
  });
  return columna;
}

// ===============================
// HANDLER PRINCIPAL
// ===============================
export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Método no permitido" });

  const form = formidable({ multiples: true });
  form.parse(req, async (err, fields, files) => {
    if (err) return res.status(500).json({ error: err.message });

    const inventarioFile = files.inventario?.[0];
    const escaneoFile = files.escaneo?.[0];

    if (!inventarioFile || !escaneoFile) {
      return res.status(400).json({ error: "Faltan archivos Excel" });
    }

    try {
      // ===============================
      // EXTRAER CODIGOS DEL ESCANEO
      // ===============================
      const codigosEscaneo = new Set();
      const wbEscaneo = new ExcelJS.stream.xlsx.WorkbookReader(escaneoFile.filepath);

      for await (const worksheet of wbEscaneo) {
        let colCodigoEscaneo = null;
        for await (const row of worksheet) {
          if (row.number === 1) {
            colCodigoEscaneo = encontrarColumnaCodigo(row);
            if (!colCodigoEscaneo) break;
            continue;
          }
          if (colCodigoEscaneo) {
            const valor = row.getCell(colCodigoEscaneo).value;
            const codigo = normalizarTexto(valor?.toString());
            if (codigo) codigosEscaneo.add(codigo);
          }
        }
      }

      if (codigosEscaneo.size === 0) {
        return res.status(400).json({ error: "No se encontraron códigos en el Excel de escaneo" });
      }

      // ===============================
      // PROCESAR INVENTARIO
      // ===============================
      const wbInventarioReader = new ExcelJS.stream.xlsx.WorkbookReader(inventarioFile.filepath);
      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: "/tmp/inventario_cruzado.xlsx" });

      for await (const worksheetReader of wbInventarioReader) {
        const ws = workbook.addWorksheet(worksheetReader.name);
        let colCodigoInventario = null;
        const noCoinciden = []; // para guardar códigos no encontrados

        for await (const rowReader of worksheetReader) {
          const rowValues = [];
          rowReader.eachCell({ includeEmpty: true }, (cell) => rowValues.push(cell.value));

          if (rowReader.number === 1) {
            colCodigoInventario = encontrarColumnaCodigo(rowReader);
            ws.addRow(rowValues).commit();
            continue;
          }

          if (colCodigoInventario) {
            const codigo = normalizarTexto(rowValues[colCodigoInventario - 1]?.toString());
            if (codigosEscaneo.has(codigo)) {
              const newRow = ws.addRow(rowValues);
              newRow.getCell(colCodigoInventario).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FF00FF00" }, // verde
              };
              newRow.commit();
              continue;
            } else {
              noCoinciden.push(rowValues); // almacenar para al final
              continue;
            }
          }

          ws.addRow(rowValues).commit();
        }

        // ===============================
        // AGREGAR CÓDIGOS NO COINCIDENTES AL FINAL
        // ===============================
        if (noCoinciden.length > 0) {
          ws.addRow([]).commit(); // fila vacía para separar
          ws.addRow(["Códigos NO encontrados"]).commit();
          noCoinciden.forEach((fila) => {
            ws.addRow(fila).commit();
          });
        }
      }

      await workbook.commit();

      const buffer = fs.readFileSync("/tmp/inventario_cruzado.xlsx");
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", "attachment; filename=inventario_cruzado.xlsx");
      res.send(buffer);
    } catch (error) {
      return res.status(500).json({ error: "Error procesando archivos", detalle: error.message });
    }
  });
}
