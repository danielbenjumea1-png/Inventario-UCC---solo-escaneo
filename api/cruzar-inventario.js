import ExcelJS from "exceljs";
import formidable from "formidable";
import fs from "fs";

export const config = {
  api: {
    bodyParser: false
  }
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  const form = formidable({ multiples: true });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      return res.status(500).json({ error: err.message });
    }

    const inventarioFile = files.inventario?.[0];
    const escaneoFile = files.escaneo?.[0];

    if (!inventarioFile || !escaneoFile) {
      return res.status(400).json({ error: "Faltan archivos Excel" });
    }

    const wbInventario = new ExcelJS.Workbook();
    const wbEscaneo = new ExcelJS.Workbook();

    await wbInventario.xlsx.readFile(inventarioFile.filepath);
    await wbEscaneo.xlsx.readFile(escaneoFile.filepath);

    const wsInventario = wbInventario.worksheets[0];
    const wsEscaneo = wbEscaneo.worksheets[0];

    // Extraer códigos escaneados
    const codigosEscaneo = new Set();
    wsEscaneo.eachRow((row, i) => {
      if (i === 1) return;
      codigosEscaneo.add(String(row.getCell(1).value || "").trim());
    });

    // Encontrar columna codigo
    let colCodigo = null;
    wsInventario.getRow(1).eachCell((cell, col) => {
      if (String(cell.value).toLowerCase() === "codigo") {
        colCodigo = col;
      }
    });

    if (!colCodigo) {
      return res.status(400).json({ error: "No existe columna 'codigo'" });
    }

    let coincidencias = 0;

    wsInventario.eachRow((row, i) => {
      if (i === 1) return;
      const cell = row.getCell(colCodigo);
      const codigo = String(cell.value || "").trim();

      if (codigosEscaneo.has(codigo)) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF00FF00" }
        };
        coincidencias++;
      }
    });

    const buffer = await wbInventario.xlsx.writeBuffer();

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=inventario_cruzado.xlsx");
    res.send(buffer);
  });
}
