import ExcelJS from "exceljs";
import formidable from "formidable";
import fs from "fs";

export const config = {
  api: { bodyParser: false },
};

function esCodigoValido(valor) {
  if (!valor) return false;
  const texto = String(valor).trim();
  return /^[A-Za-z0-9]{6,20}$/.test(texto);
}

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

      const inventarioBuffer = fs.readFileSync(inventarioFile.filepath);
      const escaneoBuffer = fs.readFileSync(escaneoFile.filepath);

      const wbInventario = new ExcelJS.Workbook();
      const wbEscaneo = new ExcelJS.Workbook();

      await wbInventario.xlsx.load(inventarioBuffer);
      await wbEscaneo.xlsx.load(escaneoBuffer);

      const wsInventario = wbInventario.worksheets[0];
      const wsEscaneo = wbEscaneo.worksheets[0];

      // =========================
      // 1️⃣ EXTRAER CODIGOS ESCANEADOS EXACTOS
      // =========================
      const codigosEscaneados = new Set();

      wsEscaneo.eachRow(row => {
        row.eachCell(cell => {
          const valor = String(cell.value ?? "").trim();
          if (esCodigoValido(valor)) {
            codigosEscaneados.add(valor);
          }
        });
      });

      // =========================
      // 2️⃣ DETECTAR COLUMNA REAL DE CODIGOS EN INVENTARIO
      // =========================
      let mejorColumna = null;
      let maxCodigos = 0;

      const totalColumnas = wsInventario.columnCount;

      for (let col = 1; col <= totalColumnas; col++) {
        let contador = 0;

        wsInventario.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;
          const valor = String(row.getCell(col).value ?? "").trim();
          if (esCodigoValido(valor)) contador++;
        });

        if (contador > maxCodigos) {
          maxCodigos = contador;
          mejorColumna = col;
        }
      }

      if (!mejorColumna) {
        return res.status(400).json({
          error: "No se detectó columna de códigos en inventario"
        });
      }

      // =========================
      // 3️⃣ CRUCE ESTRICTO ===
      // =========================
      let coincidencias = 0;
      const encontrados = new Set();

      wsInventario.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const cell = row.getCell(mejorColumna);
        const valor = String(cell.value ?? "").trim();

        if (codigosEscaneados.has(valor)) {

          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF00FF00" }
          };

          coincidencias++;
          encontrados.add(valor);
        }
      });

      // =========================
      // 4️⃣ NO ENCONTRADOS
      // =========================
      const noEncontrados = [...codigosEscaneados]
        .filter(c => !encontrados.has(c));

      if (noEncontrados.length > 0) {

        const inicio = wsInventario.rowCount + 2;

        wsInventario.getCell(`A${inicio}`).value =
          "CODIGOS ESCANEADOS NO ENCONTRADOS";

        noEncontrados.forEach((codigo, i) => {
          wsInventario.getCell(`A${inicio + i + 1}`).value = codigo;
        });
      }

      // =========================
      // 5️⃣ RESUMEN
      // =========================
      const resumenFila = wsInventario.rowCount + 2;

      wsInventario.getCell(`C${resumenFila}`).value =
        `De ${codigosEscaneados.size} códigos escaneados se hallaron ${coincidencias} coincidencias`;

      const bufferFinal = await wbInventario.xlsx.writeBuffer();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.setHeader(
        "Content-Disposition",
        "attachment; filename=inventario_cruzado.xlsx"
      );

      res.send(bufferFinal);

      fs.unlinkSync(inventarioFile.filepath);
      fs.unlinkSync(escaneoFile.filepath);

    } catch (error) {
      return res.status(500).json({
        error: "Error procesando archivos",
        detalle: error.message
      });
    }

  });
}
