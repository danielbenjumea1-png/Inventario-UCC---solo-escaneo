import ExcelJS from "exceljs";
import formidable from "formidable";

export const config = {
  api: { bodyParser: false },
};

/* =========================
   EXTRAER VALOR REAL CELDA
========================= */
function obtenerValorCelda(cell) {
  if (!cell || cell.value == null) return "";

  const v = cell.value;

  if (typeof v === "string" || typeof v === "number") {
    return String(v);
  }

  if (v.text) return String(v.text);

  if (v.richText) {
    return v.richText.map(r => r.text).join("");
  }

  if (v.result) return String(v.result);

  return String(v);
}

/* =========================
   NORMALIZADOR
========================= */
function normalizar(valor) {
  return String(valor)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s/g, "")
    .replace(/[^a-z0-9]/gi, "")
    .trim();
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
      const codigosEscaneados = new Set();
      const codigosOriginales = new Map();
      let totalEscaneados = 0;

      wsEscaneo.eachRow((row) => {
        row.eachCell((cell) => {
          const valor = obtenerValorCelda(cell);
          const limpio = normalizar(valor);

          if (limpio) {
            codigosEscaneados.add(limpio);
            codigosOriginales.set(limpio, valor);
            totalEscaneados++;
          }
        });
      });

      if (codigosEscaneados.size === 0) {
        return res.status(400).json({ error: "El archivo de escaneo no tiene códigos válidos" });
      }

      /* =========================
         CRUCE TOTAL SIN COLUMNAS
      ========================= */
      let coincidencias = 0;
      const encontrados = new Set();

      wsInventario.eachRow((row) => {
        row.eachCell((cell) => {

          const valor = obtenerValorCelda(cell);
          const limpio = normalizar(valor);

          if (!limpio) return;

          if (codigosEscaneados.has(limpio)) {

            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FF00FF00" }
            };

            coincidencias++;
            encontrados.add(limpio);
          }
        });
      });

      /* =========================
         CÓDIGOS NO ENCONTRADOS
      ========================= */
      const noEncontrados = [...codigosEscaneados].filter(c => !encontrados.has(c));

      if (noEncontrados.length > 0) {

        const inicio = wsInventario.rowCount + 2;

        wsInventario.getCell(`A${inicio}`).value =
          "CODIGOS ESCANEADOS QUE NO EXISTEN EN INVENTARIO";

        noEncontrados.forEach((codigo, index) => {
          wsInventario.getCell(`A${inicio + index + 1}`).value =
            codigosOriginales.get(codigo);
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
        detalle: error.message
      });

    }

  });

}
