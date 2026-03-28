import express from "express";
import multer from "multer";
import path from "path";
import fs from "fs";
import { createServer as createViteServer } from "vite";
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const XLSX = require("xlsx");
const PizZipRaw = require("pizzip");
const DocxtemplaterRaw = require("docxtemplater");
const JSZipRaw = require("jszip");

// Garantir acesso aos construtores corretos (lidando com ESM/CJS)
const PizZip = PizZipRaw.default || PizZipRaw;
const Docxtemplater = DocxtemplaterRaw.default || DocxtemplaterRaw;
const JSZip = JSZipRaw.default || JSZipRaw;

async function startServer() {
  const app = express();
  const PORT = 3000;

  const uploadDir = path.join(process.cwd(), "uploads");
  if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
  }

  const upload = multer({ dest: "uploads/" });

  // Health check
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // API routes
  app.post("/api/gerar", upload.fields([
    { name: "capa", maxCount: 1 },
    { name: "ficha", maxCount: 1 },
    { name: "xlsx", maxCount: 1 }
  ]), async (req: any, res) => {
    console.log("Recebida requisição para /api/gerar");
    const tempFiles: string[] = [];
    try {
      const files = req.files;
      if (!files?.capa?.[0] || !files?.ficha?.[0] || !files?.xlsx?.[0]) {
        console.warn("Arquivos faltando na requisição");
        return res.status(400).json({ error: "Arquivos faltando. Certifique-se de enviar os modelos de Capa, Ficha e a Planilha Excel." });
      }

      const capaPath = path.resolve(files.capa[0].path);
      const fichaPath = path.resolve(files.ficha[0].path);
      const xlsxPath = path.resolve(files.xlsx[0].path);
      
      tempFiles.push(capaPath, fichaPath, xlsxPath);

      console.log("Lendo planilha Excel...");
      const workbook = XLSX.readFile(xlsxPath);
      const sheets = workbook.SheetNames;
      
      if (!sheets.includes("PARECERES")) {
        console.warn("Aba 'PARECERES' não encontrada");
        return res.status(400).json({ error: "Aba 'PARECERES' não encontrada na planilha Excel." });
      }

      const sheetPareceres = XLSX.utils.sheet_to_json(workbook.Sheets["PARECERES"]) as any[];
      const mapaParecer: Record<string, string> = {};
      
      sheetPareceres.forEach(row => {
        const normalizedRow: any = {};
        Object.keys(row).forEach(key => {
          const normKey = key.trim().toLowerCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove acentos
            .replace(/ç/g, "c");
          normalizedRow[normKey] = row[key];
        });
        if (normalizedRow.codigo !== undefined && normalizedRow.texto !== undefined) {
          mapaParecer[String(normalizedRow.codigo).trim()] = String(normalizedRow.texto).trim();
        }
      });

      const resultsZip = new JSZip();
      
      console.log("Lendo modelos Word...");
      const capaTemplateContent = fs.readFileSync(capaPath);
      const fichaTemplateContent = fs.readFileSync(fichaPath);
      console.log(`Modelos lidos: Capa (${capaTemplateContent.length} bytes), Ficha (${fichaTemplateContent.length} bytes)`);

      // Iniciar processamento
      let totalFichas = 0;

      const renderDoc = (content: Buffer, data: any) => {
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
          delimiters: { start: "<<", end: ">>" },
          paragraphLoop: true,
          linebreaks: true,
        });
        doc.render(data);
        return doc.getZip().generate({ type: "nodebuffer" });
      };

      for (const sheetName of sheets) {
        if (sheetName === "PARECERES") continue;

        const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" }) as any[];
        if (rawData.length === 0) {
          console.log(`Aba ${sheetName} está vazia, pulando.`);
          continue;
        }

        console.log(`Processando turma: ${sheetName} (${rawData.length} alunos)`);
        const safeTurmaName = sheetName.replace(/[^a-zA-Z0-9]/g, "_") || "Turma_Sem_Nome";
        const turmaFolder = resultsZip.folder(safeTurmaName);

        // 1. Gerar Capa da Turma
        try {
          const capaBuffer = renderDoc(capaTemplateContent, {
            TURMA: sheetName,
            TURNO: String(rawData[0].turno || rawData[0].Turno || "").trim()
          });
          turmaFolder?.file("00_CAPA_DA_TURMA.docx", capaBuffer);
        } catch (e: any) {
          console.error(`Erro na capa da turma ${sheetName}:`, e.message);
        }

        // 2. Gerar Fichas dos Alunos
        for (let i = 0; i < rawData.length; i++) {
          const row = rawData[i];
          const normalizedRow: any = {};
          Object.keys(row).forEach(key => {
            const normKey = key.trim().toLowerCase()
              .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
              .replace(/ç/g, "c");
            normalizedRow[normKey] = row[key];
          });

          const nome = String(normalizedRow.nomealuno || normalizedRow.nome || "").trim();
          if (!nome) continue;

          const codParecer = String(normalizedRow.parecer || "").trim();
          const parecerLongText = mapaParecer[codParecer] || codParecer;
          const conceito = String(normalizedRow.parecertexto || normalizedRow.conceito || "").trim();
          const turno = String(normalizedRow.turno || "").trim();

          try {
            const fichaBuffer = renderDoc(fichaTemplateContent, {
              NOME: nome,
              PARECER: parecerLongText,
              CONCEITO: conceito,
              TURMA: sheetName,
              TURNO: turno
            });
            
            const safeNome = nome.replace(/[^a-zA-Z0-9]/g, "_") || `Aluno_${i + 1}`;
            turmaFolder?.file(`${safeNome}.docx`, fichaBuffer);
            totalFichas++;
          } catch (e: any) {
            console.error(`Erro na ficha do aluno ${nome}:`, e.message);
          }
        }
      }

      if (totalFichas === 0) {
        console.warn("Nenhuma ficha foi gerada");
        return res.status(400).json({ error: "Nenhum dado de aluno encontrado nas planilhas." });
      }

      console.log(`Gerando arquivo ZIP com ${totalFichas} fichas...`);
      const zipContent = await resultsZip.generateAsync({ 
        type: "nodebuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 }
      });
      
      console.log(`ZIP gerado com sucesso. Tamanho: ${(zipContent.length / 1024).toFixed(2)} KB`);

      res.set({
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="fichas_${Date.now()}.zip"`,
        "Content-Length": zipContent.length
      });
      
      res.end(zipContent);
      console.log("ZIP enviado com sucesso!");

    } catch (error: any) {
      console.error("Erro crítico no processamento:", error);
      res.status(500).json({ 
        error: "Erro interno no servidor", 
        details: error.message 
      });
    } finally {
      // Cleanup temp files
      tempFiles.forEach(p => {
        try {
          if (fs.existsSync(p)) fs.unlinkSync(p);
        } catch (e) {
          console.error(`Erro ao remover arquivo temporário ${p}:`, e);
        }
      });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
