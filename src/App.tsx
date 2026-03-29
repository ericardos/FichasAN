import React, { useState, useRef, useEffect } from "react";
import { Upload, FileText, Table, CheckCircle2, AlertCircle, Loader2, Download, Heart, Copy, X, Sparkles, Info, Moon, Sun, HelpCircle, Users, Layers, ChevronRight } from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import JSZip from "jszip";

export default function App() {
  const [files, setFiles] = useState<{
    capa: File | null;
    ficha: File | null;
    xlsx: File | null;
  }>({
    capa: null,
    ficha: null,
    xlsx: null,
  });

  const [status, setStatus] = useState<"idle" | "processing" | "success" | "error">("idle");
  const [message, setMessage] = useState("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [showPixModal, setShowPixModal] = useState(false);
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [copied, setCopied] = useState(false);
  const [singleFilePerTurma, setSingleFilePerTurma] = useState(true);
  const [excelStats, setExcelStats] = useState<{ turmas: number; alunos: number } | null>(null);

  const pixKey = "00020101021126580014br.gov.bcb.pix0136fdf03993-fbdd-4b89-be41-6e63d23527295204000053039865802BR5925EDSON RICARDO DOS SANTOS 6009SAO PAULO622905251KMXMGASA03HZXG41BXQ478ZS6304122C";

  const fileInputRefs = {
    capa: useRef<HTMLInputElement>(null),
    ficha: useRef<HTMLInputElement>(null),
    xlsx: useRef<HTMLInputElement>(null),
  };

  const handleFileChange = (type: keyof typeof files) => async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    setFiles((prev) => ({ ...prev, [type]: file }));

    if (type === "xlsx" && file) {
      try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(buffer), { type: "array" });
        let totalAlunos = 0;
        let totalTurmas = 0;

        workbook.SheetNames.forEach(name => {
          if (name !== "PARECERES") {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
            if (data.length > 0) {
              totalTurmas++;
              totalAlunos += data.length;
            }
          }
        });
        setExcelStats({ turmas: totalTurmas, alunos: totalAlunos });
      } catch (err) {
        console.error("Erro ao ler stats do Excel:", err);
      }
    } else if (type === "xlsx" && !file) {
      setExcelStats(null);
    }
  };

  const handleGenerate = async () => {
    if (!files.capa || !files.ficha || !files.xlsx) {
      setStatus("error");
      setMessage("Por favor, selecione todos os arquivos necessários.");
      return;
    }

    setStatus("processing");
    setMessage("Lendo e processando arquivos localmente...");

    try {
      // 1. Ler os arquivos como ArrayBuffers
      const [capaBuf, fichaBuf, xlsxBuf] = await Promise.all([
        files.capa.arrayBuffer(),
        files.ficha.arrayBuffer(),
        files.xlsx.arrayBuffer()
      ]);

      // 2. Processar Excel
      const workbook = XLSX.read(new Uint8Array(xlsxBuf), { type: "array" });
      const sheets = workbook.SheetNames;

      if (!sheets.includes("PARECERES")) {
        throw new Error("Aba 'PARECERES' não encontrada na planilha Excel.");
      }

      const sheetPareceres = XLSX.utils.sheet_to_json(workbook.Sheets["PARECERES"]) as any[];
      const mapaParecer: Record<string, string> = {};
      
      sheetPareceres.forEach(row => {
        const normalizedRow: any = {};
        Object.keys(row).forEach(key => {
          const normKey = key.trim().toLowerCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
            .replace(/ç/g, "c");
          normalizedRow[normKey] = row[key];
        });
        if (normalizedRow.codigo !== undefined && normalizedRow.texto !== undefined) {
          mapaParecer[String(normalizedRow.codigo).trim()] = String(normalizedRow.texto).trim();
        }
      });

      const resultsZip = new JSZip();
      let totalFichas = 0;

      // Função para renderizar Word no navegador
      const renderDoc = (content: ArrayBuffer, data: any) => {
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
          delimiters: { start: "<<", end: ">>" },
          paragraphLoop: true,
          linebreaks: true,
        });
        doc.render(data);
        return doc.getZip().generate({ type: "blob" });
      };

      // Função "Super Loop" - Repete o conteúdo do documento para cada aluno
      // Isso evita o erro de "letras vermelhas" pois não há união de arquivos externos
      const renderSuperLoop = (content: ArrayBuffer, studentsData: any[]) => {
        const zip = new PizZip(content);
        let xml = zip.file("word/document.xml").asText();
        
        // Estratégia robusta: Localiza o corpo e as propriedades da seção final
        const bodyStartIdx = xml.indexOf("<w:body>");
        const lastSectPrIdx = xml.lastIndexOf("<w:sectPr");
        
        if (bodyStartIdx !== -1 && lastSectPrIdx !== -1 && lastSectPrIdx > bodyStartIdx) {
          const prefix = xml.substring(0, bodyStartIdx + 8);
          const mainContent = xml.substring(bodyStartIdx + 8, lastSectPrIdx);
          const suffix = xml.substring(lastSectPrIdx);
          
          // Envolve o conteúdo principal no loop e adiciona quebra de página
          // Usamos os delimitadores << >> do usuário, mas envelopados em XML válido
          // CRITICAL: Escapamos os caracteres < e > para não quebrar o XML
          const loopStart = '<w:p><w:r><w:t>&lt;&lt;#_fichas&gt;&gt;</w:t></w:r></w:p>';
          const loopEnd = '<w:p><w:r><w:br w:type="page"/><w:t>&lt;&lt;/_fichas&gt;&gt;</w:t></w:r></w:p>';
          xml = `${prefix}${loopStart}${mainContent}${loopEnd}${suffix}`;
          zip.file("word/document.xml", xml);
        }
        
        const doc = new Docxtemplater(zip, {
          delimiters: { start: "<<", end: ">>" },
          paragraphLoop: true,
          linebreaks: true,
        });
        
        doc.render({ _fichas: studentsData });
        return doc.getZip().generate({ type: "blob" });
      };

      // 3. Processar cada aba/turma
      for (const sheetName of sheets) {
        if (sheetName === "PARECERES") continue;

        const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" }) as any[];
        if (rawData.length === 0) continue;

        const safeTurmaName = sheetName.replace(/[^a-zA-Z0-9]/g, "_") || "Turma_Sem_Nome";
        const turmaFolder = resultsZip.folder(safeTurmaName);
        
        const studentsList: any[] = [];

        // Preparar dados de todos os alunos
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

          const studentData = {
            NOME: nome,
            PARECER: parecerLongText,
            CONCEITO: conceito,
            TURMA: sheetName,
            TURNO: turno
          };

          if (singleFilePerTurma) {
            studentsList.push(studentData);
          } else {
            try {
              const fichaBlob = renderDoc(fichaBuf, studentData);
              const safeNome = nome.replace(/[^a-zA-Z0-9]/g, "_") || `Aluno_${i + 1}`;
              turmaFolder?.file(`${safeNome}.docx`, fichaBlob);
            } catch (e: any) {
              console.error(`Erro na ficha do aluno ${nome}:`, e.message);
            }
          }
          totalFichas++;
        }

        // Se for arquivo único, usar a estratégia Super Loop + Capa
        if (singleFilePerTurma && studentsList.length > 0) {
          try {
            // 1. Gera um único documento com todas as fichas
            const allFichasBlob = renderSuperLoop(fichaBuf, studentsList);
            
            // 2. Gera a capa
            const capaBlob = renderDoc(capaBuf, {
              TURMA: sheetName,
              TURNO: String(rawData[0].turno || rawData[0].Turno || "").trim()
            });

            // 3. Une APENAS a Capa com o blocão de fichas
            // @ts-ignore
            const DocxMergerModule = await import("docx-merger");
            const DocxMerger = DocxMergerModule.default || DocxMergerModule;
            
            const merger = new DocxMerger({ pageBreak: true }, [
              await capaBlob.arrayBuffer(), 
              await allFichasBlob.arrayBuffer()
            ]);
            
            await new Promise((resolve, reject) => {
              merger.save("blob", (blob: Blob) => {
                turmaFolder?.file(`DOCUMENTO_UNICO_${safeTurmaName}.docx`, blob);
                resolve(true);
              }, (err: any) => reject(err));
            });
          } catch (error: any) {
            console.error("Erro ao gerar arquivo único, tentando fallback individual:", error);
            // FALLBACK: Se a união falhar, gera os arquivos individuais para não travar o usuário
            try {
              const capaBlob = renderDoc(capaBuf, {
                TURMA: sheetName,
                TURNO: String(rawData[0].turno || rawData[0].Turno || "").trim()
              });
              turmaFolder?.file("00_CAPA_DA_TURMA.docx", capaBlob);

              for (let i = 0; i < studentsList.length; i++) {
                const s = studentsList[i];
                const fBlob = renderDoc(fichaBuf, s);
                const sNome = s.NOME.replace(/[^a-zA-Z0-9]/g, "_") || `Aluno_${i + 1}`;
                turmaFolder?.file(`${sNome}.docx`, fBlob);
              }
            } catch (fallbackError) {
              console.error("Erro no fallback:", fallbackError);
            }
          }
        } else if (!singleFilePerTurma) {
          // Gerar capa individual se não for arquivo único
          try {
            const capaBlob = renderDoc(capaBuf, {
              TURMA: sheetName,
              TURNO: String(rawData[0].turno || rawData[0].Turno || "").trim()
            });
            turmaFolder?.file("00_CAPA_DA_TURMA.docx", capaBlob);
          } catch (e: any) {
            console.error(`Erro na capa da turma ${sheetName}:`, e.message);
          }
        }
      }

      if (totalFichas === 0) {
        throw new Error("Nenhum dado de aluno encontrado nas planilhas.");
      }

      setMessage("Gerando arquivo ZIP final...");
      const zipBlob = await resultsZip.generateAsync({ 
        type: "blob",
        compression: "STORE"
      });
      
      const url = window.URL.createObjectURL(zipBlob);
      setDownloadUrl(url);
      setStatus("success");
      setMessage(`Sucesso! ${totalFichas} fichas foram geradas.`);

    } catch (error: any) {
      console.error("Erro no processamento client-side:", error);
      setStatus("error");
      setMessage(error.message || "Erro inesperado ao processar arquivos.");
    }
  };

  const copyPixKey = () => {
    navigator.clipboard.writeText(pixKey);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className={`min-h-screen transition-colors duration-500 ${isDarkMode ? 'bg-[#0F172A] text-slate-100' : 'bg-[#F8F9FA] text-[#1A1C1E]'} font-sans selection:bg-amber-100`}>
      {/* Background Pattern */}
      <div className={`fixed inset-0 z-0 opacity-[0.03] pointer-events-none ${isDarkMode ? 'invert' : ''}`} 
           style={{ backgroundImage: 'radial-gradient(#000 1px, transparent 1px)', backgroundSize: '24px 24px' }}></div>

      <div className="relative z-10 max-w-5xl mx-auto px-6 py-12 md:py-16">
        <header className="mb-8 flex flex-col md:flex-row items-start md:items-center justify-between gap-8">
          <div className="text-left">
            <motion.div 
              initial={{ opacity: 0, x: -20 }}
              animate={{ opacity: 1, x: 0 }}
              className={`inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider mb-4 border ${
                isDarkMode ? 'bg-amber-900/20 text-amber-400 border-amber-900/30' : 'bg-amber-50 text-amber-700 border-amber-100'
              }`}
            >
              <Sparkles className="w-3 h-3" />
              Processador Inteligente
            </motion.div>
            <h1 className={`text-5xl md:text-6xl font-black tracking-tight mb-4 ${isDarkMode ? 'text-white' : 'text-[#0F172A]'}`}>
              Fichas<span className="text-amber-600">.</span>io
            </h1>
            <p className={`text-lg max-w-md leading-relaxed ${isDarkMode ? 'text-slate-400' : 'text-slate-500'}`}>
              Transforme planilhas complexas em documentos Word perfeitamente formatados em segundos.
            </p>
          </div>

          <div className="flex flex-wrap gap-3">
            <button 
              onClick={() => setIsDarkMode(!isDarkMode)}
              className={`p-4 rounded-2xl border transition-all ${
                isDarkMode ? 'bg-slate-800 border-slate-700 text-amber-400 hover:bg-slate-700' : 'bg-white border-slate-200 text-slate-400 hover:bg-slate-50'
              }`}
              title="Alternar Tema"
            >
              {isDarkMode ? <Sun className="w-6 h-6" /> : <Moon className="w-6 h-6" />}
            </button>
            <button 
              onClick={() => setShowHelpModal(true)}
              className={`p-4 rounded-2xl border transition-all flex items-center gap-2 font-bold text-sm ${
                isDarkMode ? 'bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-700' : 'bg-white border-slate-200 text-slate-500 hover:bg-slate-50'
              }`}
            >
              <HelpCircle className="w-6 h-6" />
              <span className="hidden sm:inline">COMO USAR</span>
            </button>
          </div>
        </header>

        {/* Support Banner */}
        <motion.div 
          initial={{ opacity: 0, y: -20 }}
          animate={{ opacity: 1, y: 0 }}
          className={`mb-12 p-6 rounded-[32px] border-2 border-dashed flex flex-col md:flex-row items-center justify-between gap-6 transition-all ${
            isDarkMode 
              ? 'bg-amber-500/5 border-amber-500/20 text-amber-200' 
              : 'bg-amber-50 border-amber-200 text-amber-900'
          }`}
        >
          <div className="flex items-center gap-4 text-center md:text-left">
            <div className="w-12 h-12 rounded-2xl bg-amber-500 flex items-center justify-center flex-shrink-0 shadow-lg shadow-amber-500/20">
              <Heart className="w-6 h-6 text-white fill-white" />
            </div>
            <div>
              <h4 className="font-black uppercase tracking-tight text-sm">Este projeto é mantido por você!</h4>
              <p className="text-xs opacity-70">O Fichas.io é gratuito e sem anúncios. Sua doação ajuda a pagar os servidores.</p>
            </div>
          </div>
          <button 
            onClick={() => setShowPixModal(true)}
            className="px-8 py-3 bg-amber-600 text-white rounded-full font-black text-xs uppercase tracking-widest hover:bg-amber-700 transition-all shadow-lg shadow-amber-600/20 active:scale-95"
          >
            Apoiar agora
          </button>
        </motion.div>

        <main className="grid gap-8">
          {/* Bento Grid File Selection */}
          <section className="grid grid-cols-1 md:grid-cols-12 gap-4">
            <div className="md:col-span-8 grid grid-cols-1 sm:grid-cols-2 gap-4">
              <FileCard
                title="Modelo CAPA"
                subtitle="Documento .docx"
                icon={<FileText className="w-6 h-6" />}
                file={files.capa}
                isDarkMode={isDarkMode}
                onClick={() => fileInputRefs.capa.current?.click()}
              />
              <FileCard
                title="Modelo FICHA"
                subtitle="Documento .docx"
                icon={<FileText className="w-6 h-6" />}
                file={files.ficha}
                isDarkMode={isDarkMode}
                onClick={() => fileInputRefs.ficha.current?.click()}
              />
            </div>
            <div className="md:col-span-4">
              <FileCard
                title="Planilha XLSX"
                subtitle="Dados dos alunos"
                icon={<Table className="w-6 h-6" />}
                file={files.xlsx}
                isDarkMode={isDarkMode}
                fullHeight
                onClick={() => fileInputRefs.xlsx.current?.click()}
              />
            </div>
          </section>

          {/* Hidden Inputs */}
          <input type="file" ref={fileInputRefs.capa} onChange={handleFileChange("capa")} accept=".docx" className="hidden" />
          <input type="file" ref={fileInputRefs.ficha} onChange={handleFileChange("ficha")} accept=".docx" className="hidden" />
          <input type="file" ref={fileInputRefs.xlsx} onChange={handleFileChange("xlsx")} accept=".xlsx" className="hidden" />

          {/* Action Card */}
          <div className={`p-8 md:p-12 rounded-[40px] shadow-2xl border relative overflow-hidden transition-colors ${
            isDarkMode ? 'bg-slate-900 border-slate-800 shadow-black/50' : 'bg-white border-slate-100 shadow-slate-200/50'
          }`}>
            {/* Decorative elements */}
            <div className={`absolute top-0 right-0 w-64 h-64 rounded-full -mr-32 -mt-32 blur-3xl opacity-20 ${isDarkMode ? 'bg-amber-500' : 'bg-amber-200'}`}></div>
            
            <div className="relative z-10 flex flex-col items-center justify-center text-center">
              <AnimatePresence mode="wait">
                {status === "idle" && (
                  <motion.div
                    key="idle"
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -20 }}
                    className="flex flex-col items-center w-full"
                  >
                    {/* Excel Stats Preview */}
                    {excelStats && (
                      <motion.div 
                        initial={{ opacity: 0, scale: 0.9 }}
                        animate={{ opacity: 1, scale: 1 }}
                        className={`mb-8 flex gap-6 p-6 rounded-3xl border ${
                          isDarkMode ? 'bg-slate-800/50 border-slate-700' : 'bg-slate-50 border-slate-100'
                        }`}
                      >
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-xl bg-blue-500/10 flex items-center justify-center text-blue-500">
                            <Layers className="w-5 h-5" />
                          </div>
                          <div className="text-left">
                            <p className="text-[10px] uppercase font-black tracking-widest text-slate-400">Turmas</p>
                            <p className="text-xl font-black">{excelStats.turmas}</p>
                          </div>
                        </div>
                        <div className="w-px h-10 bg-slate-200 self-center"></div>
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-xl bg-green-500/10 flex items-center justify-center text-green-500">
                            <Users className="w-5 h-5" />
                          </div>
                          <div className="text-left">
                            <p className="text-[10px] uppercase font-black tracking-widest text-slate-400">Alunos</p>
                            <p className="text-xl font-black">{excelStats.alunos}</p>
                          </div>
                        </div>
                      </motion.div>
                    )}

                    <div className={`w-20 h-20 rounded-full flex items-center justify-center mb-8 ${isDarkMode ? 'bg-slate-800' : 'bg-slate-50'}`}>
                      <Upload className={`w-8 h-8 ${isDarkMode ? 'text-slate-600' : 'text-slate-400'}`} />
                    </div>
                    <h2 className={`text-2xl font-black uppercase tracking-tight mb-4 ${isDarkMode ? 'text-white' : 'text-[#0F172A]'}`}>Pronto para começar?</h2>
                    <p className={`mb-8 max-w-xs ${isDarkMode ? 'text-slate-400' : 'text-slate-500'}`}>Selecione os arquivos acima para habilitar a geração dos documentos.</p>
                    
                    <div className="flex flex-col sm:flex-row items-center gap-6 mb-10">
                      {/* Toggle Option */}
                      <div className={`flex items-center gap-3 p-4 rounded-2xl border transition-colors ${
                        isDarkMode ? 'bg-slate-800/50 border-slate-700' : 'bg-slate-50 border-slate-100'
                      }`}>
                        <div 
                          onClick={() => setSingleFilePerTurma(!singleFilePerTurma)}
                          className={`w-12 h-6 rounded-full relative cursor-pointer transition-colors ${singleFilePerTurma ? 'bg-amber-500' : 'bg-slate-300'}`}
                        >
                          <motion.div 
                            animate={{ x: singleFilePerTurma ? 24 : 4 }}
                            className="absolute top-1 w-4 h-4 bg-white rounded-full shadow-sm"
                          />
                        </div>
                        <span className={`text-sm font-bold ${isDarkMode ? 'text-slate-300' : 'text-slate-600'}`}>Arquivo Único por Turma</span>
                      </div>
                    </div>

                    <button
                      onClick={handleGenerate}
                      disabled={!files.capa || !files.ficha || !files.xlsx}
                      className={`group relative px-12 py-6 rounded-full font-black text-xl transition-all flex items-center gap-3 overflow-hidden ${
                        !files.capa || !files.ficha || !files.xlsx
                          ? 'bg-slate-200 text-slate-400 cursor-not-allowed opacity-50'
                          : 'bg-[#0F172A] text-white hover:scale-105 hover:shadow-2xl hover:shadow-amber-500/20 active:scale-95'
                      }`}
                    >
                      {files.capa && files.ficha && files.xlsx && (
                        <motion.div 
                          layoutId="glow"
                          className="absolute inset-0 bg-gradient-to-r from-amber-500/20 via-transparent to-amber-500/20 animate-pulse"
                        />
                      )}
                      <Sparkles className={`w-6 h-6 ${files.capa && files.ficha && files.xlsx ? 'text-amber-400' : 'text-slate-400'}`} />
                      GERAR DOCUMENTOS
                      <ChevronRight className="w-6 h-6 opacity-0 group-hover:opacity-100 transition-all group-hover:translate-x-1" />
                    </button>
                  </motion.div>
                )}

                {status === "processing" && (
                  <motion.div
                    key="processing"
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="flex flex-col items-center gap-8"
                  >
                    <div className="relative">
                      <div className="w-28 h-28 border-4 border-slate-100 border-t-amber-500 rounded-full animate-spin"></div>
                      <div className="absolute inset-0 flex items-center justify-center">
                        <Loader2 className="w-10 h-10 text-amber-500 animate-pulse" />
                      </div>
                    </div>
                    <div className="text-center">
                      <h3 className={`text-2xl font-black uppercase tracking-tight mb-2 ${isDarkMode ? 'text-white' : 'text-[#0F172A]'}`}>Processando...</h3>
                      <p className="text-amber-500 font-bold animate-pulse">{message}</p>
                    </div>
                  </motion.div>
                )}

                {status === "success" && (
                  <motion.div
                    key="success"
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="flex flex-col items-center"
                  >
                    <div className="w-24 h-24 bg-green-500/10 rounded-full flex items-center justify-center mb-8">
                      <CheckCircle2 className="w-12 h-12 text-green-500" />
                    </div>
                    <h3 className={`text-3xl font-black uppercase tracking-tight mb-2 ${isDarkMode ? 'text-white' : 'text-[#0F172A]'}`}>Tudo pronto!</h3>
                    <p className={`mb-8 font-medium ${isDarkMode ? 'text-slate-400' : 'text-slate-500'}`}>{message}</p>
                    
                    {/* Support Call to Action on Success */}
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ delay: 0.5 }}
                      className={`mb-10 p-6 rounded-3xl border-2 border-dashed max-w-md ${
                        isDarkMode ? 'bg-amber-500/5 border-amber-500/20' : 'bg-amber-50 border-amber-200'
                      }`}
                    >
                      <p className={`text-sm mb-4 font-bold ${isDarkMode ? 'text-amber-400' : 'text-amber-700'}`}>
                        Economizou seu tempo? Considere apoiar o projeto com qualquer valor para que ele continue gratuito!
                      </p>
                      <button 
                        onClick={() => setShowPixModal(true)}
                        className="flex items-center gap-2 text-xs font-black uppercase tracking-widest text-amber-600 hover:text-amber-700 transition-colors mx-auto"
                      >
                        <Heart className="w-4 h-4 fill-amber-600" />
                        Apoiar via PIX
                      </button>
                    </motion.div>

                    <div className="flex flex-col sm:flex-row gap-4 w-full sm:w-auto">
                      <a
                        href={downloadUrl!}
                        download="fichas_escolares.zip"
                        className="px-10 py-5 bg-green-500 text-white rounded-full font-black text-lg transition-all hover:scale-105 hover:shadow-2xl hover:shadow-green-500/30 flex items-center justify-center gap-3"
                      >
                        <Download className="w-6 h-6" />
                        BAIXAR ZIP
                      </a>
                      <button
                        onClick={() => {
                          setStatus("idle");
                          setFiles({ capa: null, ficha: null, xlsx: null });
                          setExcelStats(null);
                        }}
                        className={`px-10 py-5 rounded-full font-black text-lg transition-all ${
                          isDarkMode ? 'bg-slate-800 text-slate-300 hover:bg-slate-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                        }`}
                      >
                        REINICIAR
                      </button>
                    </div>
                  </motion.div>
                )}

                {status === "error" && (
                  <motion.div
                    key="error"
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="flex flex-col items-center"
                  >
                    <div className="w-24 h-24 bg-red-500/10 rounded-full flex items-center justify-center mb-8">
                      <AlertCircle className="w-12 h-12 text-red-500" />
                    </div>
                    <h3 className={`text-2xl font-black uppercase tracking-tight mb-2 ${isDarkMode ? 'text-white' : 'text-[#0F172A]'}`}>Ops! Algo deu errado</h3>
                    <p className="text-red-500 font-bold mb-10 max-w-md">{message}</p>
                    <button
                      onClick={() => setStatus("idle")}
                      className="px-12 py-5 bg-red-500 text-white rounded-full font-black transition-all hover:scale-105 shadow-xl shadow-red-500/20"
                    >
                      TENTAR NOVAMENTE
                    </button>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </div>
        </main>

        <footer className={`mt-20 pt-10 border-t flex flex-col md:flex-row items-center justify-between gap-6 transition-colors ${
          isDarkMode ? 'border-slate-800 text-slate-500' : 'border-slate-200 text-slate-400'
        }`}>
          <div className={`flex items-center gap-2 font-black ${isDarkMode ? 'text-slate-700' : 'text-slate-200'}`}>
            <Sparkles className="w-4 h-4" />
            FICHAS.IO
          </div>
          <p className="text-sm font-medium">Desenvolvido para facilitar a vida de educadores • v2.9 (UX Enhanced) • 2026</p>
          <div className="flex items-center gap-6">
            <button 
              onClick={() => setShowHelpModal(true)}
              className="text-sm font-bold hover:text-amber-600 transition-colors"
            >
              Instruções
            </button>
            <button 
              onClick={() => setShowPixModal(true)}
              className="flex items-center gap-2 text-sm font-bold text-amber-600 hover:text-amber-700 transition-colors"
            >
              <Heart className="w-4 h-4 fill-amber-600" />
              Apoie o Projeto
            </button>
          </div>
        </footer>
      </div>

      {/* Floating Support Button */}
      <motion.button 
        initial={{ scale: 0, opacity: 0 }}
        animate={{ scale: 1, opacity: 1 }}
        whileHover={{ scale: 1.1 }}
        whileTap={{ scale: 0.9 }}
        onClick={() => setShowPixModal(true)}
        className="fixed bottom-8 right-8 z-40 w-16 h-16 bg-amber-600 text-white rounded-full shadow-2xl flex items-center justify-center group"
      >
        <Heart className="w-8 h-8 fill-white group-hover:scale-110 transition-transform" />
        <div className="absolute right-full mr-4 bg-[#0F172A] text-white px-4 py-2 rounded-xl text-xs font-black uppercase tracking-widest opacity-0 group-hover:opacity-100 transition-opacity whitespace-nowrap pointer-events-none">
          Apoie o Projeto
        </div>
      </motion.button>

      {/* PIX Modal */}
      <AnimatePresence>
        {showPixModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowPixModal(false)}
              className="absolute inset-0 bg-[#0F172A]/80 backdrop-blur-sm"
            ></motion.div>
            
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className={`relative w-full max-w-md rounded-[40px] p-10 text-center shadow-2xl ${isDarkMode ? 'bg-slate-900' : 'bg-white'}`}
            >
              <button 
                onClick={() => setShowPixModal(false)}
                className={`absolute top-6 right-6 p-2 rounded-full transition-colors ${isDarkMode ? 'hover:bg-slate-800' : 'hover:bg-slate-100'}`}
              >
                <X className="w-6 h-6 text-slate-400" />
              </button>

              <div className="w-16 h-16 bg-amber-500/10 rounded-2xl flex items-center justify-center mx-auto mb-6">
                <Heart className="w-8 h-8 text-amber-600 fill-amber-600" />
              </div>

              <h2 className={`text-2xl font-black uppercase tracking-tight mb-2 ${isDarkMode ? 'text-white' : 'text-[#0F172A]'}`}>Apoio Voluntário</h2>
              <p className={`mb-8 font-medium ${isDarkMode ? 'text-slate-400' : 'text-slate-500'}`}>Sua contribuição ajuda a manter o site no ar e gratuito!</p>

              <div className={`p-8 rounded-[32px] mb-8 ${isDarkMode ? 'bg-slate-800' : 'bg-slate-50'}`}>
                <div className="bg-white p-4 rounded-2xl shadow-sm inline-block mb-6 border border-slate-100">
                  <img 
                    src={`https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(pixKey)}`} 
                    alt="PIX QR Code"
                    className="w-[180px] h-[180px]"
                  />
                </div>

                <div className="flex gap-2">
                  <div className={`flex-1 border rounded-2xl px-4 py-3 text-xs truncate flex items-center ${
                    isDarkMode ? 'bg-slate-900 border-slate-700 text-slate-500' : 'bg-white border-slate-200 text-slate-400'
                  }`}>
                    {pixKey}
                  </div>
                  <button 
                    onClick={copyPixKey}
                    className={`p-4 rounded-2xl transition-all ${copied ? 'bg-green-500 text-white' : 'bg-amber-600 text-white hover:bg-amber-700'}`}
                  >
                    {copied ? <CheckCircle2 className="w-5 h-5" /> : <Copy className="w-5 h-5" />}
                  </button>
                </div>
              </div>

              <button 
                onClick={() => setShowPixModal(false)}
                className="text-slate-400 font-black uppercase tracking-widest text-sm hover:text-slate-600 transition-colors"
              >
                Fechar
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Help Modal */}
      <AnimatePresence>
        {showHelpModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowHelpModal(false)}
              className="absolute inset-0 bg-[#0F172A]/80 backdrop-blur-sm"
            ></motion.div>
            
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className={`relative w-full max-w-2xl rounded-[40px] p-8 md:p-12 shadow-2xl overflow-y-auto max-h-[90vh] ${isDarkMode ? 'bg-slate-900' : 'bg-white'}`}
            >
              <button 
                onClick={() => setShowHelpModal(false)}
                className={`absolute top-6 right-6 p-2 rounded-full transition-colors ${isDarkMode ? 'hover:bg-slate-800' : 'hover:bg-slate-100'}`}
              >
                <X className="w-6 h-6 text-slate-400" />
              </button>

              <div className="flex items-center gap-4 mb-8">
                <div className="w-12 h-12 bg-blue-500/10 rounded-2xl flex items-center justify-center text-blue-500">
                  <Info className="w-6 h-6" />
                </div>
                <h2 className={`text-3xl font-black uppercase tracking-tight ${isDarkMode ? 'text-white' : 'text-[#0F172A]'}`}>Como Usar</h2>
              </div>

              <div className="space-y-8 text-left">
                <section>
                  <h3 className={`text-lg font-black uppercase tracking-widest mb-4 flex items-center gap-2 ${isDarkMode ? 'text-amber-400' : 'text-amber-600'}`}>
                    <span className="w-6 h-6 rounded-full bg-amber-500/10 flex items-center justify-center text-xs">1</span>
                    Planilha Excel
                  </h3>
                  <p className={`mb-4 ${isDarkMode ? 'text-slate-400' : 'text-slate-600'}`}>
                    Sua planilha deve ter abas para cada turma. Uma aba especial chamada <strong className={isDarkMode ? 'text-white' : 'text-slate-900'}>PARECERES</strong> deve conter os textos longos.
                  </p>
                  <div className={`p-4 rounded-2xl border ${isDarkMode ? 'bg-slate-800 border-slate-700' : 'bg-slate-50 border-slate-100'}`}>
                    <p className="text-xs font-bold uppercase tracking-widest text-slate-400 mb-2">Colunas Necessárias:</p>
                    <ul className={`text-sm space-y-1 ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                      <li>• <strong>NOMEALUNO</strong>: Nome completo do aluno</li>
                      <li>• <strong>PARECER</strong>: Código do parecer (ex: 1, 2, A, B)</li>
                      <li>• <strong>CONCEITO</strong>: Nota ou conceito rápido</li>
                    </ul>
                  </div>
                </section>

                <section>
                  <h3 className={`text-lg font-black uppercase tracking-widest mb-4 flex items-center gap-2 ${isDarkMode ? 'text-amber-400' : 'text-amber-600'}`}>
                    <span className="w-6 h-6 rounded-full bg-amber-500/10 flex items-center justify-center text-xs">2</span>
                    Modelos Word (.docx)
                  </h3>
                  <p className={`mb-4 ${isDarkMode ? 'text-slate-400' : 'text-slate-600'}`}>
                    Use marcadores entre chaves duplas para indicar onde os dados devem entrar:
                  </p>
                  <div className="grid grid-cols-2 gap-2">
                    <div className={`p-3 rounded-xl border text-center ${isDarkMode ? 'bg-slate-800 border-slate-700' : 'bg-slate-50 border-slate-100'}`}>
                      <code className="text-amber-500 font-bold">{'<<NOME>>'}</code>
                    </div>
                    <div className={`p-3 rounded-xl border text-center ${isDarkMode ? 'bg-slate-800 border-slate-700' : 'bg-slate-50 border-slate-100'}`}>
                      <code className="text-amber-500 font-bold">{'<<PARECER>>'}</code>
                    </div>
                    <div className={`p-3 rounded-xl border text-center ${isDarkMode ? 'bg-slate-800 border-slate-700' : 'bg-slate-50 border-slate-100'}`}>
                      <code className="text-amber-500 font-bold">{'<<TURMA>>'}</code>
                    </div>
                    <div className={`p-3 rounded-xl border text-center ${isDarkMode ? 'bg-slate-800 border-slate-700' : 'bg-slate-50 border-slate-100'}`}>
                      <code className="text-amber-500 font-bold">{'<<CONCEITO>>'}</code>
                    </div>
                  </div>
                </section>

                <section>
                  <h3 className={`text-lg font-black uppercase tracking-widest mb-4 flex items-center gap-2 ${isDarkMode ? 'text-amber-400' : 'text-amber-600'}`}>
                    <span className="w-6 h-6 rounded-full bg-amber-500/10 flex items-center justify-center text-xs">3</span>
                    Processamento
                  </h3>
                  <p className={`${isDarkMode ? 'text-slate-400' : 'text-slate-600'}`}>
                    O sistema processa tudo localmente no seu navegador. Seus dados não são enviados para nenhum servidor, garantindo total privacidade.
                  </p>
                </section>
              </div>

              <button 
                onClick={() => setShowHelpModal(false)}
                className="mt-10 w-full py-5 bg-[#0F172A] text-white rounded-full font-black uppercase tracking-widest hover:bg-slate-800 transition-colors"
              >
                Entendi!
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function FileCard({ title, subtitle, icon, file, onClick, isDarkMode, fullHeight }: { title: string; subtitle: string; icon: React.ReactNode; file: File | null; onClick: () => void; isDarkMode: boolean; fullHeight?: boolean }) {
  return (
    <motion.div
      whileHover={{ y: -4, scale: 1.01 }}
      whileTap={{ scale: 0.98 }}
      onClick={onClick}
      className={`p-8 rounded-[32px] border-2 transition-all cursor-pointer flex flex-col items-center justify-center text-center gap-4 group relative overflow-hidden ${
        fullHeight ? 'h-full min-h-[200px]' : ''
      } ${
        file 
          ? (isDarkMode ? "border-green-500/30 bg-green-500/5" : "border-green-100 bg-green-50/30") 
          : (isDarkMode ? "border-slate-800 bg-slate-900/50 hover:border-amber-500/30" : "border-transparent bg-white shadow-sm hover:shadow-xl hover:shadow-slate-200/50")
      }`}
    >
      {file && (
        <motion.div 
          initial={{ opacity: 0, scale: 0 }}
          animate={{ opacity: 1, scale: 1 }}
          className="absolute top-4 right-4"
        >
          <CheckCircle2 className="w-6 h-6 text-green-500" />
        </motion.div>
      )}
      
      <div className={`w-16 h-16 rounded-2xl flex items-center justify-center transition-all duration-500 ${
        file 
          ? (isDarkMode ? "bg-green-500/20 text-green-400" : "bg-green-100 text-green-600") 
          : (isDarkMode ? "bg-slate-800 text-slate-600 group-hover:bg-amber-500/10 group-hover:text-amber-500" : "bg-slate-50 text-slate-400 group-hover:bg-amber-50 group-hover:text-amber-600")
      }`}>
        {icon}
      </div>
      <div>
        <h4 className={`font-black uppercase tracking-tight ${isDarkMode ? 'text-white' : 'text-[#0F172A]'}`}>{title}</h4>
        <p className={`text-xs mt-1 font-bold truncate max-w-[150px] ${file ? "text-green-500" : "text-slate-400"}`}>
          {file ? file.name : subtitle}
        </p>
      </div>

      {!file && (
        <div className={`mt-2 px-4 py-1.5 rounded-full text-[10px] font-black uppercase tracking-widest transition-colors ${
          isDarkMode ? 'bg-slate-800 text-slate-500 group-hover:bg-amber-500 group-hover:text-white' : 'bg-slate-50 text-slate-400 group-hover:bg-amber-600 group-hover:text-white'
        }`}>
          Selecionar
        </div>
      )}
    </motion.div>
  );
}
