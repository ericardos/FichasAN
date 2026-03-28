import React, { useState, useRef, useEffect } from "react";
import { Upload, FileText, Table, CheckCircle2, AlertCircle, Loader2, Download, Heart, Copy, X, Sparkles } from "lucide-react";
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
  const [copied, setCopied] = useState(false);

  const pixKey = "fdf03993-fbdd-4b89-be41-6e63d2352729";

  const fileInputRefs = {
    capa: useRef<HTMLInputElement>(null),
    ficha: useRef<HTMLInputElement>(null),
    xlsx: useRef<HTMLInputElement>(null),
  };

  const handleFileChange = (type: keyof typeof files) => (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    setFiles((prev) => ({ ...prev, [type]: file }));
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

      // 3. Processar cada aba/turma
      for (const sheetName of sheets) {
        if (sheetName === "PARECERES") continue;

        const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" }) as any[];
        if (rawData.length === 0) continue;

        const safeTurmaName = sheetName.replace(/[^a-zA-Z0-9]/g, "_") || "Turma_Sem_Nome";
        const turmaFolder = resultsZip.folder(safeTurmaName);

        // Capa da Turma
        try {
          const capaBlob = renderDoc(capaBuf, {
            TURMA: sheetName,
            TURNO: String(rawData[0].turno || rawData[0].Turno || "").trim()
          });
          turmaFolder?.file("00_CAPA_DA_TURMA.docx", capaBlob);
        } catch (e: any) {
          console.error(`Erro na capa da turma ${sheetName}:`, e.message);
        }

        // Fichas dos Alunos
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
            const fichaBlob = renderDoc(fichaBuf, {
              NOME: nome,
              PARECER: parecerLongText,
              CONCEITO: conceito,
              TURMA: sheetName,
              TURNO: turno
            });
            
            const safeNome = nome.replace(/[^a-zA-Z0-9]/g, "_") || `Aluno_${i + 1}`;
            turmaFolder?.file(`${safeNome}.docx`, fichaBlob);
            totalFichas++;
          } catch (e: any) {
            console.error(`Erro na ficha do aluno ${nome}:`, e.message);
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
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1C1E] font-sans selection:bg-amber-100">
      {/* Background Pattern */}
      <div className="fixed inset-0 z-0 opacity-[0.03] pointer-events-none" 
           style={{ backgroundImage: 'radial-gradient(#000 1px, transparent 1px)', backgroundSize: '24px 24px' }}></div>

      <div className="relative z-10 max-w-4xl mx-auto px-6 py-12 md:py-20">
        <header className="mb-16 text-center md:text-left flex flex-col md:flex-row md:items-end justify-between gap-6">
          <div>
            <motion.div 
              initial={{ opacity: 0, x: -20 }}
              animate={{ opacity: 1, x: 0 }}
              className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-amber-50 text-amber-700 text-xs font-bold uppercase tracking-wider mb-4 border border-amber-100"
            >
              <Sparkles className="w-3 h-3" />
              Processador Inteligente
            </motion.div>
            <h1 className="text-5xl md:text-6xl font-bold tracking-tight text-[#0F172A] mb-4">
              Fichas<span className="text-amber-600">.</span>io
            </h1>
            <p className="text-lg text-slate-500 max-w-md leading-relaxed">
              Transforme planilhas complexas em documentos Word perfeitamente formatados em segundos.
            </p>
          </div>

          {/* Donation Mini Card */}
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white p-5 rounded-[24px] shadow-sm border border-slate-100 flex items-center gap-4 max-w-sm"
          >
            <div className="w-12 h-12 rounded-2xl bg-amber-50 flex items-center justify-center flex-shrink-0">
              <Heart className="w-6 h-6 text-amber-600 fill-amber-600" />
            </div>
            <div className="text-left">
              <h4 className="font-bold text-sm text-[#0F172A] uppercase tracking-tight">Apoie este projeto</h4>
              <p className="text-xs text-slate-400 mb-2">Este site é gratuito. Considere apoiar!</p>
              <button 
                onClick={() => setShowPixModal(true)}
                className="w-full py-2 bg-[#0F172A] text-white rounded-full text-xs font-bold flex items-center justify-center gap-2 hover:bg-slate-800 transition-colors"
              >
                VER CHAVE PIX
                <Sparkles className="w-3 h-3 text-amber-400" />
              </button>
            </div>
          </motion.div>
        </header>

        <main className="grid gap-8">
          {/* File Selection Section */}
          <section className="grid md:grid-cols-3 gap-4">
            <FileCard
              title="Modelo CAPA"
              subtitle="Documento .docx"
              icon={<FileText className="w-6 h-6" />}
              file={files.capa}
              onClick={() => fileInputRefs.capa.current?.click()}
            />
            <FileCard
              title="Modelo FICHA"
              subtitle="Documento .docx"
              icon={<FileText className="w-6 h-6" />}
              file={files.ficha}
              onClick={() => fileInputRefs.ficha.current?.click()}
            />
            <FileCard
              title="Planilha XLSX"
              subtitle="Dados dos alunos"
              icon={<Table className="w-6 h-6" />}
              file={files.xlsx}
              onClick={() => fileInputRefs.xlsx.current?.click()}
            />
          </section>

          {/* Hidden Inputs */}
          <input type="file" ref={fileInputRefs.capa} onChange={handleFileChange("capa")} accept=".docx" className="hidden" />
          <input type="file" ref={fileInputRefs.ficha} onChange={handleFileChange("ficha")} accept=".docx" className="hidden" />
          <input type="file" ref={fileInputRefs.xlsx} onChange={handleFileChange("xlsx")} accept=".xlsx" className="hidden" />

          {/* Action Card */}
          <div className="bg-white p-10 md:p-16 rounded-[40px] shadow-xl shadow-slate-200/50 border border-slate-100 relative overflow-hidden">
            {/* Decorative elements */}
            <div className="absolute top-0 right-0 w-64 h-64 bg-amber-50 rounded-full -mr-32 -mt-32 blur-3xl opacity-50"></div>
            
            <div className="relative z-10 flex flex-col items-center justify-center text-center">
              <AnimatePresence mode="wait">
                {status === "idle" && (
                  <motion.div
                    key="idle"
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -20 }}
                    className="flex flex-col items-center"
                  >
                    <div className="w-20 h-20 bg-slate-50 rounded-full flex items-center justify-center mb-8">
                      <Upload className="w-8 h-8 text-slate-400" />
                    </div>
                    <h2 className="text-2xl font-bold text-[#0F172A] mb-4">Pronto para começar?</h2>
                    <p className="text-slate-500 mb-10 max-w-xs">Selecione os arquivos acima para habilitar a geração dos documentos.</p>
                    <button
                      onClick={handleGenerate}
                      disabled={!files.capa || !files.ficha || !files.xlsx}
                      className="group relative px-12 py-5 bg-[#0F172A] text-white rounded-full font-bold text-lg transition-all hover:scale-105 hover:shadow-2xl hover:shadow-slate-400 disabled:opacity-20 disabled:hover:scale-100 disabled:hover:shadow-none flex items-center gap-3"
                    >
                      <Sparkles className="w-5 h-5 text-amber-400 group-hover:animate-pulse" />
                      Gerar Documentos
                    </button>
                  </motion.div>
                )}

                {status === "processing" && (
                  <motion.div
                    key="processing"
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="flex flex-col items-center gap-6"
                  >
                    <div className="relative">
                      <div className="w-24 h-24 border-4 border-slate-100 border-t-amber-500 rounded-full animate-spin"></div>
                      <div className="absolute inset-0 flex items-center justify-center">
                        <Loader2 className="w-8 h-8 text-amber-500 animate-pulse" />
                      </div>
                    </div>
                    <div className="text-center">
                      <h3 className="text-xl font-bold text-[#0F172A] mb-2">Processando...</h3>
                      <p className="text-slate-500 animate-pulse">{message}</p>
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
                    <div className="w-24 h-24 bg-green-50 rounded-full flex items-center justify-center mb-8">
                      <CheckCircle2 className="w-12 h-12 text-green-500" />
                    </div>
                    <h3 className="text-3xl font-bold text-[#0F172A] mb-2">Tudo pronto!</h3>
                    <p className="text-slate-500 mb-10">{message}</p>
                    <div className="flex flex-col sm:flex-row gap-4 w-full sm:w-auto">
                      <a
                        href={downloadUrl!}
                        download="fichas_escolares.zip"
                        className="px-10 py-5 bg-green-500 text-white rounded-full font-bold text-lg transition-all hover:scale-105 hover:shadow-xl hover:shadow-green-200 flex items-center justify-center gap-3"
                      >
                        <Download className="w-6 h-6" />
                        Baixar Arquivos
                      </a>
                      <button
                        onClick={() => {
                          setStatus("idle");
                          setFiles({ capa: null, ficha: null, xlsx: null });
                        }}
                        className="px-10 py-5 bg-slate-100 text-slate-600 rounded-full font-bold text-lg transition-all hover:bg-slate-200"
                      >
                        Reiniciar
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
                    <div className="w-24 h-24 bg-red-50 rounded-full flex items-center justify-center mb-8">
                      <AlertCircle className="w-12 h-12 text-red-500" />
                    </div>
                    <h3 className="text-2xl font-bold text-[#0F172A] mb-2">Ops! Algo deu errado</h3>
                    <p className="text-red-500 mb-10 max-w-md">{message}</p>
                    <button
                      onClick={() => setStatus("idle")}
                      className="px-10 py-5 bg-slate-900 text-white rounded-full font-bold transition-all hover:scale-105"
                    >
                      Tentar Novamente
                    </button>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </div>
        </main>

        <footer className="mt-20 pt-10 border-t border-slate-200 flex flex-col md:flex-row items-center justify-between gap-6 text-slate-400">
          <div className="flex items-center gap-2 font-bold text-slate-300">
            <Sparkles className="w-4 h-4" />
            FICHAS.IO
          </div>
          <p className="text-sm">Desenvolvido para facilitar a vida de educadores • 2026</p>
          <button 
            onClick={() => setShowPixModal(true)}
            className="flex items-center gap-2 text-sm font-bold text-amber-600 hover:text-amber-700 transition-colors"
          >
            <Heart className="w-4 h-4 fill-amber-600" />
            Apoie o Projeto
          </button>
        </footer>
      </div>

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
              className="relative bg-white w-full max-w-md rounded-[40px] p-10 text-center shadow-2xl"
            >
              <button 
                onClick={() => setShowPixModal(false)}
                className="absolute top-6 right-6 p-2 rounded-full hover:bg-slate-100 transition-colors"
              >
                <X className="w-6 h-6 text-slate-400" />
              </button>

              <div className="w-16 h-16 bg-amber-50 rounded-2xl flex items-center justify-center mx-auto mb-6">
                <Heart className="w-8 h-8 text-amber-600 fill-amber-600" />
              </div>

              <h2 className="text-2xl font-black text-[#0F172A] uppercase tracking-tight mb-2">Apoio Voluntário</h2>
              <p className="text-slate-500 mb-8">Sua contribuição ajuda a manter o site no ar e gratuito!</p>

              <div className="bg-slate-50 p-8 rounded-[32px] mb-8">
                {/* Placeholder for QR Code */}
                <div className="bg-white p-4 rounded-2xl shadow-sm inline-block mb-6 border border-slate-100">
                  <img 
                    src="https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=fdf03993-fbdd-4b89-be41-6e63d2352729" 
                    alt="PIX QR Code"
                    className="w-[180px] h-[180px]"
                  />
                </div>

                <div className="flex gap-2">
                  <div className="flex-1 bg-white border border-slate-200 rounded-2xl px-4 py-3 text-xs text-slate-400 truncate flex items-center">
                    {pixKey}
                  </div>
                  <button 
                    onClick={copyPixKey}
                    className={`p-4 rounded-2xl transition-all ${copied ? 'bg-green-500 text-white' : 'bg-[#0F172A] text-white hover:bg-slate-800'}`}
                  >
                    {copied ? <CheckCircle2 className="w-5 h-5" /> : <Copy className="w-5 h-5" />}
                  </button>
                </div>
              </div>

              <button 
                onClick={() => setShowPixModal(false)}
                className="text-slate-400 font-bold uppercase tracking-widest text-sm hover:text-slate-600 transition-colors"
              >
                Fechar
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function FileCard({ title, subtitle, icon, file, onClick }: { title: string; subtitle: string; icon: React.ReactNode; file: File | null; onClick: () => void }) {
  return (
    <motion.div
      whileHover={{ y: -4 }}
      onClick={onClick}
      className={`p-8 rounded-[32px] border-2 transition-all cursor-pointer flex flex-col items-center justify-center text-center gap-4 group ${
        file 
          ? "border-green-100 bg-green-50/30" 
          : "border-transparent bg-white shadow-sm hover:shadow-xl hover:shadow-slate-200/50"
      }`}
    >
      <div className={`w-14 h-14 rounded-2xl flex items-center justify-center transition-colors ${
        file ? "bg-green-100 text-green-600" : "bg-slate-50 text-slate-400 group-hover:bg-amber-50 group-hover:text-amber-600"
      }`}>
        {file ? <CheckCircle2 className="w-7 h-7" /> : icon}
      </div>
      <div>
        <h4 className="font-bold text-[#0F172A]">{title}</h4>
        <p className={`text-xs mt-1 font-medium ${file ? "text-green-600" : "text-slate-400"}`}>
          {file ? file.name : subtitle}
        </p>
      </div>
    </motion.div>
  );
}
