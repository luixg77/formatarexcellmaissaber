'use client';

import { useState, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, Loader2, Download, AlertCircle } from 'lucide-react';

export default function Home() {
  const [files, setFiles] = useState<File[]>([]);
  const [formattedFiles, setFormattedFiles] = useState<{ nome: string; conteudoBase64: string }[] | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    setError(null);
    setFormattedFiles(null);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFiles = Array.from(e.dataTransfer.files).filter(
        f => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
      );
      
      if (droppedFiles.length > 0) {
        setFiles(prev => [...prev, ...droppedFiles]);
      } else {
        setError('Por favor, selecione apenas arquivos Excel (.xlsx ou .xls).');
      }
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    setFormattedFiles(null);
    if (e.target.files && e.target.files.length > 0) {
      const selectedFiles = Array.from(e.target.files).filter(
        f => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
      );
      setFiles(prev => [...prev, ...selectedFiles]);
    }
  };

  const removeFile = (e: React.MouseEvent, index: number) => {
    e.stopPropagation();
    setFormattedFiles(null);
    setFiles(prev => prev.filter((_, i) => i !== index));
  };

  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  const handleFormat = async () => {
    if (files.length === 0) return;

    setIsProcessing(true);
    setError(null);
    setFormattedFiles(null);

    try {
      const allProcessed: { nome: string; conteudoBase64: string }[] = [];

      for (const fileToProcess of files) {
        const formData = new FormData();
        formData.append('file', fileToProcess);

        const response = await fetch('/api/format', {
          method: 'POST',
          body: formData,
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(`Erro ao processar ${fileToProcess.name}: ${errorData.error}`);
        }

        const responseData = await response.json();
        
        if (!responseData.success || !responseData.files) {
          throw new Error('Erro ao processar o arquivo. Resposta inválida do servidor.');
        }

        allProcessed.push(...responseData.files);
      }

      setFormattedFiles(allProcessed);
    } catch (err: any) {
      setError(err.message || 'Erro inesperado.');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownload = async () => {
    if (!formattedFiles) return;

    if (formattedFiles.length >= 7) {
      // Create ZIP
      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();

      for (const fileObj of formattedFiles) {
        const byteCharacters = atob(fileObj.conteudoBase64);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        
        zip.file(fileObj.nome, byteArray);
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = window.URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `planilhas_formatadas_${new Date().getTime()}.zip`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } else {
      for (const fileObj of formattedFiles) {
        // Decode Base64 string to Blob
        const byteCharacters = atob(fileObj.conteudoBase64);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileObj.nome;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        // Maior delay entre downloads para o navegador não bloquear múltiplos arquivos da mesma planilha
        await new Promise(r => setTimeout(r, 800));
      }
    }

    // Reset after downloading
    setFiles([]);
    setFormattedFiles(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <main className="min-h-screen bg-gray-50 flex flex-col items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-lg w-full p-8">
        <div className="text-center mb-8">
          <div className="bg-blue-100 text-blue-600 w-16 h-16 rounded-full flex items-center justify-center mx-auto mb-4">
            <FileSpreadsheet size={32} />
          </div>
          <h1 className="text-2xl font-bold text-gray-800">Formatador de Planilhas</h1>
          <p className="text-gray-500 mt-2 text-sm">
            Importe a planilha da escola e ela será convertida automaticamente para o template de importação.
          </p>
        </div>

        <div
          className={`border-2 border-dashed rounded-xl p-10 text-center cursor-pointer transition-colors ${
            isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-blue-400 hover:bg-gray-50'
          }`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={handleUploadClick}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept=".xlsx, .xls"
            multiple
            className="hidden"
          />
          
          {files.length > 0 ? (
            <div className="flex flex-col items-center">
              <FileSpreadsheet size={48} className="text-green-500 mb-3" />
              <p className="font-medium text-gray-700">{files.length} arquivo(s) selecionado(s)</p>
              
              <div className="w-full mt-4 space-y-2 text-left">
                {files.map((f, i) => (
                  <div key={i} className="flex justify-between items-center bg-white border rounded p-2 text-sm text-gray-600">
                    <span className="truncate max-w-[250px]">{f.name}</span>
                    <button 
                      onClick={(e) => removeFile(e, i)}
                      className="text-red-500 hover:text-red-700 font-bold px-2"
                    >
                      X
                    </button>
                  </div>
                ))}
              </div>
              <p className="text-xs text-gray-400 mt-4">Clique para adicionar mais arquivos</p>
            </div>
          ) : (
            <div className="flex flex-col items-center">
              <UploadCloud size={48} className="text-gray-400 mb-3" />
              <p className="font-medium text-gray-600">Arraste a planilha aqui</p>
              <p className="text-sm text-gray-400 mt-1">ou clique para selecionar do computador</p>
            </div>
          )}
        </div>

        {error && (
          <div className="mt-4 p-4 bg-red-50 border border-red-200 rounded-lg flex items-start text-red-600">
            <AlertCircle size={20} className="mr-2 flex-shrink-0 mt-0.5" />
            <p className="text-sm">{error}</p>
          </div>
        )}

        <div className="mt-8">
          {!formattedFiles ? (
            <button
              onClick={handleFormat}
              disabled={files.length === 0 || isProcessing}
              className={`w-full py-3 px-4 rounded-xl font-medium text-white flex items-center justify-center transition-all duration-200 active:scale-95 ${
                files.length === 0 || isProcessing
                  ? 'bg-blue-300 cursor-not-allowed'
                  : 'bg-blue-600 hover:bg-blue-700 shadow-md hover:shadow-lg hover:-translate-y-1'
              }`}
            >
              {isProcessing ? (
                <>
                  <Loader2 size={20} className="animate-spin mr-2" />
                  Processando arquivo...
                </>
              ) : (
                <>
                  <FileSpreadsheet size={20} className="mr-2" />
                  Formatar Planilha
                </>
              )}
            </button>
          ) : (
            <button
              onClick={handleDownload}
              className="w-full py-3 px-4 rounded-xl font-medium text-white flex items-center justify-center transition-all duration-200 active:scale-95 bg-green-600 hover:bg-green-700 shadow-md hover:shadow-lg hover:-translate-y-1"
            >
              <Download size={20} className="mr-2" />
              Baixar {formattedFiles.length > 1 ? 'Planilhas Prontas' : 'Planilha Pronta'}
            </button>
          )}
        </div>

        <div className="mt-6 border-t pt-4">
          <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Regras Aplicadas:</h3>
          <ul className="text-xs text-gray-500 space-y-1 list-disc list-inside">
            <li>As 3 primeiras linhas ficarão em branco.</li>
            <li>Colunas Mapeadas: Nome (A), Escola (C), Turma (D), Turno (E), Ano (F).</li>
            <li>A Coluna B (Matrícula) ficará vazia.</li>
            <li>Formatação automática do Ano (ex: 2 para 2º).</li>
          </ul>
        </div>
      </div>

      <footer className="mt-8 text-gray-400 text-sm flex items-center justify-center gap-2">
        <span>luizteste82@gmail.com</span>
        <span className="text-gray-300">|</span>
        <span className="text-gray-400 text-xs font-mono">v1.0.4</span>
      </footer>
    </main>
  );
}
