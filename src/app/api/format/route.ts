import { NextRequest, NextResponse } from 'next/server';
import * as xlsx from 'xlsx';
import { compareTwoStrings } from 'string-similarity';

export async function POST(req: NextRequest) {
  try {
    const data = await req.formData();
    const file: File | null = data.get('file') as unknown as File;

    if (!file) {
      return NextResponse.json({ success: false, error: 'Nenhum arquivo enviado.' }, { status: 400 });
    }

    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);

    // Read the original file
    const workbook = xlsx.read(buffer, { type: 'buffer' });

    // Extract data from all sheets and concatenate them
    let rows: any[][] = [];
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const sheetRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1 });
      if (sheetRows.length > 0) {
        rows = rows.concat(sheetRows);
      }
    }

    if (rows.length === 0) {
      return NextResponse.json({ success: false, error: 'A planilha está vazia.' }, { status: 400 });
    }

    // Heuristics to find columns in the original file
    let headerRowIndex = -1;
    let colIndices = {
      nome: -1,
      escola: -1,
      turma: -1,
      turno: -1,
      ano: -1,
    };

    // Normalization helper
    const normalizeString = (str: string) => {
      return str
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // remove acentos
        .replace(/\s+/g, ' ') // remove espaços extras
        .toLowerCase()
        .trim();
    };

    const synonymsNome = ['nome', 'aluno', 'estudante', 'candidato', 'criança', 'crianca'];
    const synonymsEscola = ['escola', 'instituição', 'unidade', 'colegio', 'centro'];
    const synonymsTurma = ['turma', 'classe', 'sala', 'agrupamento'];
    const synonymsTurno = ['turno', 'periodo', 'horario'];
    const synonymsAno = ['ano', 'serie', 'nivel', 'etapa'];

    for (let i = 0; i < Math.min(50, rows.length); i++) {
      const row = rows[i];
      if (!row || !Array.isArray(row)) continue;

      let foundNome = -1;
      let foundEscola = -1;
      let foundTurma = -1;
      let foundTurno = -1;
      let foundAno = -1;

      for (let j = 0; j < row.length; j++) {
        const normalizedCell = normalizeString(String(row[j] || ''));
        if (!normalizedCell) continue;
        
        // Não confundir "Nome Completo" com "Nome da Escola", ou "Nome da Mãe/Pai"
        if (synonymsNome.some(s => normalizedCell.includes(s)) && !synonymsEscola.some(s => normalizedCell.includes(s)) && !normalizedCell.includes('mae') && !normalizedCell.includes('pai') && !normalizedCell.includes('responsavel')) {
          foundNome = j;
        } else if (synonymsEscola.some(s => normalizedCell.includes(s)) && !normalizedCell.includes('ano escolar')) {
          // Ignora 'ano escolar' para não sobrescrever a coluna de escola
          foundEscola = j;
        } else if (synonymsTurma.some(s => normalizedCell.includes(s))) {
          foundTurma = j;
        } else if (synonymsTurno.some(s => normalizedCell.includes(s))) {
          foundTurno = j;
        } else if (synonymsAno.some(s => normalizedCell.includes(s)) && !normalizedCell.includes('letivo') && !normalizedCell.includes('nascimento')) {
          // Ignora 'ano letivo' (ex: 2026) e 'data de nascimento'
          foundAno = j;
        }
      }

      if (foundNome !== -1 && (foundEscola !== -1 || foundTurma !== -1 || foundAno !== -1)) {
        headerRowIndex = i;
        colIndices = {
          nome: foundNome,
          escola: foundEscola,
          turma: foundTurma,
          turno: foundTurno,
          ano: foundAno,
        };
        break;
      }
    }

    if (headerRowIndex === -1) {
      return NextResponse.json({ success: false, error: 'Não foi possível encontrar o cabeçalho com colunas como Nome, Turma, etc.' }, { status: 400 });
    }

    // Build the new data matrix grouped by school
    const groupedRows: Record<string, any[][]> = {};
    const knownSchools: { original: string; core: string }[] = [];
    
    // Função para extrair apenas o "nome núcleo" da escola, ignorando palavras genéricas
    const getCoreSchoolName = (name: string) => {
      let core = normalizeString(name);
      // Remover pontuações
      core = core.replace(/[.,-]/g, ' ');
      // Remover termos genéricos que inflam a similaridade
      const genericTerms = [
        'escola', 'municipal', 'estadual', 'centro', 'educacional', 'colegio', 'instituto',
        'creche', 'infantil', 'emef', 'cmei', 'cei', 'e m', 'e e', 'c e', 'c m', 'em', 'ee', 'ce', 'cm',
        'de', 'da', 'do', 'das', 'dos', 'e'
      ];
      
      const words = core.split(/\s+/);
      const filteredWords = words.filter(w => !genericTerms.includes(w));
      return filteredWords.join(' ').trim();
    };
    
    // Iterate through data rows in the original file
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !Array.isArray(row) || row.length === 0) continue;

      // Extract raw values and remove leading/trailing spaces
      const valNome = colIndices.nome !== -1 && row[colIndices.nome] ? String(row[colIndices.nome]).trim() : '';
      const valEscola = colIndices.escola !== -1 && row[colIndices.escola] ? String(row[colIndices.escola]).trim() : '';
      const valTurma = colIndices.turma !== -1 && row[colIndices.turma] ? String(row[colIndices.turma]).trim() : '';
      const valTurno = colIndices.turno !== -1 && row[colIndices.turno] ? String(row[colIndices.turno]).trim() : '';
      let valAno = colIndices.ano !== -1 && row[colIndices.ano] ? String(row[colIndices.ano]).trim() : '';

      // Fallback: Se a escola não mandou coluna "Ano", tentar extrair da "Turma" (ex: "2º Ano A" -> "2")
      if (!valAno && valTurma) {
         const anoMatch = valTurma.match(/\d+/);
         if (anoMatch) {
             valAno = anoMatch[0];
         }
      }

      // Skip if completely empty row
      if (!valNome && !valEscola && !valTurma && !valTurno && !valAno) {
        continue;
      }

      // Filter out repeated headers mixed in the data
      if (
        valTurma.toLowerCase() === 'turma' || 
        valEscola.toLowerCase().includes('nome da escola') || 
        valNome.toLowerCase() === 'nome completo' ||
        valNome.toLowerCase() === 'nome'
      ) {
        continue;
      }

      const turmaUpper = valTurma.toUpperCase();
      const anoUpper = valAno.toUpperCase();

      // Ignorar Creche, Infantil e outras nomenclaturas sem ano numerado
      if (
        turmaUpper.includes('CRECHE') || anoUpper.includes('CRECHE') ||
        turmaUpper.includes('INFANTIL') || anoUpper.includes('INFANTIL') ||
        turmaUpper.includes('MATERNAL') || anoUpper.includes('MATERNAL') ||
        turmaUpper.includes('BERÇARIO') || anoUpper.includes('BERÇARIO') ||
        turmaUpper.includes('BERCARIO') || anoUpper.includes('BERCARIO')
      ) {
        continue;
      }

      // Se a turma/ano não contém nenhum dígito numérico (ex: "1", "2", "3"), ignora os alunos
      const hasYearNumber = /\d/.test(valAno) || /\d/.test(valTurma);
      if (!hasYearNumber) {
        continue;
      }

      // Se a linha não tiver o nome do aluno, ignora (evita linhas semi-vazias)
      if (!valNome) {
        continue;
      }

      // Format Ano to always include "° ANO" if needed
      if (valAno) {
        // Ensure it has the degree symbol
        if (!valAno.includes('º') && !valAno.includes('°')) {
          const match = valAno.match(/\d+/);
          if (match) {
            valAno = valAno.replace(match[0], match[0] + 'º');
          }
        }
        
        // Ensure it has " ANO" (case-insensitive check)
        if (!valAno.toUpperCase().includes('ANO')) {
          valAno = valAno + ' ANO';
        }
      }

      let canonicalSchoolName = 'Sem_Escola';
      if (valEscola) {
        const coreName = getCoreSchoolName(valEscola);
        let bestMatch = { index: -1, score: 0 };
        
        for (let j = 0; j < knownSchools.length; j++) {
          const score = compareTwoStrings(coreName, knownSchools[j].core);
          if (score > bestMatch.score) {
            bestMatch = { index: j, score };
          }
        }

        if (bestMatch.score >= 0.70) {
          // Utiliza o primeiro nome oficial encontrado que tem >= 70% de similaridade
          canonicalSchoolName = knownSchools[bestMatch.index].original;
        } else {
          // É uma escola nova
          knownSchools.push({ original: valEscola, core: coreName });
          canonicalSchoolName = valEscola;
        }
      }
      
      // Target template:
      // A: Nome
      // B: (blank)
      // C: Escola
      // D: Turma
      // E: Turno
      // F: Ano
      
      const newRow = [
        valNome || '',
        null, // Coluna B em branco
        canonicalSchoolName, // Usamos o nome padronizado
        valTurma || '',
        valTurno || '',
        valAno || ''
      ];

      if (!groupedRows[canonicalSchoolName]) {
        // Rule 1: First 3 rows must be blank for each new school sheet
        groupedRows[canonicalSchoolName] = [[], [], []];
      }
      
      groupedRows[canonicalSchoolName].push(newRow);
    }

    const processedFiles: { nome: string; conteudoBase64: string }[] = [];
    const originalBaseName = file.name.replace(/\.[^/.]+$/, ""); // strip extension

    for (const [schoolName, schoolRows] of Object.entries(groupedRows)) {
      // Create the new workbook and worksheet
      const newWb = xlsx.utils.book_new();
      const newWs = xlsx.utils.aoa_to_sheet(schoolRows);
      xlsx.utils.book_append_sheet(newWb, newWs, 'Planilha Formatada');

      // Write to buffer
      const outBuffer = xlsx.write(newWb, { type: 'buffer', bookType: 'xlsx' });
      
      // Convert to Base64
      const base64Content = outBuffer.toString('base64');
      
      // Safe filename: remove invalid characters if any
      const safeSchoolName = schoolName.replace(/[\\/:*?"<>|]/g, '-').trim();
      
      let finalFilename = `Formatado_${originalBaseName}.xlsx`;
      // If there's more than one school in this sheet, append the school name to the file name
      if (Object.keys(groupedRows).length > 1) {
        finalFilename = `Formatado_${originalBaseName}_${safeSchoolName}.xlsx`;
      }
      
      processedFiles.push({
        nome: finalFilename,
        conteudoBase64: base64Content
      });
    }

    // Return JSON array of formatted files
    return NextResponse.json({ success: true, files: processedFiles }, { status: 200 });

  } catch (error: any) {
    console.error('Error processing file:', error);
    return NextResponse.json({ success: false, error: 'Erro interno ao processar o arquivo: ' + error.message }, { status: 500 });
  }
}
