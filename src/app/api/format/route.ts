import { NextRequest, NextResponse } from 'next/server';
import * as xlsx from 'xlsx';

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
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert to array of arrays
    const rows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1 });

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

    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i];
      if (!row || !Array.isArray(row)) continue;

      let foundNome = -1;
      let foundEscola = -1;
      let foundTurma = -1;
      let foundTurno = -1;
      let foundAno = -1;

      for (let j = 0; j < row.length; j++) {
        const cellValue = String(row[j] || '').toLowerCase().trim();
        // Não confundir "Nome Completo" com "Nome da Escola"
        if (cellValue.includes('nome') && !cellValue.includes('escola')) foundNome = j;
        else if (cellValue === 'aluno' || cellValue === 'alunos') foundNome = j;
        
        if (cellValue.includes('escola')) foundEscola = j;
        if (cellValue.includes('turma')) foundTurma = j;
        if (cellValue.includes('turno')) foundTurno = j;
        if (cellValue.includes('ano') || cellValue.includes('série') || cellValue.includes('serie')) foundAno = j;
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

    // Build the new data matrix
    const newRows: any[][] = [];
    
    // Rule 1: First 3 rows must be blank
    newRows.push([]);
    newRows.push([]);
    newRows.push([]);

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
        valEscola || '',
        valTurma || '',
        valTurno || '',
        valAno || ''
      ];

      newRows.push(newRow);
    }

    // Create the new workbook and worksheet
    const newWb = xlsx.utils.book_new();
    const newWs = xlsx.utils.aoa_to_sheet(newRows);
    xlsx.utils.book_append_sheet(newWb, newWs, 'Planilha Formatada');

    // Write to buffer
    const outBuffer = xlsx.write(newWb, { type: 'buffer', bookType: 'xlsx' });

    // Return the formatted file as response
    return new NextResponse(outBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="Formatado.xlsx"',
      },
    });

  } catch (error: any) {
    console.error('Error processing file:', error);
    return NextResponse.json({ success: false, error: 'Erro interno ao processar o arquivo: ' + error.message }, { status: 500 });
  }
}
