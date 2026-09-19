const fs = require('fs');
const path = require('path');
const {
  AlignmentType, BorderStyle, Document, Footer, Header, PageNumber,
  Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, WidthType, Packer
} = require('docx');

const base = 'C:\\Users\\Samuel Gomes\\Desktop\\Processos - Decof_LIC\\Contração direta - Serviços de água e esgoto';
const outDir = path.join(base, 'Documentos_da_Instrução');
fs.mkdirSync(outDir, { recursive: true });

const BLUE = '0000FF';
const GRAY = '666666';
const YELLOW = 'FFF2CC';
const border = { style: BorderStyle.SINGLE, size: 4, color: 'A6A6A6' };

function run(value, opts = {}) {
  return new TextRun({ text: value, font: 'Arial', size: 20, ...opts });
}

function paragraph(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [run(children)],
    spacing: { after: 100 },
    ...opts,
  });
}

function cell(value, opts = {}) {
  const children = Array.isArray(value) ? value : [run(value, { bold: !!opts.header, size: opts.size || 17, color: opts.color })];
  return new TableCell({
    width: { size: opts.width || 3000, type: WidthType.DXA },
    shading: opts.shading ? { type: ShadingType.CLEAR, color: opts.shading, fill: opts.shading } : undefined,
    borders: { top: border, bottom: border, left: border, right: border },
    verticalAlign: 'center',
    margins: { top: 50, bottom: 50, left: 70, right: 70 },
    children: [paragraph(children, { spacing: { after: 0 }, alignment: opts.alignment || AlignmentType.LEFT })],
  });
}

function table(rows, widths) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((row, index) => new TableRow({
      children: row.map((entry, j) => cell(entry, {
        width: widths[j],
        header: index === 0,
        shading: index === 0 ? 'D9EAF7' : undefined,
        size: index === 0 ? 17 : 16,
      })),
    })),
  });
}

function headerFooter(label) {
  return {
    headers: { default: new Header({ children: [paragraph([run(label, { bold: true, size: 16, color: GRAY })], { alignment: AlignmentType.RIGHT, spacing: { after: 0 } })] }) },
    footers: { default: new Footer({ children: [paragraph([run('CPII — rascunho para conferência — página ', { size: 16, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GRAY })], { alignment: AlignmentType.CENTER, spacing: { after: 0 } })] }) },
  };
}

const roteiroRows = [
  ['CÓDIGO', 'DOCUMENTO', 'SITUAÇÃO E PROVIDÊNCIA'],
  ['DOC ÁGUA - 01', 'Documento de Formalização da Demanda — DFD', 'Já existe. Conferir se descreve as duas unidades/matrículas e a necessidade contínua.'],
  ['DOC ÁGUA - 02', 'Portaria da Equipe de Planejamento', 'Já existe. Conferir designação e assinaturas.'],
  ['DOC ÁGUA - 03', 'Estudo Técnico Preliminar — ETP', 'Já existe. Conferir aderência com o TR e com a inexigibilidade do art. 74, I.'],
  ['DOC ÁGUA - 04', 'Mapa de Riscos', 'Já existe. Conferir riscos de interrupção, faturamento, reajuste tarifário e fiscalização.'],
  ['DOC ÁGUA - 05', 'Nota Técnica de Justificativa de Preços + faturas/memória de cálculo', 'Área requisitante: preencher os dois campos anuais com R$ 1.180.116,40, retirar a fórmula-modelo em amarelo e assinar. Não há pesquisa competitiva; permanecem a justificativa tarifária e as faturas.'],
  ['DOC ÁGUA - 06', 'Formulário de Enquadramento da Inexigibilidade', 'Licitações, com validação técnica: equivalente ao documento interno citado pela UFRRJ. Juntar AGENERSA e declaração/contrato de concessão vinculando os endereços e matrículas do CPII.'],
  ['DOC ÁGUA - 07', 'Termo de Referência — padrão AGU maio/2026', 'Licitações/EPC: minuta adaptada diretamente do modelo oficial da AGU, com estrutura e cores do padrão. Completar campos azuis, compatibilizar com ETP e colher assinaturas.'],
  ['DOC ÁGUA - 08', 'Habilitação e regularidade da contratada', 'Licitações: consulta SICAF e demais documentos aplicáveis; representação da empresa; CNPJ e certidões que não constem válidas no SICAF.'],
  ['DOC ÁGUA - 09', 'Declaração de disponibilidade orçamentária', 'Área financeira/orçamentária: indicar crédito, fonte, natureza de despesa e disponibilidade para o exercício.'],
  ['DOC ÁGUA - 10', 'Enquadramento no Parecer Referencial AGU 00003/2025 e lista de verificação AGU', 'Licitações: usar somente se a orientação for aceita pelo fluxo jurídico interno do CPII; caso contrário, encaminhar para análise jurídica.'],
  ['DOC ÁGUA - 11', 'Autorização da inexigibilidade', 'Autoridade competente, após a instrução e os controles internos.'],
  ['DOC ÁGUA - 12', 'Instrumento contratual/registro e publicação no PNCP', 'Licitações/contratos: providência posterior à autorização, conforme rito interno e orientação jurídica.'],
];

const roteiro = new Document({
  sections: [{
    properties: { page: { size: { width: 15840, height: 12240, orientation: 'landscape' }, margin: { top: 520, right: 520, bottom: 520, left: 520 } } },
    ...headerFooter('ROTEIRO DE INSTRUÇÃO — ÁGUAS DO RIO'),
    children: [
      paragraph([run('ROTEIRO DE INSTRUÇÃO — INEXIGIBILIDADE ÁGUAS DO RIO / CPII', { bold: true, size: 26 })], { alignment: AlignmentType.CENTER, spacing: { after: 50 } }),
      paragraph([run('Uso interno de conferência. Referências estruturais: processos IFRJ/UFRRJ; modelo AGU de TR (maio/2026).', { italic: true, size: 15, color: GRAY })], { alignment: AlignmentType.CENTER, spacing: { after: 120 } }),
      table(roteiroRows, [1900, 3500, 9300]),
      paragraph([run('Atenção: a AGENERSA confirma São Cristóvão no Bloco 4 da Águas do Rio; para o atesto conclusivo, falta vincular a prova às matrículas/endereço das unidades do CPII.', { bold: true, size: 15, color: '7F6000' })], { shading: { type: ShadingType.CLEAR, color: YELLOW, fill: YELLOW }, spacing: { before: 120, after: 0 } }),
    ],
  }],
});

const atesto = new Document({
  sections: [{
    properties: { page: { margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 } } },
    ...headerFooter('DOC ÁGUA - 06 — ATESTO TÉCNICO'),
    children: [
      paragraph([run('ATESTO TÉCNICO DE EXCLUSIVIDADE E RAZÃO DA ESCOLHA DA CONTRATADA', { bold: true, size: 27 })], { alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
      paragraph([run('Processo administrativo: ', { bold: true }), run('[preencher NUP/SUAP]', { color: BLUE })]),
      paragraph([run('Objeto: ', { bold: true }), run('prestação contínua de serviços públicos de fornecimento de água tratada e coleta/tratamento de esgoto sanitário para as unidades do Colégio Pedro II no Complexo de São Cristóvão.')]),
      paragraph([run('1. BASE DOCUMENTAL', { bold: true, size: 22 })], { spacing: { before: 120, after: 80 } }),
      paragraph([run('1.1. A Nota Técnica de Justificativa de Preços e as faturas que instruem os autos identificam as duas matrículas consumidoras e a estimativa anual da contratação.')]),
      paragraph([run('1.2. O Relatório de Caracterização da AGENERSA relativo ao Bloco 4 informa que o bairro de São Cristóvão, no Município do Rio de Janeiro, integra a área concedida à Águas do Rio 4 SPE S.A. para a prestação regionalizada dos serviços de abastecimento de água e esgotamento sanitário.')]),
      paragraph([run('1.3. Deve ser juntada, antes da assinatura deste atesto, ', { bold: true }), run('declaração atual da Águas do Rio e/ou trecho do contrato de concessão que confirme expressamente que as matrículas/endereço abaixo são atendidos pela concessionária exclusiva local:', { bold: true }), run('\n• Matrícula 1: [preencher número e endereço completo]', { color: BLUE }), run('\n• Matrícula 2: [preencher número e endereço completo]', { color: BLUE })]),
      paragraph([run('2. ANÁLISE', { bold: true, size: 22 })], { spacing: { before: 120, after: 80 } }),
      paragraph([run('2.1. O serviço é essencial ao funcionamento das atividades pedagógicas, administrativas, sanitárias e de saúde da comunidade escolar. A contratação deve abranger exclusivamente as unidades consumidoras identificadas acima.')]),
      paragraph([run('2.2. Não há alternativa competitiva para a prestação do serviço público de água e esgotamento nas unidades atendidas pela rede local. A inviabilidade de competição decorre da concessão territorial, devendo a prova documental mencionada no item 1.3 integrar os autos.')]),
      paragraph([run('2.3. A escolha da Águas do Rio 4 SPE S.A., CNPJ 42.644.220/0001-06, decorre de sua condição de concessionária responsável pela área de atendimento das unidades, e não de comparação de propostas comerciais.')]),
      paragraph([run('3. CONCLUSÃO', { bold: true, size: 22 })], { spacing: { before: 120, after: 80 } }),
      paragraph([run('3.1. Após a juntada e a conferência do documento indicado no item 1.3, atesta-se a inviabilidade de competição para o objeto delimitado, recomendando-se a instrução da contratação direta por inexigibilidade, com fundamento no art. 74, inciso I, da Lei nº 14.133/2021.')]),
      paragraph([run('3.2. Este atesto não substitui a análise dos demais requisitos da contratação direta, especialmente a justificativa de preço, a habilitação da contratada, a disponibilidade orçamentária, o Termo de Referência, a autorização da autoridade competente e o rito jurídico interno.')]),
      paragraph([run('Local e data: ', { bold: true }), run('[preencher]', { color: BLUE })], { spacing: { before: 300 } }),
      paragraph([run('____________________________________________\n[Nome do responsável técnico/requisitante]\n[Cargo — SIAPE — unidade]', { bold: true })], { alignment: AlignmentType.CENTER, spacing: { before: 400, after: 120 } }),
      paragraph([run('Conferência da instrução: ____________________________________________\n[Nome do servidor de Licitações — cargo — SIAPE]', { size: 18 })], { alignment: AlignmentType.CENTER }),
      paragraph([run('RASCUNHO CONDICIONADO — não assinar antes de anexar a declaração da concessionária ou documento de concessão que vincule as matrículas/endereço do CPII.', { bold: true, size: 16, color: '7F6000' })], { shading: { type: ShadingType.CLEAR, color: YELLOW, fill: YELLOW }, spacing: { before: 180, after: 0 } }),
    ],
  }],
});

Packer.toBuffer(roteiro)
  .then(buffer => fs.writeFileSync(path.join(outDir, 'DOC ÁGUA - 00 — Roteiro de Instrução (uma página).docx'), buffer))
  .catch(error => { console.error(error); process.exitCode = 1; });
