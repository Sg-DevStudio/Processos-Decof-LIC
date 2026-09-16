const fs = require('fs');
const path = require('path');
const {
  AlignmentType, BorderStyle, Document, Footer, Header, HeadingLevel,
  PageNumber, Paragraph, ShadingType, Table, TableCell, TableRow,
  TextRun, WidthType, Packer
} = require('docx');

const base = 'C:\\Users\\Samuel Gomes\\Desktop\\Processos - Decof_LIC\\Contração direta - Serviços de água e esgoto';
const out = path.join(base, 'Termo_de_Referencia', 'Rascunho_para_Conferencia', 'TR_Aguas_do_Rio_CPII_RASCUNHO.docx');

const BLUE = '0000FF';
const RED = 'C00000';
const GRAY = '666666';
const border = { style: BorderStyle.SINGLE, size: 4, color: 'A6A6A6' };

function text(value, options = {}) {
  return new TextRun({ text: value, font: 'Arial', size: 22, ...options });
}

function p(children, options = {}) {
  return new Paragraph({ children: Array.isArray(children) ? children : [text(children)], spacing: { after: 120 }, ...options });
}

function title(value) {
  return p([text(value, { bold: true, size: 30 })], { alignment: AlignmentType.CENTER, spacing: { after: 180 } });
}

function h(value, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ text: value, heading: level, spacing: { before: 220, after: 120 } });
}

function cell(value, opts = {}) {
  return new TableCell({
    width: { size: opts.width || 4500, type: WidthType.DXA },
    shading: opts.header ? { type: ShadingType.CLEAR, color: 'D9EAF7', fill: 'D9EAF7' } : undefined,
    borders: { top: border, bottom: border, left: border, right: border },
    children: [p([text(value, { bold: !!opts.header, size: 19, color: opts.color })], { spacing: { after: 0 } })]
  });
}

function twoColumn(rows) {
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [2800, 6200],
    rows: rows.map(([a, b], i) => new TableRow({ children: [cell(a, { header: i === 0, width: 2800 }), cell(b, { header: i === 0, width: 6200, color: b.startsWith('[') ? BLUE : undefined })] }))
  });
}

function note(value) {
  return p([text(value, { italic: true, color: GRAY, size: 18 })], { shading: { type: ShadingType.CLEAR, color: 'FFF2CC', fill: 'FFF2CC' }, spacing: { after: 140 } });
}

const body = [
  title('TERMO DE REFERÊNCIA - RASCUNHO PARA CONFERÊNCIA'),
  p([text('Contratação direta por inexigibilidade - fornecimento de água tratada e esgotamento sanitário', { bold: true, color: RED })], { alignment: AlignmentType.CENTER }),
  note('RASCUNHO. Não assinar enquanto houver textos em azul ou decisões marcadas para validação. Estrutura baseada no Modelo de Termo de Referência - Serviços e Obras - Lei nº 14.133/2021, AGU, maio de 2026.'),
  twoColumn([
    ['Informação', 'Conteúdo'],
    ['Processo administrativo', '[preencher NUP/SUAP/SEI]'],
    ['Unidade contratante', 'Colégio Pedro II - UASG 153167'],
    ['Modalidade/fundamento', 'Inexigibilidade de licitação - art. 74, I, da Lei nº 14.133/2021'],
    ['Contratada', 'Águas do Rio 4 SPE S.A. - CNPJ 42.644.220/0001-06'],
    ['Unidades consumidoras', 'Complexo de São Cristóvão - duas matrículas da concessionária; identificar números, endereços e economias no Anexo I'],
    ['Estimativa anual', 'R$ 1.180.116,40 (um milhão, cento e oitenta mil, cento e dezesseis reais e quarenta centavos)'],
  ]),

  h('1. CONDIÇÕES GERAIS DA CONTRATAÇÃO'),
  p([text('1.1. ') , text('Objeto. ', { bold: true }), text('Contratação da Águas do Rio 4 SPE S.A. para a prestação contínua de serviços públicos de fornecimento de água tratada e coleta/tratamento de esgoto sanitário para as unidades do Colégio Pedro II integrantes do Complexo de São Cristóvão, nas respectivas matrículas e categorias tarifárias, conforme Anexo I.')]),
  p([text('1.2. ') , text('Natureza. ', { bold: true }), text('Trata-se de serviço público essencial, de natureza continuada, sem dedicação exclusiva de mão de obra, cuja interrupção compromete as atividades educacionais, administrativas, sanitárias e de saúde da comunidade escolar.')]),
  p([text('1.3. ') , text('Fundamento da contratação direta. ', { bold: true }), text('A contratação será realizada por inexigibilidade de licitação, com fundamento no art. 74, I, da Lei nº 14.133/2021, em razão da inviabilidade de competição decorrente da atuação exclusiva da concessionária local.')]),
  p([text('1.4. ') , text('Vigência. ', { bold: true }), text('Propõe-se a vigência por prazo indeterminado, na forma do art. 109 da Lei nº 14.133/2021, condicionada à confirmação anual da disponibilidade de créditos orçamentários e da estimativa de consumo. ', { }), text('[Validar com a área jurídica/administrativa do CPII antes da assinatura.]', { color: BLUE })]),

  h('2. FUNDAMENTAÇÃO E DESCRIÇÃO DA NECESSIDADE DA CONTRATAÇÃO'),
  p([text('2.1. O Complexo de São Cristóvão necessita de abastecimento contínuo de água potável e de coleta/tratamento de esgoto para assegurar higiene, salubridade, funcionamento regular das atividades pedagógicas e administrativas e prevenção de riscos à saúde pública.')]),
  p([text('2.2. A estimativa foi formada a partir do histórico de faturamento de agosto de 2025 a julho de 2026, reunindo doze meses de medições efetivas das duas matrículas, com atualização das faturas anteriores ao reajuste tarifário de dezembro de 2025, conforme Nota Técnica de Justificativa de Preços e faturas anexadas aos autos.')]),
  p([text('2.3. A AGENERSA identifica São Cristóvão como bairro do Município do Rio de Janeiro inserido no Bloco 4, concedido à Águas do Rio 4 para prestação regionalizada dos serviços de abastecimento de água e esgotamento sanitário. A área técnica deve juntar, ainda, declaração atual da concessionária e/ou o contrato de concessão que confirme as matrículas e endereços concretos do CPII.')]),

  h('3. REQUISITOS DA CONTRATAÇÃO'),
  p([text('3.1. A contratada deverá manter a prestação dos serviços segundo as normas regulatórias aplicáveis, as categorias tarifárias e as condições do contrato de concessão.')]),
  p([text('3.2. Cada fatura deverá identificar, por matrícula: período de competência, leitura/consumo medido, tarifas aplicadas, valor de água, valor de esgoto, recursos hídricos ou tributos incidentes, eventuais ajustes e valor total.')]),
  p([text('3.3. A contratada deverá comunicar interrupções programadas, emergências e providências de reparo pelos canais oficiais, observando os padrões regulatórios aplicáveis.')]),
  p([text('3.4. Não se aplica subcontratação, pois a execução decorre da concessão e da atuação exclusiva da concessionária na área atendida.')]),
  p([text('3.5. Não será exigida garantia contratual, salvo decisão fundamentada da autoridade competente. ', { }), text('[Confirmar a opção no processo.]', { color: BLUE })]),

  h('4. MODELO DE EXECUÇÃO DO OBJETO'),
  p([text('4.1. Os serviços serão prestados de forma contínua nas unidades consumidoras indicadas no Anexo I, mediante disponibilidade da rede, medição regular do consumo e faturamento mensal pela contratada.')]),
  p([text('4.2. A execução compreende o fornecimento de água tratada, a coleta e o tratamento de esgoto sanitário, nos limites da infraestrutura disponível e conforme as regras da concessão.')]),
  p([text('4.3. A contratada deverá disponibilizar faturas e registros de consumo que permitam a conferência mensal pelo fiscal técnico/requisitante.')]),

  h('5. MODELO DE GESTÃO DO CONTRATO'),
  p([text('5.1. A execução será acompanhada e fiscalizada por servidores formalmente designados pelo CPII, nos termos do art. 117 da Lei nº 14.133/2021.')]),
  p([text('5.2. Gestor do contrato: ', { bold: true }), text('[nome, SIAPE e unidade]', { color: BLUE })]),
  p([text('5.3. Fiscal técnico/requisitante: ', { bold: true }), text('[nome, SIAPE e unidade]', { color: BLUE })]),
  p([text('5.4. Fiscal substituto: ', { bold: true }), text('[nome, SIAPE e unidade]', { color: BLUE })]),
  p([text('5.5. Compete à fiscalização conferir as faturas, registrar ocorrências, atestar a execução correspondente ao consumo medido e comunicar irregularidades à contratada e à gestão.')]),

  h('6. CRITÉRIOS DE MEDIÇÃO E PAGAMENTO'),
  p([text('6.1. A medição será mensal, com base no consumo efetivamente registrado nas matrículas constantes do Anexo I e nas tarifas públicas/regulatórias vigentes.')]),
  p([text('6.2. O pagamento ocorrerá após o ateste da fatura pelo fiscal competente, no prazo e pelas regras de pagamento aplicáveis ao CPII, observado o valor efetivamente devido no mês.')]),
  p([text('6.3. A estimativa anual não representa consumo mínimo garantido nem autoriza pagamento por valor fixo; é teto estimativo para a contratação.')]),
  p([text('6.4. Reajustes tarifários homologados pela agência reguladora serão aplicados à faturação, devendo ser comprovados e conferidos nos autos.')]),

  h('7. FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR'),
  p([text('7.1. A seleção será por contratação direta, mediante inexigibilidade de licitação, em favor da Águas do Rio 4 SPE S.A., por ser a concessionária exclusiva do serviço público na área de São Cristóvão, condicionada à comprovação documental da exclusividade para as unidades do CPII.')]),
  p([text('7.2. Antes da contratação, deverão ser juntados os documentos de habilitação e regularidade exigíveis, bem como a comprovação de representação da empresa e os documentos de exclusividade previstos no art. 74, § 1º, da Lei nº 14.133/2021.')]),

  h('8. ESTIMATIVAS DO VALOR DA CONTRATAÇÃO'),
  p([text('8.1. O valor anual estimado é de R$ 1.180.116,40 (um milhão, cento e oitenta mil, cento e dezesseis reais e quarenta centavos), apurado na Nota Técnica de Justificativa de Preços a partir das faturas do período de agosto de 2025 a julho de 2026, corrigidas quando cabível.')]),
  p([text('8.2. A justificativa de preço decorre da aplicação das tarifas públicas homologadas e da inexistência de concorrência, devendo permanecer juntadas a Nota Técnica, as faturas de suporte e a regulamentação/tarifa aplicável.')]),

  h('9. ADEQUAÇÃO ORÇAMENTÁRIA'),
  p([text('9.1. As despesas correrão à conta dos créditos orçamentários próprios do CPII. ', { }), text('[Preencher programa de trabalho, natureza da despesa, fonte, plano interno e declaração de disponibilidade orçamentária.]', { color: BLUE })]),
  p([text('9.2. Para a vigência por prazo indeterminado, a existência de créditos orçamentários deverá ser comprovada em cada exercício financeiro, nos termos do art. 109 da Lei nº 14.133/2021.')]),

  h('10. ANEXOS E DOCUMENTOS DE SUPORTE'),
  p([text('10.1. Integram ou devem instruir os autos:')]),
  p([text('a) DFD, ETP e Mapa de Riscos aprovados;')]),
  p([text('b) Nota Técnica de Justificativa de Preços assinada;')]),
  p([text('c) Faturas das duas matrículas e memória de cálculo;')]),
  p([text('d) documento de exclusividade: AGENERSA/contrato de concessão e declaração atual da Águas do Rio que vincule os endereços/matrículas do CPII;')]),
  p([text('e) documentos de habilitação/regularidade da contratada;')]),
  p([text('f) declaração de disponibilidade orçamentária;')]),
  p([text('g) razão da escolha da contratada, justificativa do preço, parecer jurídico/referencial aplicável, autorização da autoridade competente e checklist da contratação direta.')]),

  h('11. RESPONSÁVEIS'),
  p([text('Elaborado por: ', { bold: true }), text('[nome, cargo, SIAPE e unidade]', { color: BLUE })]),
  p([text('Integrante requisitante/técnico: ', { bold: true }), text('[nome, cargo, SIAPE e unidade]', { color: BLUE })]),
  p([text('Aprovado por: ', { bold: true }), text('[autoridade competente]', { color: BLUE })]),
  note('Conferência obrigatória: preencher todos os campos azuis; confirmar a vigência indeterminada, as duas matrículas/endereços, a dotação, os fiscais e a correspondência integral com o ETP e o Mapa de Riscos já elaborados. O documento deve passar pelo fluxo e pelas assinaturas do CPII antes de qualquer anexação ao sistema.'),
];

const doc = new Document({
  sections: [{
    properties: { page: { margin: { top: 1100, right: 1100, bottom: 1100, left: 1100 } } },
    headers: { default: new Header({ children: [p([text('COLÉGIO PEDRO II - TERMO DE REFERÊNCIA', { bold: true, size: 18, color: GRAY })], { alignment: AlignmentType.RIGHT, spacing: { after: 0 } })] }) },
    footers: { default: new Footer({ children: [p([text('Rascunho para conferência - página ', { size: 18, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GRAY })], { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 } })] }) },
    children: body
  }]
});

Packer.toBuffer(doc).then(buffer => fs.writeFileSync(out, buffer));
