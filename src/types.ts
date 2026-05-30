export interface ExcelRow {
  'Código da turma'?: string;
  'Série'?: string;
  'Identificador Estudante'?: string;
  'Nome Completo'?: string;
  'Pai'?: string;
  'Telefone do Pai'?: string;
  'Mãe'?: string;
  'Telefone da Mãe'?: string;
  'Responsável Legal'?: string;
  'Telefone do Responsável Legal'?: string;
  'Responsável Financeiro'?: string;
  'Telefone do Responsável Financeiro'?: string;
}

export interface ConvertedContact {
  id: string; // unique identification
  name: string; // full generated nomenclature
  phone: string; // phone number
  studentName: string;
  studentId: string;
  classCode: string;
  className: string;
  guardianName: string;
  guardianType: 'P' | 'M' | 'RL' | 'RF';
  tags: string[]; // e.g. ['P', 'RL', 'RF']
}

export interface SchoolUnit {
  code: string;
  name: string;
}

export const SCHOOL_UNITS: SchoolUnit[] = [
  { code: '401', name: '401 - Escola Adventista de Americanópolis' },
  { code: '501', name: '501 - Escola Adventista de Araçariguama' },
  { code: '601', name: '601 - Colégio Adventista do Brooklin' },
  { code: '901', name: '901 - Escola Adventista do Campo Grande' },
  { code: '1001', name: '1001 - Escola Adventista de Cidade Ademar' },
  { code: '1201', name: '1201 - Colégio Adventista da Liberdade' },
  { code: '1401', name: '1401 - Colégio Adventista de Cotia' },
  { code: '1901', name: '1901 - Colégio Adventista de Granja Viana' },
  { code: '2101', name: '2101 - Escola Adventista de Ibiúna' },
  { code: '2201', name: '2201 - Colégio Adventista de Interlagos' },
  { code: '2601', name: '2601 - Colégio Adventista da Lapa' },
  { code: '2901', name: '2901 - Escola Adventista de Pedreira' },
  { code: '3501', name: '3501 - Colégio Adventista de Santo Amaro' },
  { code: '4101', name: '4101 - Colégio Adventista de São Roque' },
  { code: '1015', name: '1015 - Colégio Adventista de Vila Yara' }
];

export const REQUIRED_COLUMNS = [
  'Código da turma',
  'Série',
  'Identificador Estudante',
  'Nome Completo',
  'Pai',
  'Telefone do Pai',
  'Mãe',
  'Telefone da Mãe',
  'Responsável Legal',
  'Telefone do Responsável Legal',
  'Responsável Financeiro',
  'Telefone do Responsável Financeiro'
];
