import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  FileSpreadsheet, 
  Trash2, 
  Download, 
  Copy, 
  Check, 
  Search, 
  HelpCircle, 
  Moon, 
  Sun, 
  Users, 
  User, 
  FileText, 
  AlertCircle, 
  BookOpen
} from 'lucide-react';
import type { ExcelRow, ConvertedContact } from './types';
import { 
  SCHOOL_UNITS, 
  REQUIRED_COLUMNS 
} from './types';

export default function App() {
  // Theme state
  const [theme, setTheme] = useState<'light' | 'dark'>('light');

  // App core state
  const [selectedUnitCode, setSelectedUnitCode] = useState<string>('2101');
  const [fileName, setFileName] = useState<string>('');
  const [fileSize, setFileSize] = useState<string>('');
  const [fileDate, setFileDate] = useState<string>('');
  const [excelData, setExcelData] = useState<ExcelRow[]>([]);
  const [errorMsg, setErrorMsg] = useState<string>('');
  const [missingCols, setMissingCols] = useState<string[]>([]);
  
  // Filters state
  const [filterPais, setFilterPais] = useState<boolean>(true);
  const [filterMaes, setFilterMaes] = useState<boolean>(true);
  const [filterRL, setFilterRL] = useState<boolean>(true);
  const [filterRF, setFilterRF] = useState<boolean>(true);
  
  // Search and Pagination
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize] = useState<number>(100);
  
  // Modals state
  const [isHelpOpen, setIsHelpOpen] = useState<boolean>(false);
  const [copiedAll, setCopiedAll] = useState<boolean>(false);
  const [copiedRowId, setCopiedRowId] = useState<string | null>(null);
  
  // File input ref
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);

  // Toggle Theme
  const toggleTheme = () => {
    const nextTheme = theme === 'light' ? 'dark' : 'light';
    setTheme(nextTheme);
    document.documentElement.setAttribute('data-theme', nextTheme);
  };

  // VBA Helper - reduzir_campo_telefone (100% validated logic matching MS Access dead-path behavior)
  const reduzirCampoTelefone = (telefones: string): string => {
    if (!telefones) return '';
    let tel = telefones.toString().trim();
    
    const telClean = tel.replace(/\s/g, '').replace(/-/g, '');
    
    // If the number is extremely long (like double country code 55 55...), strip the first 55
    if (telClean.length >= 15 && telClean.startsWith('5555')) {
      const stripped = tel.substring(2).trim();
      if (stripped.startsWith('55')) {
        const remaining = stripped.substring(2).trim();
        if (remaining.startsWith('15')) {
          return '55 15 ' + remaining.substring(2).trim();
        }
      }
      return stripped;
    }
    
    if (telClean.length > 15) {
      tel = tel.substring(2);
    }
    return tel;
  };

  // File drag & drop handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    setErrorMsg('');
    setMissingCols([]);
    
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      processFile(files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setErrorMsg('');
    setMissingCols([]);
    const files = e.target.files;
    if (files && files.length > 0) {
      processFile(files[0]);
    }
  };

  // Parse Excel spreadsheet
  const processFile = (file: File) => {
    if (!file.name.endsWith('.xlsx')) {
      setErrorMsg('Arquivo inválido. Por favor, selecione uma planilha do Excel (.xlsx).');
      return;
    }

    // Set file meta
    setFileName(file.name);
    setFileSize((file.size / 1024).toFixed(1) + ' KB');
    
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) return;
        
        const workbook = XLSX.read(data, { type: 'binary', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        // Convert to JSON
        const rawRows = XLSX.utils.sheet_to_json<any>(sheet, { defval: '' });
        
        if (rawRows.length === 0) {
          setErrorMsg('A planilha selecionada está vazia.');
          return;
        }
        
        // Validate columns
        const headers = Object.keys(rawRows[0]);
        const missing = REQUIRED_COLUMNS.filter(col => !headers.includes(col));
        
        if (missing.length > 0) {
          setMissingCols(missing);
          setErrorMsg('O arquivo selecionado está incompleto ou não é compatível para a geração dos contatos.');
          return;
        }
        
        // Save date metadata (formatted nicely)
        const now = new Date();
        setFileDate(now.toLocaleString('pt-BR'));
        
        // Save valid data
        setExcelData(rawRows);
        setCurrentPage(1);
      } catch (err) {
        console.error(err);
        setErrorMsg('Erro ao ler a planilha. Verifique se o arquivo não está corrompido.');
      }
    };
    reader.readAsBinaryString(file);
  };

  // Process & Convert data based on exact Access VBA logic
  const allConvertedContacts = useMemo((): ConvertedContact[] => {
    if (excelData.length === 0) return [];
    
    const list: ConvertedContact[] = [];
    
    excelData.forEach((row, index) => {
      const serie = (row['Série'] || '').toString().trim();
      const turma = (row['Código da turma'] || '').toString().trim();
      const idEstudante = (row['Identificador Estudante'] || '').toString().trim();
      const nomeCompleto = (row['Nome Completo'] || '').toString().trim();
      
      const pai = (row['Pai'] || '').toString().trim();
      const telPai = (row['Telefone do Pai'] || '').toString().trim();
      
      const mae = (row['Mãe'] || '').toString().trim();
      const telMae = (row['Telefone da Mãe'] || '').toString().trim();
      
      const respLegal = (row['Responsável Legal'] || '').toString().trim();
      const telRL = (row['Telefone do Responsável Legal'] || '').toString().trim();
      
      const respFin = (row['Responsável Financeiro'] || '').toString().trim();
      const telRF = (row['Telefone do Responsável Financeiro'] || '').toString().trim();
      
      // 1. Father Contact
      if (pai && telPai) {
        const isRL = respLegal === pai;
        const isRF = respFin === pai;
        
        let tag = 'P';
        if (isRL) {
          tag = isRF ? 'P - RL - RF' : 'P - RL';
        }
        
        const name = `${serie} (${turma}) - (${tag}) ${pai} - ${selectedUnitCode}/${idEstudante} - (A) ${nomeCompleto}`;
        const phone = reduzirCampoTelefone(telPai);
        
        list.push({
          id: `${index}-P`,
          name,
          phone,
          studentName: nomeCompleto,
          studentId: idEstudante,
          classCode: turma,
          className: serie,
          guardianName: pai,
          guardianType: 'P',
          tags: tag.split(' - ')
        });
      }
      
      // 2. Mother Contact
      if (mae && telMae) {
        const isRL = respLegal === mae;
        const isRF = respFin === mae;
        
        let tag = 'M';
        if (isRL) {
          tag = isRF ? 'M - RL - RF' : 'M - RL';
        }
        
        const name = `${serie} (${turma}) - (${tag}) ${mae} - ${selectedUnitCode}/${idEstudante} - (A) ${nomeCompleto}`;
        const phone = reduzirCampoTelefone(telMae);
        
        list.push({
          id: `${index}-M`,
          name,
          phone,
          studentName: nomeCompleto,
          studentId: idEstudante,
          classCode: turma,
          className: serie,
          guardianName: mae,
          guardianType: 'M',
          tags: tag.split(' - ')
        });
      }
      
      // 3. Responsável Legal Contact (only if NOT father and NOT mother)
      if (respLegal && telRL && respLegal !== pai && respLegal !== mae) {
        const isRF = respLegal === respFin;
        const tag = isRF ? 'RL - RF' : 'RL';
        
        const name = `${serie} (${turma}) - (${tag}) ${respLegal} - ${selectedUnitCode}/${idEstudante} - (A) ${nomeCompleto}`;
        const phone = reduzirCampoTelefone(telRL);
        
        list.push({
          id: `${index}-RL`,
          name,
          phone,
          studentName: nomeCompleto,
          studentId: idEstudante,
          classCode: turma,
          className: serie,
          guardianName: respLegal,
          guardianType: 'RL',
          tags: tag.split(' - ')
        });
      }
      
      // 4. Responsável Financeiro Contact (only if NOT father, mother, and RL)
      if (respFin && telRF && respFin !== pai && respFin !== mae && respFin !== respLegal) {
        const tag = 'RF';
        
        const name = `${serie} (${turma}) - (${tag}) ${respFin} - ${selectedUnitCode}/${idEstudante} - (A) ${nomeCompleto}`;
        const phone = reduzirCampoTelefone(telRF);
        
        list.push({
          id: `${index}-RF`,
          name,
          phone,
          studentName: nomeCompleto,
          studentId: idEstudante,
          classCode: turma,
          className: serie,
          guardianName: respFin,
          guardianType: 'RF',
          tags: [tag]
        });
      }
    });
    
    return list;
  }, [excelData, selectedUnitCode]);

  // Calculate filtered counts of available categories in raw parsed data
  const rawCounts = useMemo(() => {
    let paisCount = 0;
    let maesCount = 0;
    let rlCount = 0;
    let rfCount = 0;

    allConvertedContacts.forEach(contact => {
      if (contact.guardianType === 'P') paisCount++;
      if (contact.guardianType === 'M') maesCount++;
      if (contact.guardianType === 'RL') rlCount++;
      if (contact.guardianType === 'RF') rfCount++;
    });

    return { paisCount, maesCount, rlCount, rfCount };
  }, [allConvertedContacts]);

  // Apply visual category filtering
  const filteredByCategory = useMemo(() => {
    return allConvertedContacts.filter(contact => {
      if (contact.guardianType === 'P') return filterPais;
      if (contact.guardianType === 'M') return filterMaes;
      if (contact.guardianType === 'RL') return filterRL;
      if (contact.guardianType === 'RF') return filterRF;
      return false;
    });
  }, [allConvertedContacts, filterPais, filterMaes, filterRL, filterRF]);

  // Apply search query filtering
  const finalFilteredContacts = useMemo(() => {
    if (!searchTerm.trim()) return filteredByCategory;
    
    const query = searchTerm.toLowerCase();
    return filteredByCategory.filter(c => 
      c.name.toLowerCase().includes(query) || 
      c.phone.toLowerCase().includes(query) || 
      c.studentName.toLowerCase().includes(query) || 
      c.guardianName.toLowerCase().includes(query) ||
      c.studentId.includes(query) ||
      c.classCode.toLowerCase().includes(query)
    );
  }, [filteredByCategory, searchTerm]);

  // Pagination Logic
  const totalPages = Math.ceil(finalFilteredContacts.length / pageSize) || 1;
  const paginatedContacts = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return finalFilteredContacts.slice(start, start + pageSize);
  }, [finalFilteredContacts, currentPage, pageSize]);

  // Copy individual row
  const handleCopyRow = (contact: ConvertedContact) => {
    const text = `${contact.name};${contact.phone}`;
    navigator.clipboard.writeText(text).then(() => {
      setCopiedRowId(contact.id);
      setTimeout(() => setCopiedRowId(null), 2000);
    });
  };

  // Copy all filtered CSV data
  const handleCopyAll = () => {
    if (finalFilteredContacts.length === 0) return;
    
    let csvContent = 'Name;Phone\n';
    finalFilteredContacts.forEach(c => {
      csvContent += `${c.name};${c.phone}\n`;
    });
    
    navigator.clipboard.writeText(csvContent).then(() => {
      setCopiedAll(true);
      setTimeout(() => setCopiedAll(false), 2500);
    });
  };

  // Export CSV matching exact format
  const handleExportCSV = () => {
    if (finalFilteredContacts.length === 0) return;
    
    // DELIMITER is semicolon (;), exactly like the Access app
    let csvContent = 'Name;Phone\r\n';
    finalFilteredContacts.forEach(c => {
      csvContent += `${c.name};${c.phone}\r\n`;
    });
    
    // Encoding: Use UTF-8 with BOM (\uFEFF) so Excel opens accented characters flawlessly on Windows!
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    
    // File nomenclature: Code - date - contatos.csv (ex: 2101 - 19-01-2026 - contatos.csv)
    const rawDate = fileDate.split(' ')[0] || new Date().toLocaleDateString('pt-BR');
    const formattedDate = rawDate.replace(/\//g, '-');
    const filename = `${selectedUnitCode} - ${formattedDate} - contatos.csv`;
    
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Clear application state
  const handleClear = () => {
    setFileName('');
    setFileSize('');
    setFileDate('');
    setExcelData([]);
    setErrorMsg('');
    setMissingCols([]);
    setSearchTerm('');
    setCurrentPage(1);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="app-container">
      {/* 1. Header Section */}
      <header className="app-header glass-card">
        <div className="brand-wrapper">
          <div className="logo-container">
            <BookOpen size={28} className="text-white" />
          </div>
          <div className="brand-title-group">
            <h1>Conversor de Contatos</h1>
            <p>Educação Adventista • Sistema de Nomenclatura 7edu</p>
          </div>
        </div>
        
        <div className="school-selector-container">
          <label htmlFor="school-select" className="text-secondary font-bold text-xs uppercase tracking-wide">
            Unidade:
          </label>
          <select 
            id="school-select"
            className="school-select"
            value={selectedUnitCode}
            onChange={(e) => setSelectedUnitCode(e.target.value)}
          >
            {SCHOOL_UNITS.map(unit => (
              <option key={unit.code} value={unit.code}>
                {unit.name}
              </option>
            ))}
          </select>
        </div>
      </header>

      {/* 2. Drag & Drop Upload Zone */}
      <section className="glass-card">
        {excelData.length === 0 ? (
          <div 
            className={`upload-zone ${isDragging ? 'dragging' : ''}`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
          >
            <input 
              type="file"
              ref={fileInputRef}
              className="file-input"
              accept=".xlsx"
              onChange={handleFileChange}
            />
            <div className="upload-icon-wrapper">
              <Upload size={32} />
            </div>
            <h3>Importar lista de alunos</h3>
            <p>Arraste e solte sua planilha <strong>.xlsx</strong> exportada do 7edu ou clique para navegar</p>
            <span className="text-light text-xs font-semibold mt-2">
              (Apenas processamento local e seguro no navegador)
            </span>
          </div>
        ) : (
          <div className="file-loaded-state">
            <div className="file-loaded-info">
              <FileSpreadsheet size={40} className="file-loaded-icon" />
              <div className="file-loaded-details">
                <h4>{fileName}</h4>
                <div className="flex gap-4 text-xs">
                  <span>Tamanho: {fileSize}</span>
                  <span className="ml-3">Processado em: {fileDate}</span>
                </div>
              </div>
            </div>
            <button className="btn btn-secondary" onClick={handleClear}>
              <Trash2 size={16} /> Substituir
            </button>
          </div>
        )}

        {/* Error State */}
        {errorMsg && (
          <div className="mt-4 p-4 rounded-xl border border-red-500/20 bg-red-500/5 text-red-500 flex gap-3">
            <AlertCircle size={22} className="flex-shrink-0 mt-0.5" />
            <div>
              <p className="font-bold text-sm text-red-600 dark:text-red-400">{errorMsg}</p>
              {missingCols.length > 0 && (
                <div className="mt-2">
                  <p className="text-xs font-semibold text-secondary">Colunas obrigatórias que não foram encontradas:</p>
                  <ul className="list-disc pl-5 mt-1 text-xs text-secondary grid grid-cols-2 gap-x-4">
                    {missingCols.map((c, i) => (
                      <li key={i} className="mt-0.5">{c}</li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          </div>
        )}
      </section>

      {excelData.length > 0 && (
        <>
          {/* 3. Filter Options Card */}
          <section className="glass-card">
            <h3 className="toggle-section-title">Contatos de:</h3>
            <div className="toggles-container">
              <div 
                className={`toggle-card ${filterPais ? 'active' : ''}`}
                onClick={() => setFilterPais(!filterPais)}
              >
                <div className="toggle-label-wrapper">
                  <span className="toggle-title">Pais</span>
                  <span className="toggle-counter">{rawCounts.paisCount} registros</span>
                </div>
                <div className="switch-ui" />
              </div>

              <div 
                className={`toggle-card ${filterMaes ? 'active' : ''}`}
                onClick={() => setFilterMaes(!filterMaes)}
              >
                <div className="toggle-label-wrapper">
                  <span className="toggle-title">Mães</span>
                  <span className="toggle-counter">{rawCounts.maesCount} registros</span>
                </div>
                <div className="switch-ui" />
              </div>

              <div 
                className={`toggle-card ${filterRL ? 'active' : ''}`}
                onClick={() => setFilterRL(!filterRL)}
              >
                <div className="toggle-label-wrapper">
                  <span className="toggle-title">Responsável Legal</span>
                  <span className="toggle-counter">{rawCounts.rlCount} registros</span>
                </div>
                <div className="switch-ui" />
              </div>

              <div 
                className={`toggle-card ${filterRF ? 'active' : ''}`}
                onClick={() => setFilterRF(!filterRF)}
              >
                <div className="toggle-label-wrapper">
                  <span className="toggle-title">Responsável Financeiro</span>
                  <span className="toggle-counter">{rawCounts.rfCount} registros</span>
                </div>
                <div className="switch-ui" />
              </div>
            </div>
          </section>

          {/* 4. Dynamic Dashboard Cards */}
          <section className="metrics-grid">
            <div className="metric-card total">
              <div className="metric-icon-wrapper">
                <Users size={20} />
              </div>
              <div className="metric-details">
                <span className="metric-value">{excelData.length}</span>
                <span className="metric-label">Estudantes</span>
              </div>
            </div>
            
            <div className="metric-card father">
              <div className="metric-icon-wrapper">
                <User size={20} />
              </div>
              <div className="metric-details">
                <span className="metric-value">{filteredByCategory.length}</span>
                <span className="metric-label">Contatos Gerados</span>
              </div>
            </div>

            <div className="metric-card mother">
              <div className="metric-icon-wrapper">
                <FileText size={20} />
              </div>
              <div className="metric-details">
                <span className="metric-value">
                  {searchTerm.trim() ? finalFilteredContacts.length : filteredByCategory.length}
                </span>
                <span className="metric-label">Filtrados / Exibir</span>
              </div>
            </div>

            <div className="metric-card others">
              <div className="metric-icon-wrapper font-bold text-sm">
                CSV
              </div>
              <div className="metric-details">
                <span className="metric-value">{selectedUnitCode}</span>
                <span className="metric-label">Código Escola</span>
              </div>
            </div>
          </section>

          {/* 5. Main Preview Grid & Action Section */}
          <section className="glass-card">
            <div className="action-controls mb-6">
              <div className="search-wrapper">
                <Search size={18} className="search-icon" />
                <input 
                  type="text"
                  placeholder="Pesquisar contatos..."
                  className="search-input"
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value);
                    setCurrentPage(1);
                  }}
                />
              </div>

              <div className="button-group">
                <button 
                  className="btn btn-secondary"
                  onClick={handleCopyAll}
                  disabled={finalFilteredContacts.length === 0}
                >
                  {copiedAll ? <Check size={16} className="text-success" /> : <Copy size={16} />}
                  {copiedAll ? 'Copiado!' : 'Copiar Lista'}
                </button>
                
                <button 
                  className="btn btn-primary"
                  onClick={handleExportCSV}
                  disabled={finalFilteredContacts.length === 0}
                >
                  <Download size={16} /> Exportar csv
                </button>

                <button 
                  className="btn btn-secondary text-danger"
                  onClick={handleClear}
                >
                  Limpar
                </button>
              </div>
            </div>

            {/* Custom Table Grid */}
            {finalFilteredContacts.length > 0 ? (
              <>
                <div className="table-scroll-container">
                  <table className="preview-table">
                    <thead>
                      <tr>
                        <th>Nome do Contato (WhatsApp Nomenclature)</th>
                        <th>Telefone</th>
                        <th style={{ width: '100px', textAlign: 'center' }}>Ações</th>
                      </tr>
                    </thead>
                    <tbody>
                      {paginatedContacts.map((contact) => (
                        <tr key={contact.id}>
                          <td>
                            <div className="font-semibold text-sm">{contact.name}</div>
                            <div className="text-xs text-secondary mt-1 flex items-center gap-2">
                              <span>Aluno: {contact.studentName} ({contact.studentId})</span>
                              <span>•</span>
                              <span>Turma: {contact.classCode}</span>
                              <span>•</span>
                              <span>Parentesco: 
                                <span className="ml-1 font-bold text-primary dark:text-accent">
                                  {contact.guardianType === 'P' ? 'Pai' : 
                                   contact.guardianType === 'M' ? 'Mãe' : 
                                   contact.guardianType === 'RL' ? 'Responsável Legal' : 'Responsável Financeiro'}
                                </span>
                              </span>
                            </div>
                          </td>
                          <td className="font-mono text-sm tracking-wide">
                            {contact.phone}
                          </td>
                          <td style={{ textAlign: 'center' }}>
                            <div className="row-actions justify-center">
                              <button 
                                className={`action-icon-btn ${copiedRowId === contact.id ? 'success-copied' : ''}`}
                                title="Copiar Linha no formato CSV"
                                onClick={() => handleCopyRow(contact)}
                              >
                                {copiedRowId === contact.id ? <Check size={18} /> : <Copy size={18} />}
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Table Pagination Controls */}
                <div className="flex justify-between items-center mt-4 text-xs font-semibold text-secondary flex-wrap gap-2">
                  <span>
                    Exibindo {Math.min(finalFilteredContacts.length, (currentPage - 1) * pageSize + 1)} a{' '}
                    {Math.min(finalFilteredContacts.length, currentPage * pageSize)} de{' '}
                    {finalFilteredContacts.length} registro(s)
                  </span>
                  
                  <div className="flex items-center gap-1">
                    <button 
                      className="btn btn-secondary px-3 py-1"
                      disabled={currentPage === 1}
                      onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                    >
                      Anterior
                    </button>
                    
                    <span className="px-3">
                      Página {currentPage} de {totalPages}
                    </span>
                    
                    <button 
                      className="btn btn-secondary px-3 py-1"
                      disabled={currentPage === totalPages}
                      onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                    >
                      Próxima
                    </button>
                  </div>
                </div>
              </>
            ) : (
              <div className="empty-state">
                <Users size={48} className="empty-state-icon" />
                <h4>Nenhum contato encontrado</h4>
                <p>Verifique os filtros selecionados ou digite outros termos de busca no campo de pesquisa.</p>
              </div>
            )}
          </section>
        </>
      )}

      {/* 6. Footer section */}
      <footer className="app-footer glass-card">
        <span className="footer-credits">
          Conversor de Contatos Web • por <strong>Fabio Alves Feitoza</strong>
        </span>
        
        <div className="flex gap-3">
          <button 
            className="theme-toggle-btn" 
            onClick={toggleTheme} 
            title="Alternar Tema Escuro/Claro"
          >
            {theme === 'light' ? <Moon size={18} /> : <Sun size={18} />}
          </button>
          
          <button 
            className="btn btn-secondary" 
            onClick={() => setIsHelpOpen(true)}
          >
            <HelpCircle size={16} /> Como Importar no Celular
          </button>
        </div>
      </footer>

      {/* 7. Instruction Modal Drawer */}
      {isHelpOpen && (
        <div className="modal-overlay" onClick={() => setIsHelpOpen(false)}>
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>Como importar os contatos no celular?</h3>
              <button className="modal-close-btn" onClick={() => setIsHelpOpen(false)}>
                &times;
              </button>
            </div>
            
            <div className="modal-body">
              <h4>1. Importando no Google Contatos (Recomendado para Android)</h4>
              <p>Ideal para celulares Android e sincronização automática com o WhatsApp.</p>
              <ol className="mt-2">
                <li>Acesse o site <a href="https://contacts.google.com" target="_blank" rel="noopener noreferrer" className="text-primary font-bold hover:underline">Google Contacts</a> no computador logado no mesmo e-mail do celular.</li>
                <li>No menu esquerdo, clique em <strong>Importar</strong>.</li>
                <li>Clique em <strong>Selecionar arquivo</strong> e selecione o arquivo <strong>.csv</strong> exportado por este conversor.</li>
                <li>Clique em <strong>Importar</strong>. Os contatos serão salvos na nuvem e, em minutos, estarão sincronizados no celular e disponíveis no WhatsApp!</li>
              </ol>

              <h4>2. Importando no iCloud (Recomendado para iPhone)</h4>
              <p>Passos para importar e exibir os contatos no iOS/WhatsApp.</p>
              <ol className="mt-2">
                <li>Abra o site <a href="https://www.icloud.com" target="_blank" rel="noopener noreferrer" className="text-primary font-bold hover:underline">iCloud.com</a> e faça login com seu Apple ID.</li>
                <li>Abra o aplicativo <strong>Contatos</strong>.</li>
                <li>Clique no ícone de engrenagem no canto inferior esquerdo e escolha <strong>Importar vCard...</strong> (ou importe diretamente convertendo o CSV para vCard em conversores online).</li>
                <li>Alternativamente, você pode importar o arquivo CSV na conta do Gmail sincronizada no seu iPhone indo em Ajustes &gt; Contatos &gt; Contas.</li>
              </ol>

              <h4>3. Como forçar o WhatsApp a atualizar a lista</h4>
              <ol className="mt-2">
                <li>No WhatsApp do celular, inicie uma nova conversa (ícone de balão de mensagem).</li>
                <li>Toque no menu de três pontos no canto superior direito (Android) ou puxe a lista para baixo para atualizar (iOS).</li>
                <li>Toque em <strong>Atualizar</strong>. Prontinho! Os contatos com a nova nomenclatura aparecerão instantaneamente para busca.</li>
              </ol>
            </div>
            
            <div className="modal-footer">
              <button className="btn btn-primary" onClick={() => setIsHelpOpen(false)}>
                Entendido
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
