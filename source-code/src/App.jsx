import React, { useState, useEffect } from 'react';
import './App.css';
import WelcomeBrand from './components/WelcomeBrand';
import FileUpload from './components/FileUpload';
import ColumnSelector from './components/ColumnSelector';
import FacilityColumnSelector from './components/FacilityColumnSelector';
import MissingMappingsForm from './components/MissingMappingsForm';
import AuditPeriodSelector from './components/AuditPeriodSelector';
import ReportMapper, { TARGET_FIELDS } from './components/ReportMapper';
import { parseExcelHeaders, parseMappingFile } from './utils/excelParser';
import { processAuditData, exportToExcel } from './utils/reportGenerator';
import { Loader2, CheckCircle, XCircle, BarChart3 } from 'lucide-react';

const STEPS = [
  { label: 'Audit Period', state: 'AUDIT_PERIOD' },
  { label: 'Upload File', state: 'UPLOAD' },
  { label: 'Select Columns', state: 'SELECT_COLUMNS' },
  { label: 'Map Fields', state: 'REPORT_MAPPING' },
];

const ACTIVE_STATES = ['AUDIT_PERIOD', 'UPLOAD', 'SELECT_COLUMNS', 'SELECT_FACILITY_COLUMN', 'RESOLVE_MISSING_MAPPINGS', 'REPORT_MAPPING', 'GENERATING_REPORT', 'PARSING', 'PARSING_MAPPING'];

function StepProgress({ currentState }) {
  const getStepStatus = (stepIdx) => {
    const order = ['AUDIT_PERIOD', 'UPLOAD', 'SELECT_COLUMNS', 'REPORT_MAPPING'];
    const stateOrder = {
      'AUDIT_PERIOD': 0,
      'UPLOAD': 1,
      'PARSING': 1,
      'PARSING_MAPPING': 1,
      'SELECT_COLUMNS': 2,
      'SELECT_FACILITY_COLUMN': 2,
      'RESOLVE_MISSING_MAPPINGS': 2,
      'REPORT_MAPPING': 3,
      'GENERATING_REPORT': 3,
    };
    const curr = stateOrder[currentState] ?? 0;
    if (stepIdx < curr) return 'done';
    if (stepIdx === curr) return 'active';
    return 'pending';
  };

  return (
    <div className="step-progress">
      {STEPS.map((step, i) => {
        const status = getStepStatus(i);
        return (
          <React.Fragment key={i}>
            <div className="step-item">
              <div className={`step-circle ${status}`}>
                {status === 'done' ? '✓' : i + 1}
              </div>
              <span className={`step-label ${status}`}>{step.label}</span>
            </div>
            {i < STEPS.length - 1 && (
              <div className={`step-connector ${status === 'done' ? 'done' : ''}`} />
            )}
          </React.Fragment>
        );
      })}
    </div>
  );
}

function Navbar() {
  return (
    <nav className="app-navbar">
      <div className="app-navbar-brand">
        <div className="app-navbar-logo">JS</div>
        <div>
          <div className="app-navbar-title">Bank Audit Intelligence Suite</div>
        </div>
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        <BarChart3 size={15} color="var(--text-muted)" />
        <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)', fontWeight: 500 }}>CA Jainam Shah</span>
      </div>
    </nav>
  );
}

function LoadingCard({ message, sub }) {
  return (
    <div className="card card-body loading-card animate-fade">
      <div className="loading-icon-wrap">
        <Loader2 size={32} color="var(--brand-primary)" className="spin" />
      </div>
      <div className="loading-title">{message}</div>
      {sub && <p className="loading-sub">{sub}</p>}
    </div>
  );
}

function App() {
  const [appState, setAppState] = useState('WELCOME');
  const [auditPeriod, setAuditPeriod] = useState(null);
  const [npaConfig, setNpaConfig] = useState({ available: false, source: '' });
  const [fileDetails, setFileDetails] = useState({ file: null, headers: [], rawData: [], selectedColumns: [] });
  const [npaFileDetails, setNpaFileDetails] = useState({ file: null, headers: [], rawData: [], selectedColumns: [] });
  const [mappingDetails, setMappingDetails] = useState({ dictionary: {}, fileMappings: {}, facilityColumn: '', missingCodes: [], mappingFileProvided: false });
  const [error, setError] = useState(null);
  const [successMsg, setSuccessMsg] = useState(null);

  // Load saved product mappings on mount
  useEffect(() => {
    const saved = localStorage.getItem('dataSummarizer_productMappings');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        setMappingDetails(prev => ({ ...prev, dictionary: { ...prev.dictionary, ...parsed } }));
      } catch (e) {
        console.error("Failed to load saved mappings:", e);
      }
    }
  }, []);

  const showError = (msg) => {
    setError(msg);
    setTimeout(() => setError(null), 5000);
  };

  const showSuccess = (msg) => {
    setSuccessMsg(msg);
    setTimeout(() => setSuccessMsg(null), 5000);
  };

  const handleAuditPeriodSubmit = (periodData) => {
    const { fromDate, toDate, npaAvailable, npaSource } = periodData;
    setAuditPeriod({ fromDate, toDate });
    setNpaConfig({ available: npaAvailable, source: npaSource });
    setAppState('UPLOAD');
  };

  const handleFileUpload = async (dataFile, mappingFile, npaFile) => {
    setAppState('PARSING');
    setError(null);
    try {
      const { headers, rawData } = await parseExcelHeaders(dataFile);
      setFileDetails({ file: dataFile, headers, rawData, selectedColumns: [] });

      if (npaFile) {
        const npaResult = await parseExcelHeaders(npaFile);
        setNpaFileDetails({ file: npaFile, headers: npaResult.headers, rawData: npaResult.rawData, selectedColumns: [] });
      }

      if (mappingFile) {
        setAppState('PARSING_MAPPING');
        const fileDictionary = await parseMappingFile(mappingFile);

        // Update both dictionary (for runtime) and fileMappings (to track source)
        setMappingDetails(prev => {
          const updatedDict = { ...prev.dictionary, ...fileDictionary };
          return { ...prev, dictionary: updatedDict, fileMappings: fileDictionary, mappingFileProvided: true };
        });
      } else {
        setMappingDetails(prev => ({ ...prev, fileMappings: {}, mappingFileProvided: false }));
      }
      setAppState('SELECT_COLUMNS');
    } catch (err) {
      console.error(err);
      showError('Failed to parse Excel file. Please ensure it is a valid .xlsx or .xls format.');
      setAppState('UPLOAD');
    }
  };

  const handleColumnsSelected = (selectedColumns) => {
    setFileDetails(prev => ({ ...prev, selectedColumns }));
    setAppState(mappingDetails.mappingFileProvided ? 'SELECT_FACILITY_COLUMN' : 'REPORT_MAPPING');
  };

  const handleFacilityColumnSelected = (columnName) => {
    const uniqueFacilityCodes = new Set();
    fileDetails.rawData.forEach(row => {
      const code = String(row[columnName] || '').trim();
      if (code) uniqueFacilityCodes.add(code);
    });

    // Only show codes that are NOT in the uploaded mapping file
    const codesMissingFromFile = [...uniqueFacilityCodes].filter(code => !mappingDetails.fileMappings[code]);

    if (codesMissingFromFile.length > 0) {
      setMappingDetails(prev => ({ ...prev, facilityColumn: columnName, missingCodes: codesMissingFromFile }));
      setAppState('RESOLVE_MISSING_MAPPINGS');
    } else {
      setMappingDetails(prev => ({ ...prev, facilityColumn: columnName, missingCodes: [] }));
      finalizePhase2(mappingDetails.dictionary, columnName);
    }
  };

  const handleMissingMappingsResolved = (manualMappings) => {
    const updatedDictionary = { ...mappingDetails.dictionary, ...manualMappings };

    // Persist to localStorage
    localStorage.setItem('dataSummarizer_productMappings', JSON.stringify(updatedDictionary));

    setMappingDetails(prev => ({ ...prev, dictionary: updatedDictionary }));
    finalizePhase2(updatedDictionary, mappingDetails.facilityColumn);
  };

  const finalizePhase2 = (dictionary, facilityColumn) => {
    const updatedData = fileDetails.rawData.map(row => {
      const code = String(row[facilityColumn] || '').trim();
      return { ...row, 'Mapped Loan Type': dictionary[code] || 'Unmapped' };
    });
    setFileDetails(prev => ({ ...prev, rawData: updatedData }));
    setAppState('REPORT_MAPPING');
  };

  const handleGenerateAndDownload = async (fieldMappings, npaMappings) => {
    setAppState('GENERATING_REPORT');
    setTimeout(async () => {
      try {
        // If separate NPA file, create a map for easy lookup
        let externalNpaData = null;
        if (npaConfig.available && npaConfig.source === 'separate' && npaFileDetails.file) {
          externalNpaData = {};
          const joinKey = npaMappings.joinCol;
          const valKey = npaMappings.provisionCol;
          npaFileDetails.rawData.forEach(row => {
            const id = String(row[joinKey] || '').trim();
            if (id) externalNpaData[id] = row[valKey];
          });
        }

        const processedData = processAuditData(fileDetails.rawData, fieldMappings, auditPeriod, TARGET_FIELDS, externalNpaData, npaMappings.joinCol);
        // Guard: if no data rows, show error instead of exporting a blank file
        if (!processedData || processedData.length === 0) {
          showError('No data to export. Please check your column mappings — at least one row must be present.');
          setAppState('REPORT_MAPPING');
          return;
        }
        await exportToExcel(processedData, TARGET_FIELDS, auditPeriod, npaConfig);
        showSuccess('✓ Excel report downloaded successfully!');
        setAppState('REPORT_MAPPING');
      } catch (e) {
        console.error(e);
        showError('Failed to generate or export the report. Please check your data.');
        setAppState('REPORT_MAPPING');
      }
    }, 100);
  };

  // Welcome screen full-page
  if (appState === 'WELCOME') return <WelcomeBrand onStart={() => setAppState('AUDIT_PERIOD')} />;

  const showProgress = ACTIVE_STATES.includes(appState);

  return (
    <div className="app-wrapper">
      <Navbar />
      <main className="app-main">
        {showProgress && <StepProgress currentState={appState} />}

        {error && (
          <div className="error-banner animate-fade">
            <XCircle size={16} />
            {error}
          </div>
        )}
        {successMsg && (
          <div className="alert alert-success animate-fade" style={{ marginBottom: '1.5rem' }}>
            <CheckCircle size={16} style={{ flexShrink: 0 }} />
            {successMsg}
          </div>
        )}

        {appState === 'AUDIT_PERIOD' && <AuditPeriodSelector onProceed={handleAuditPeriodSubmit} />}
        {appState === 'UPLOAD' && <FileUpload onFileUpload={handleFileUpload} npaConfig={npaConfig} />}
        {appState === 'PARSING' && <LoadingCard message="Reading file…" sub="Parsing your Excel sheet contents." />}
        {appState === 'PARSING_MAPPING' && <LoadingCard message="Loading product mappings…" sub="Parsing the Product Mapping file." />}
        {appState === 'SELECT_COLUMNS' && (
          <ColumnSelector
            headers={fileDetails.headers}
            onProceed={handleColumnsSelected}
            onBack={() => setAppState('UPLOAD')}
          />
        )}
        {appState === 'SELECT_FACILITY_COLUMN' && (
          <FacilityColumnSelector
            availableColumns={fileDetails.selectedColumns}
            onSelect={handleFacilityColumnSelected}
            onBack={() => setAppState('SELECT_COLUMNS')}
          />
        )}
        {appState === 'RESOLVE_MISSING_MAPPINGS' && (
          <MissingMappingsForm
            missingCodes={mappingDetails.missingCodes}
            existingMappings={mappingDetails.dictionary}
            onComplete={handleMissingMappingsResolved}
            onBack={() => setAppState('SELECT_FACILITY_COLUMN')}
          />
        )}
        {appState === 'REPORT_MAPPING' && (
          <ReportMapper
            availableColumns={fileDetails.selectedColumns}
            onGenerate={handleGenerateAndDownload}
            onBack={() => setAppState(mappingDetails.mappingFileProvided ? 'SELECT_FACILITY_COLUMN' : 'SELECT_COLUMNS')}
            npaConfig={npaConfig}
            npaFileDetails={npaFileDetails}
          />
        )}
        {appState === 'GENERATING_REPORT' && (
          <LoadingCard
            message="Building your report…"
            sub="Calculating fields, formulas, and structuring your audit workbook."
          />
        )}
      </main>
    </div>
  );
}

export default App;
