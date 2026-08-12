import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer
} from 'recharts';
import { Loader2, AlertCircle, Search, Filter, FileText, BarChart3 } from 'lucide-react';
import './App.css';

import planilhaPath from './Patentes_IFAL.xlsx';

export default function PatentDashboard() {
  const [rawData, setRawData] = useState([]);
  const [sheetNames, setSheetNames] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [selectedSheet, setSelectedSheet] = useState('TODAS');
  const [selectedType, setSelectedType] = useState('TODOS');
  const [searchTerm, setSearchTerm] = useState('');

  useEffect(() => {
  async function loadExcelFile() {
    try {
      setLoading(true);
      const response = await fetch(planilhaPath);
      
      if (!response.ok) {
        throw new Error('Não foi possível carregar a planilha Patentes_IFAL.xlsx');
      }

      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });

      setSheetNames(workbook.SheetNames);

      let consolidated = [];

      workbook.SheetNames.forEach((sheetName) => {
        if (sheetName.includes('Quantitativo') || sheetName.includes('Unificado') || sheetName.includes('Sheet')) {
          return;
        }

        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        let currentType = 'Patente';
        let colIndexes = { projeto: -1, orientador: -1, campus: -1, status: -1 };

        rows.forEach((row, rowIndex) => {
          if (!row || row.length === 0) return;

          const rowText = row.map(cell => String(cell).toLowerCase().trim()).join(' ');

          if (rowText.includes('programa de computador')) {
            currentType = 'Programa de Computador';
            colIndexes = { projeto: -1, orientador: -1, campus: -1, status: -1 };
            return;
          }
          if (rowText.includes('desenho industrial')) {
            currentType = 'Desenho Industrial';
            colIndexes = { projeto: -1, orientador: -1, campus: -1, status: -1 };
            return;
          }
          if (rowText.includes('pedidos de patentes')) {
            currentType = 'Patente';
            colIndexes = { projeto: -1, orientador: -1, campus: -1, status: -1 };
            return;
          }

          if (rowText.includes('projeto') || rowText.includes('invento') || rowText.includes('programa')) {
            row.forEach((cell, i) => {
              const c = String(cell).toLowerCase().trim();
              if (c.includes('projeto') || c.includes('invento') || c.includes('programa')) colIndexes.projeto = i;
              if (c.includes('orientador') || c.includes('autor')) colIndexes.orientador = i;
              if (c.includes('campus')) colIndexes.campus = i;
              if (c.includes('status')) colIndexes.status = i;
            });
            return;
          }

          const pIdx = colIndexes.projeto !== -1 ? colIndexes.projeto : 1;
          const oIdx = colIndexes.orientador !== -1 ? colIndexes.orientador : 2;
          const cIdx = colIndexes.campus !== -1 ? colIndexes.campus : 3;
          const sIdx = colIndexes.status !== -1 ? colIndexes.status : 5;

          const projeto = String(row[pIdx] || '').trim();
          const orientador = String(row[oIdx] || '').trim();
          const campus = String(row[cIdx] || '').trim();
          const status = String(row[sIdx] || '').trim();

          if (
            projeto && 
            projeto !== '' && 
            !projeto.toLowerCase().includes('solicitações') &&
            !projeto.toLowerCase().includes('pedidos de') &&
            !projeto.toLowerCase().includes('registros de')
          ) {
            consolidated.push({
              id: `${sheetName}-${rowIndex}`,
              tipoPI: currentType,
              projeto: projeto,
              orientador: orientador || 'Não Informado',
              campus: campus || 'Não Informado',
              status: status || 'Acompanhamento',
              __abaOrigem: sheetName,
            });
          }
        });
      });

      setRawData(consolidated);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  loadExcelFile();
}, []);

  const filteredData = useMemo(() => {
    return rawData.filter((item) => {
      const matchesSheet = selectedSheet === 'TODAS' || item.__abaOrigem === selectedSheet;
      const matchesType = selectedType === 'TODOS' || item.tipoPI === selectedType;
      
      const term = searchTerm.toLowerCase();
      const matchesSearch =
        item.projeto.toLowerCase().includes(term) ||
        item.orientador.toLowerCase().includes(term) ||
        item.campus.toLowerCase().includes(term) ||
        item.status.toLowerCase().includes(term);

      return matchesSheet && matchesType && matchesSearch;
    });
  }, [rawData, selectedSheet, selectedType, searchTerm]);

  const chartDataByYear = useMemo(() => {
    const counts = {};
    rawData.forEach((item) => {
      counts[item.__abaOrigem] = (counts[item.__abaOrigem] || 0) + 1;
    });

    return Object.keys(counts)
      .sort()
      .map((ano) => ({ ano, quantidade: counts[ano] }));
  }, [rawData]);

  const chartDataByType = useMemo(() => {
    const counts = {};
    filteredData.forEach((item) => {
      counts[item.tipoPI] = (counts[item.tipoPI] || 0) + 1;
    });

    return Object.entries(counts).map(([nome, quantidade]) => ({ nome, quantidade }));
  }, [filteredData]);

  if (loading) return (
    <div className="loading-container">
      <div className="loading-content">
        <Loader2 size={32} className="loading-icon" />
        <p className="loading-text">Processando dados...</p>
      </div>
    </div>
  );
  if (error) return (
    <div className="error-container">
      <div className="error-content">
        <AlertCircle size={32} className="error-icon" />
        <p className="error-text">Erro: {error}</p>
      </div>
    </div>
  );

  return (
    <div className="app">
      <header className="header">
        <div className="header-content">
          <BarChart3 size={24} className="header-icon" />
          <h1 className="header-title">Propriedade Intelectual</h1>
        </div>
      </header>

      <main className="main">
        <div className="card filters-section">
          <div className="filters-header">
            <Filter size={18} className="filters-icon" />
            <h2 className="filters-title">Filtros</h2>
          </div>
          
          <div className="filters-grid">
            <div className="filter-group">
              <label className="filter-label">Ano/Aba</label>
              <select
                value={selectedSheet}
                onChange={(e) => setSelectedSheet(e.target.value)}
                className="filter-select"
              >
                <option value="TODAS">Todos os Anos</option>
                {sheetNames.map((name) => (
                  <option key={name} value={name}>{name}</option>
                ))}
              </select>
            </div>

            <div className="filter-group">
              <label className="filter-label">Tipo de PI</label>
              <select
                value={selectedType}
                onChange={(e) => setSelectedType(e.target.value)}
                className="filter-select"
              >
                <option value="TODOS">Todos os Tipos</option>
                <option value="Patente">Patente</option>
                <option value="Programa de Computador">Programa de Computador</option>
                <option value="Desenho Industrial">Desenho Industrial</option>
              </select>
            </div>

            <div className="filter-group">
              <label className="filter-label">Pesquisar</label>
              <div className="search-wrapper">
                <input
                  type="text"
                  placeholder="Projeto, autor, campus..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="search-input"
                />
                <Search size={16} className="search-icon" />
              </div>
            </div>
          </div>
        </div>

        <div className="card kpi-card">
          <div className="kpi-content">
            <FileText size={18} className="kpi-icon" />
            <div className="kpi-info">
              <p className="kpi-label">Total de Registros</p>
              <h3 className="kpi-value">{filteredData.length}</h3>
            </div>
          </div>
        </div>

        <div className="charts-grid">
          <div className="card chart-card">
            <h3 className="chart-title">Evolução por Ano</h3>
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={chartDataByYear}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                <XAxis dataKey="ano" stroke="#9ca3af" style={{ fontSize: '12px' }} />
                <YAxis stroke="#9ca3af" style={{ fontSize: '12px' }} />
                <Tooltip contentStyle={{ borderRadius: '6px', border: '1px solid #e5e7eb' }} />
                <Line type="monotone" dataKey="quantidade" stroke="#1f2937" strokeWidth={2} dot={{ fill: '#1f2937', r: 4 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>

          <div className="card chart-card">
            <h3 className="chart-title">Distribuição por Tipo</h3>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={chartDataByType}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" />
                <XAxis dataKey="nome" stroke="#9ca3af" style={{ fontSize: '12px' }} />
                <YAxis stroke="#9ca3af" style={{ fontSize: '12px' }} />
                <Tooltip contentStyle={{ borderRadius: '6px', border: '1px solid #e5e7eb' }} />
                <Bar dataKey="quantidade" fill="#6366f1" radius={[6, 6, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="table-container">
          <div className="table-header">
            <h3 className="table-title">Projetos ({filteredData.length})</h3>
          </div>
          
          <div className="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>Ano</th>
                  <th>Tipo</th>
                  <th>Projeto</th>
                  <th>Autor</th>
                  <th>Campus</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {filteredData.slice(0, 25).map((item) => (
                  <tr key={item.id}>
                    <td>{item.__abaOrigem}</td>
                    <td className="type-cell">{item.tipoPI}</td>
                    <td className="project-cell">{item.projeto}</td>
                    <td>{item.orientador}</td>
                    <td>{item.campus}</td>
                    <td className="status-cell">{item.status}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {filteredData.length > 25 && (
            <div className="table-footer">
              Exibindo os primeiros 25 de {filteredData.length} resultados
            </div>
          )}
        </div>
      </main>
    </div>
  );
}