import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer
} from 'recharts';
import { Loader2, AlertCircle, Search, Filter, FileText, BarChart3 } from 'lucide-react';

import planilhaPath from './Patentes_IFAL.xlsx';

export default function PatentDashboard() {
  const [rawData, setRawData] = useState([]);
  const [sheetNames, setSheetNames] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Filtros
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
        // Ignora abas de resumo estatístico
        if (sheetName.includes('Quantitativo') || sheetName.includes('Unificado') || sheetName.includes('Sheet')) {
          return;
        }

        const sheet = workbook.Sheets[sheetName];
        // Converte a aba para matriz de linhas
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        let currentType = 'Patente';
        let colIndexes = { projeto: -1, orientador: -1, campus: -1, status: -1 };

        rows.forEach((row, rowIndex) => {
          if (!row || row.length === 0) return;

          // Junta o texto da linha inteira para identificar seções ou cabeçalhos
          const rowText = row.map(cell => String(cell).toLowerCase().trim()).join(' ');

          // 1. Detecta mudança de seção (Tipo de Propriedade Intelectual)
          if (rowText.includes('programa de computador')) {
            currentType = 'Programa de Computador';
            colIndexes = { projeto: -1, orientador: -1, campus: -1, status: -1 }; // Reseta cabeçalhos
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

          // 2. Mapeia a posição real das colunas quando encontra a linha de cabeçalho
          if (rowText.includes('projeto') || rowText.includes('invento') || rowText.includes('programa')) {
            row.forEach((cell, i) => {
              const c = String(cell).toLowerCase().trim();
              if (c.includes('projeto') || c.includes('invento') || c.includes('programa')) colIndexes.projeto = i;
              if (c.includes('orientador') || c.includes('autor')) colIndexes.orientador = i;
              if (c.includes('campus')) colIndexes.campus = i;
              if (c.includes('status')) colIndexes.status = i;
            });
            return; // Pula a linha do próprio cabeçalho
          }

          // 3. Se ainda não achou cabeçalho específico para a seção, usa valores padrão comuns (colunas B, C, D, F)
          const pIdx = colIndexes.projeto !== -1 ? colIndexes.projeto : 1;
          const oIdx = colIndexes.orientador !== -1 ? colIndexes.orientador : 2;
          const cIdx = colIndexes.campus !== -1 ? colIndexes.campus : 3;
          const sIdx = colIndexes.status !== -1 ? colIndexes.status : 5;

          const projeto = String(row[pIdx] || '').trim();
          const orientador = String(row[oIdx] || '').trim();
          const campus = String(row[cIdx] || '').trim();
          const status = String(row[sIdx] || '').trim();

          // 4. Captura o registro se houver um nome de projeto válido
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

  // Aplicação dos Filtros
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

  // Gráfico: Registros por Ano
  const chartDataByYear = useMemo(() => {
    const counts = {};
    rawData.forEach((item) => {
      counts[item.__abaOrigem] = (counts[item.__abaOrigem] || 0) + 1;
    });

    return Object.keys(counts)
      .sort()
      .map((ano) => ({ ano, quantidade: counts[ano] }));
  }, [rawData]);

  // Gráfico: Distribuição por Tipo de PI
  const chartDataByType = useMemo(() => {
    const counts = {};
    filteredData.forEach((item) => {
      counts[item.tipoPI] = (counts[item.tipoPI] || 0) + 1;
    });

    return Object.entries(counts).map(([nome, quantidade]) => ({ nome, quantidade }));
  }, [filteredData]);

  if (loading) return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '100vh', backgroundColor: '#f9fafb' }}>
      <div style={{ textAlign: 'center' }}>
        <Loader2 size={32} style={{ animation: 'spin 1s linear infinite', marginBottom: '12px', color: '#1f2937' }} />
        <p style={{ color: '#6b7280', fontSize: '14px' }}>Processando dados...</p>
      </div>
    </div>
  );
  if (error) return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '100vh', backgroundColor: '#f9fafb' }}>
      <div style={{ textAlign: 'center', padding: '24px' }}>
        <AlertCircle size={32} style={{ color: '#ef4444', marginBottom: '12px' }} />
        <p style={{ color: '#dc2626', fontSize: '14px' }}>Erro: {error}</p>
      </div>
    </div>
  );

  return (
    <div style={{ minHeight: '100vh', backgroundColor: '#f9fafb', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif' }}>
      {/* Header */}
      <header style={{ borderBottom: '1px solid #e5e7eb', backgroundColor: '#fff' }}>
        <div style={{ maxWidth: '1400px', margin: '0 auto', padding: '20px 24px', display: 'flex', alignItems: 'center', gap: '12px' }}>
          <BarChart3 size={24} style={{ color: '#1f2937' }} />
          <h1 style={{ margin: 0, fontSize: '20px', fontWeight: '600', color: '#1f2937' }}>Propriedade Intelectual</h1>
        </div>
      </header>

      <main style={{ maxWidth: '1400px', margin: '0 auto', padding: '24px' }}>
        {/* Filtros */}
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '24px', border: '1px solid #e5e7eb' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '16px' }}>
            <Filter size={18} style={{ color: '#6b7280' }} />
            <h2 style={{ margin: 0, fontSize: '14px', fontWeight: '600', color: '#374151' }}>Filtros</h2>
          </div>
          
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))', gap: '16px' }}>
            <div>
              <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: '500', color: '#374151' }}>Ano/Aba</label>
              <select
                value={selectedSheet}
                onChange={(e) => setSelectedSheet(e.target.value)}
                style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '13px', backgroundColor: '#fff', cursor: 'pointer', whiteSpace: 'nowrap' }}
              >
                <option value="TODAS">Todos os Anos</option>
                {sheetNames.map((name) => (
                  <option key={name} value={name}>{name}</option>
                ))}
              </select>
            </div>

            <div>
              <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: '500', color: '#374151' }}>Tipo de PI</label>
              <select
                value={selectedType}
                onChange={(e) => setSelectedType(e.target.value)}
                style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '13px', backgroundColor: '#fff', cursor: 'pointer', whiteSpace: 'nowrap' }}
              >
                <option value="TODOS">Todos os Tipos</option>
                <option value="Patente">Patente</option>
                <option value="Programa de Computador">Programa de Computador</option>
                <option value="Desenho Industrial">Desenho Industrial</option>
              </select>
            </div>

            <div>
              <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: '500', color: '#374151' }}>Pesquisar</label>
              <div style={{ position: 'relative' }}>
                <input
                  type="text"
                  placeholder="Projeto, autor, campus..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  style={{ width: '100%', padding: '8px 36px 8px 12px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '13px', boxSizing: 'border-box' }}
                />
                <Search size={16} style={{ position: 'absolute', right: '12px', top: '50%', transform: 'translateY(-50%)', color: '#9ca3af', pointerEvents: 'none' }} />
              </div>
            </div>
          </div>
        </div>

        {/* KPI */}
        <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '24px', border: '1px solid #e5e7eb' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <FileText size={18} style={{ color: '#6b7280' }} />
            <div>
              <p style={{ margin: '0 0 4px 0', fontSize: '13px', color: '#6b7280' }}>Total de Registros</p>
              <h3 style={{ margin: 0, fontSize: '28px', fontWeight: '700', color: '#1f2937' }}>{filteredData.length}</h3>
            </div>
          </div>
        </div>

        {/* Gráficos */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '24px', marginBottom: '24px' }}>
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb' }}>
            <h3 style={{ margin: '0 0 16px 0', fontSize: '14px', fontWeight: '600', color: '#1f2937' }}>Evolução por Ano</h3>
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

          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #e5e7eb' }}>
            <h3 style={{ margin: '0 0 16px 0', fontSize: '14px', fontWeight: '600', color: '#1f2937' }}>Distribuição por Tipo</h3>
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

        {/* Tabela */}
        <div style={{ backgroundColor: '#fff', borderRadius: '8px', border: '1px solid #e5e7eb', overflow: 'hidden' }}>
          <div style={{ padding: '20px', borderBottom: '1px solid #e5e7eb' }}>
            <h3 style={{ margin: 0, fontSize: '14px', fontWeight: '600', color: '#1f2937' }}>Projetos ({filteredData.length})</h3>
          </div>
          
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr style={{ backgroundColor: '#f3f4f6', borderBottom: '1px solid #e5e7eb' }}>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontSize: '12px', fontWeight: '600', color: '#374151' }}>Ano</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontSize: '12px', fontWeight: '600', color: '#374151' }}>Tipo</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontSize: '12px', fontWeight: '600', color: '#374151' }}>Projeto</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontSize: '12px', fontWeight: '600', color: '#374151' }}>Autor</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontSize: '12px', fontWeight: '600', color: '#374151' }}>Campus</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontSize: '12px', fontWeight: '600', color: '#374151' }}>Status</th>
                </tr>
              </thead>
              <tbody>
                {filteredData.slice(0, 25).map((item) => (
                  <tr key={item.id} style={{ borderBottom: '1px solid #e5e7eb', transition: 'background-color 0.2s' }} onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#f9fafb'} onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'transparent'}>
                    <td style={{ padding: '12px 16px', fontSize: '13px', color: '#6b7280' }}>{item.__abaOrigem}</td>
                    <td style={{ padding: '12px 16px', fontSize: '13px', fontWeight: '500', color: '#1f2937' }}>{item.tipoPI}</td>
                    <td style={{ padding: '12px 16px', fontSize: '13px', color: '#374151' }}>{item.projeto}</td>
                    <td style={{ padding: '12px 16px', fontSize: '13px', color: '#6b7280' }}>{item.orientador}</td>
                    <td style={{ padding: '12px 16px', fontSize: '13px', color: '#6b7280' }}>{item.campus}</td>
                    <td style={{ padding: '12px 16px', fontSize: '13px', color: '#059669' }}>{item.status}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {filteredData.length > 25 && (
            <div style={{ padding: '16px', backgroundColor: '#f9fafb', borderTop: '1px solid #e5e7eb', fontSize: '13px', color: '#6b7280', textAlign: 'center' }}>
              Exibindo os primeiros 25 de {filteredData.length} resultados
            </div>
          )}
        </div>
      </main>

      <style>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
        
        @media (max-width: 768px) {
          select, input {
            font-size: 16px !important;
          }
        }
      `}</style>
    </div>
  );
}