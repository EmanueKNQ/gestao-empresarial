import React, { useEffect, useState, useMemo } from 'react';
import {
  Building2, Landmark, CalendarClock, HandCoins, Plus, X, Trash2, Pencil,
  LayoutDashboard, TrendingDown, AlertCircle, CheckCircle2, FileBarChart,
  ChevronRight, ArrowLeft, Download, ArrowLeftRight, ChevronDown, ChevronUp,
  CreditCard, Receipt, Percent, RefreshCw, Shield, ShoppingCart,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, Cell } from 'recharts';
import * as db from '../lib/data';

const BANDEIRAS = ['Visa', 'Mastercard', 'Elo', 'American Express', 'Hipercard', 'Outra'];
const PARCELAS_RANGE = Array.from({ length: 11 }, (_, i) => i + 2);
const LINHAS_CREDITO = ['PRONAMP', 'FNE', 'FGI', 'Recurso Próprio', 'Parcel. Imp. Fed. PERT III', 'Capital de giro', 'Conta garantida', 'Financiamento', 'Cheque especial', 'Antecipação de recebíveis', 'Desconto de duplicatas', 'Outra'];
const FORMAS_PAGAMENTO_SEGURO = [
  { value: 'boleto', label: 'Boleto' },
  { value: 'debito', label: 'Débito em conta' },
  { value: 'cartao', label: 'Cartão' },
];
const SETORES_VENDA = [
  { value: 'loja', label: 'Loja' },
  { value: 'filial', label: 'Filial' },
  { value: 'maquinas', label: 'Máquinas' },
  { value: 'assistencia_tecnica', label: 'Assistência Técnica' },
  { value: 'reparticoes', label: 'Repartições' },
  { value: 'locacoes', label: 'Locações' },
];
const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const ANO_INICIAL_VENDAS = 2018;
function anosDisponiveis() {
  const atual = new Date().getFullYear();
  const arr = [];
  for (let y = atual; y >= ANO_INICIAL_VENDAS; y--) arr.push(y);
  return arr;
}
const PDF_NAVY = [31, 58, 95];
const PDF_TEXT = [27, 39, 51];
const PDF_MUTED_BG = [241, 244, 247];

const fmtBRL = (v) => (Number(v) || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const fmtPct = (v) => (v === null || v === undefined || v === '') ? '—' : `${Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
const fmtDate = (iso) => { if (!iso) return '—'; const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };
const todayISO = () => new Date().toISOString().slice(0, 10);
const daysUntil = (iso) => Math.round((new Date(iso) - new Date(todayISO())) / 86400000);
function addMonthsLocal(iso, n) {
  const [y, m, d] = iso.split('-').map(Number);
  const date = new Date(y, (m - 1) + n, d);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

// ---------- Reusable UI ----------

function Field({ label, children }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
      <style jsx>{`
        .field { display:flex; flex-direction:column; gap:6px; }
        span { font-size:12px; letter-spacing:.04em; text-transform:uppercase; color:var(--text-muted); }
      `}</style>
    </label>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div className="backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="head"><h3>{title}</h3><button className="icon-btn" onClick={onClose}><X size={18} /></button></div>
        <div className="body">{children}</div>
      </div>
      <style jsx>{`
        .backdrop { position:fixed; inset:0; background:rgba(8,10,14,0.72); display:flex; align-items:center; justify-content:center; z-index:50; padding:16px; }
        .modal { background:var(--surface); border:1px solid var(--border); border-radius:12px; width:100%; max-width:460px; max-height:88vh; overflow-y:auto; padding:20px; box-shadow:0 12px 32px rgba(27,39,51,0.14); }
        .head { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; }
        .head h3 { font-family:'Space Grotesk',sans-serif; font-size:18px; color:var(--text); margin:0; font-weight:600; }
        .body { display:flex; flex-direction:column; gap:14px; }
        .icon-btn { background:none; border:none; color:var(--text-muted); cursor:pointer; padding:4px; }
      `}</style>
    </div>
  );
}

const inputCss = `
  .input, .select { background:var(--surface-alt); border:1px solid var(--border); border-radius:10px; padding:11px 12px; color:var(--text); font-size:15px; outline:none; width:100%; box-sizing:border-box; }
  .input:focus, .select:focus { border-color:var(--accent); }
`;

function PrimaryButton({ children, onClick, full, tone }) {
  return (
    <button onClick={onClick} className={`btn ${full ? 'full' : ''}`}>
      {children}
      <style jsx>{`
        .btn { background:${tone === 'danger' ? 'var(--danger)' : 'var(--accent)'}; color:var(--accent-contrast); border:none; border-radius:10px; padding:12px 18px; font-weight:600; font-size:14px; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; gap:6px; }
        .full { width:100%; }
      `}</style>
    </button>
  );
}

function GhostButton({ children, onClick, full }) {
  return (
    <button onClick={onClick} className={`btn ${full ? 'full' : ''}`}>
      {children}
      <style jsx>{`
        .btn { background:transparent; color:var(--text-muted); border:1px solid var(--border); border-radius:10px; padding:11px 18px; font-weight:600; font-size:14px; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; gap:6px; }
        .full { width:100%; }
      `}</style>
    </button>
  );
}

function EmptyState({ icon: Icon, text }) {
  return (
    <div className="empty">
      <Icon size={28} strokeWidth={1.5} />
      <p>{text}</p>
      <style jsx>{`
        .empty { display:flex; flex-direction:column; align-items:center; gap:10px; padding:48px 20px; color:var(--text-dim); text-align:center; }
        p { margin:0; font-size:14px; }
      `}</style>
    </div>
  );
}

// Campo de valor com máscara no padrão brasileiro: milhar separado por ponto,
// decimais separados por vírgula (ex: 1.234.567,89). Digita-se da direita
// para a esquerda, como em caixas eletrônicos — os 2 últimos dígitos são
// sempre os centavos, e o valor cresce por unidade, dezena, centena, milhar...
function centsToBRL(cents) {
  return (cents / 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function DecimalMaskInput({ value, onChange, placeholder, prefix }) {
  const [display, setDisplay] = useState(() => {
    if (value === '' || value === null || value === undefined) return '';
    return centsToBRL(Math.round(Number(value) * 100));
  });

  useEffect(() => {
    if (value === '' || value === null || value === undefined) { setDisplay(''); return; }
    const expected = centsToBRL(Math.round(Number(value) * 100));
    // só ressincroniza se o valor externo mudou por outro motivo (ex: carregar edição)
    if (Number(value) !== Number(parseDisplay(display))) setDisplay(expected);
  }, [value]); // eslint-disable-line react-hooks/exhaustive-deps

  function parseDisplay(str) {
    if (!str) return 0;
    return Number(str.replace(/\./g, '').replace(',', '.')) || 0;
  }

  const handleChange = (e) => {
    const digits = e.target.value.replace(/\D/g, '');
    if (!digits) { setDisplay(''); onChange(''); return; }
    const cents = parseInt(digits, 10);
    setDisplay(centsToBRL(cents));
    onChange(cents / 100);
  };

  return (
    <div className="dmi-wrap">
      {prefix && <span className="dmi-prefix">{prefix}</span>}
      <input className="input dmi-input" type="text" inputMode="decimal" value={display} onChange={handleChange} placeholder={placeholder || '0,00'} />
      <style jsx>{`
        .dmi-wrap { position:relative; display:flex; align-items:center; }
        .dmi-prefix { position:absolute; left:12px; color:var(--text-dim); font-size:14px; pointer-events:none; }
        .dmi-input { ${prefix ? 'padding-left:30px;' : ''} text-align:right; font-family:'IBM Plex Mono',monospace; }
      `}</style>
    </div>
  );
}

function SectionTitle({ children }) {
  return (
    <h2 className="st">
      {children}
      <style jsx>{`.st { font-family:'Space Grotesk',sans-serif; font-size:14px; letter-spacing:.02em; color:var(--text-muted); margin:20px 0 10px; font-weight:600; text-transform:uppercase; }`}</style>
    </h2>
  );
}

function ExportButton({ onClick, disabled, label = 'Exportar Excel' }) {
  return (
    <button className="exp" onClick={onClick} disabled={disabled}>
      <Download size={14} /> {label}
      <style jsx>{`
        .exp { display:inline-flex; align-items:center; gap:6px; background:none; border:1px solid var(--border); color:var(--accent); font-size:12.5px; font-weight:600; padding:8px 12px; border-radius:9px; cursor:pointer; }
        .exp:disabled { color:var(--text-dim); }
      `}</style>
    </button>
  );
}

function exportToExcel(rows, columns, sheetName, fileName) {
  const aoa = [columns.map((c) => c.header), ...rows.map((r) => columns.map((c) => c.get(r)))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = columns.map((c) => ({ wch: c.width || 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `${fileName}-${todayISO()}.xlsx`);
}

// Agrupa linhas por empresa (ou outro critério) e exporta em Excel com
// cabeçalho de grupo + subtotal, espelhando a organização usada nas telas do app.
function groupRows(rows, groupKeyFn, groupLabelFn) {
  const map = new Map();
  rows.forEach((r) => {
    const key = groupKeyFn(r);
    if (!map.has(key)) map.set(key, { label: groupLabelFn(r), rows: [] });
    map.get(key).rows.push(r);
  });
  return Array.from(map.values());
}

function exportGroupedExcel({ groups, columns, valueColIndex, sheetName, fileName, reportTitle }) {
  const aoa = [];
  aoa.push([reportTitle]);
  aoa.push([`Gerado em ${fmtDate(todayISO())}`]);
  aoa.push([]);
  let grandTotal = 0;
  groups.forEach((g) => {
    aoa.push([g.label]);
    aoa.push(columns.map((c) => c.header));
    let subtotal = 0;
    g.rows.forEach((r) => {
      const rowVals = columns.map((c) => c.get(r));
      if (valueColIndex != null) subtotal += Number(rowVals[valueColIndex]) || 0;
      aoa.push(rowVals);
    });
    if (valueColIndex != null) {
      const subtotalRow = columns.map((_, i) => (i === valueColIndex ? subtotal : (i === 0 ? 'Subtotal' : '')));
      aoa.push(subtotalRow);
      grandTotal += subtotal;
    }
    aoa.push([]);
  });
  if (valueColIndex != null) {
    const totalRow = columns.map((_, i) => (i === valueColIndex ? grandTotal : (i === 0 ? 'TOTAL GERAL' : '')));
    aoa.push(totalRow);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = columns.map((c) => ({ wch: c.width || 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `${fileName}-${todayISO()}.xlsx`);
}

// Gera um PDF do relatório, agrupado por empresa (ou outro critério), com
// cabeçalho de seção, subtotal por grupo e total geral — mesma estrutura das telas.
function exportToPDF({ title, subtitle, groups, columns, valueColIndex, fileName }) {
  const doc = new jsPDF();
  doc.setFontSize(15);
  doc.setTextColor(...PDF_NAVY);
  doc.text(title, 14, 18);
  doc.setFontSize(9);
  doc.setTextColor(120, 130, 140);
  doc.text(subtitle || `Gerado em ${fmtDate(todayISO())}`, 14, 24);

  let startY = 30;
  let grandTotal = 0;

  groups.forEach((g) => {
    doc.setFontSize(11);
    doc.setTextColor(...PDF_NAVY);
    doc.text(g.label, 14, startY);
    startY += 4;

    let subtotal = 0;
    const body = g.rows.map((r) => columns.map((c) => {
      const v = c.get(r);
      if (valueColIndex != null && columns.indexOf(c) === valueColIndex) subtotal += Number(v) || 0;
      return v;
    }));
    if (valueColIndex != null) grandTotal += subtotal;

    const foot = valueColIndex != null
      ? [columns.map((c, i) => (i === valueColIndex ? fmtBRL(subtotal) : (i === 0 ? 'Subtotal' : '')))]
      : undefined;

    autoTable(doc, {
      startY,
      head: [columns.map((c) => c.header)],
      body,
      foot,
      theme: 'grid',
      styles: { fontSize: 8.5, textColor: PDF_TEXT, cellPadding: 3 },
      headStyles: { fillColor: PDF_NAVY, textColor: 255, fontStyle: 'bold' },
      footStyles: { fillColor: PDF_MUTED_BG, textColor: PDF_TEXT, fontStyle: 'bold' },
      margin: { left: 14, right: 14 },
    });
    startY = doc.lastAutoTable.finalY + 10;
    if (startY > 260) { doc.addPage(); startY = 20; }
  });

  if (valueColIndex != null) {
    doc.setFontSize(11);
    doc.setTextColor(...PDF_NAVY);
    doc.text(`Total geral: ${fmtBRL(grandTotal)}`, 14, startY);
  }

  doc.save(`${fileName}-${todayISO()}.pdf`);
}

// ---------- Shared card styles ----------
function ViewStyle() {
  return (
    <style jsx global>{`
      .view { padding:16px 20px 24px; display:flex; flex-direction:column; gap:12px; }
      .card { background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:14px 16px; box-shadow:0 1px 2px rgba(27,39,51,0.04); }
      .card-head { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; gap:8px; }
      .card-title { display:flex; align-items:center; gap:8px; font-family:'Space Grotesk',sans-serif; font-size:15px; color:var(--text); font-weight:600; }
      .card-actions { display:flex; gap:2px; flex-shrink:0; }
      .card-grid { display:grid; grid-template-columns:1fr 1fr; gap:10px 14px; }
      .k { display:block; font-size:11px; color:var(--text-dim); text-transform:uppercase; letter-spacing:.03em; margin-bottom:2px; }
      .v { display:block; font-size:14px; color:var(--text); }
      .mono { font-family:'IBM Plex Mono',monospace; }
      .icon-btn { background:none; border:none; color:var(--text-dim); cursor:pointer; padding:5px; border-radius:6px; display:flex; }
      .icon-btn:hover { color:var(--accent); }
      .icon-btn.danger:hover { color:var(--danger); }
      .row { display:flex; justify-content:space-between; align-items:center; padding:12px 14px; background:var(--surface); border:1px solid var(--border); border-radius:12px; margin-bottom:8px; font-size:14px; color:var(--text); }
      .row-sub { font-size:12px; margin-top:2px; }
      .badge { display:inline-block; background:var(--surface-alt); border:1px solid var(--border); color:var(--text-muted); font-size:10.5px; font-weight:600; padding:2px 7px; border-radius:6px; margin-left:6px; }
      .add-inline-btn { display:inline-flex; align-items:center; gap:5px; background:var(--accent); color:var(--accent-contrast); border:none; border-radius:9px; padding:9px 12px; font-size:12.5px; font-weight:700; cursor:pointer; white-space:nowrap; }
      .chart-box { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:12px 8px 4px; margin-bottom:4px; box-shadow:0 1px 2px rgba(27,39,51,0.04); }
      .report-filters { display:flex; gap:8px; padding:4px 20px 0; flex-wrap:wrap; }
      .rtable { width:calc(100% - 40px); margin:0 20px; border-collapse:collapse; font-size:13px; }
      .rtable thead th { text-align:left; font-size:10.5px; text-transform:uppercase; color:var(--text-dim); padding:0 8px 8px 0; border-bottom:1px solid var(--border); }
      .rtable tbody td { padding:10px 8px 10px 0; border-bottom:1px solid var(--border); color:var(--text); }
      .rtable .dim { color:var(--text-muted); }
      .rtable .right { text-align:right; }
      ${inputCss}
    `}</style>
  );
}

// ============================================================
// MAIN APP
// ============================================================

export default function Home() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState(null);
  const [tab, setTab] = useState('dashboard');
  const [modal, setModal] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [reportView, setReportView] = useState(null);
  const [taxasSub, setTaxasSub] = useState(null);
  const [saving, setSaving] = useState(false);

  const reload = async () => {
    try {
      setErrorMsg(null);
      const next = await db.fetchAll();
      setData(next);
    } catch (e) {
      console.error(e);
      setErrorMsg('Não foi possível conectar ao banco de dados. Confira as credenciais do Supabase.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { reload(); }, []);

  const withSave = async (fn) => {
    setSaving(true);
    try { await fn(); await reload(); } catch (e) { console.error(e); alert('Erro ao salvar. Tente novamente.'); }
    setSaving(false);
  };

  if (loading) return <CenterMsg text="Carregando…" />;
  if (errorMsg) return <CenterMsg text={errorMsg} error />;

  const saldoTotal = data.bancos.reduce((s, b) => s + (Number(b.saldo) || 0), 0);

  return (
    <div className="app">
      <header className="header">
        <div>
          <div className="eyebrow">Saldo consolidado</div>
          <div className="saldo-total">{fmtBRL(saldoTotal)}</div>
        </div>
        <div className="header-meta">
          <span>{data.empresas.length} empresa{data.empresas.length === 1 ? '' : 's'}</span>
          <span>·</span>
          <span>{data.bancos.length} conta{data.bancos.length === 1 ? '' : 's'} bancária{data.bancos.length === 1 ? '' : 's'}</span>
          <button className="refresh-btn" onClick={reload} title="Atualizar"><RefreshCw size={13} /></button>
        </div>
      </header>

      <main className="content">
        {tab === 'dashboard' && <Dashboard data={data} onGoReports={() => setTab('relatorios')} />}
        {tab === 'empresas' && (
          <EmpresasView data={data} onEdit={(it) => setModal({ type: 'empresa', item: it })}
            onDelete={(id, label) => setConfirmDelete({ type: 'empresa', id, label })} />
        )}
        {tab === 'bancos' && (
          <BancosView data={data}
            onEdit={(it) => setModal({ type: 'banco', item: it })}
            onDelete={(id, label) => setConfirmDelete({ type: 'banco', id, label })}
            onNewTransfer={() => setModal({ type: 'transferencia', item: null })}
            onDeleteTransfer={(t) => withSave(() => db.excluirTransferencia(t, data.bancos))} />
        )}
        {tab === 'contas' && (
          <ContasView data={data}
            onDarBaixa={(c, dataPagamento) => withSave(() => db.darBaixaConta(c, dataPagamento))}
            onReverterBaixa={(c) => withSave(() => db.reverterBaixaConta(c))}
            onEdit={(it) => setModal({ type: 'conta', item: it })}
            onDelete={(id, label) => setConfirmDelete({ type: 'conta', id, label })} />
        )}
        {tab === 'emprestimos' && (
          <EmprestimosView data={data}
            onEdit={(it) => setModal({ type: 'emprestimo', item: it })}
            onDelete={(id, label) => setConfirmDelete({ type: 'emprestimo', id, label })}
            onToggleParcela={(p) => withSave(() => db.toggleParcela(p))}
            onGoReport={(view) => { setTab('relatorios'); setReportView(view); }} />
        )}
        {tab === 'taxas' && (
          <TaxasHome data={data} sub={taxasSub} setSub={setTaxasSub}
            onNewCartao={() => setModal({ type: 'cartao', item: null })}
            onEditCartao={(it) => setModal({ type: 'cartao', item: it })}
            onDeleteCartao={(id, label) => setConfirmDelete({ type: 'cartao', id, label })}
            onNewBoleto={() => setModal({ type: 'boleto', item: null })}
            onEditBoleto={(it) => setModal({ type: 'boleto', item: it })}
            onDeleteBoleto={(id, label) => setConfirmDelete({ type: 'boleto', id, label })} />
        )}
        {tab === 'seguros' && (
          <SegurosView data={data}
            onEdit={(it) => setModal({ type: 'seguro', item: it })}
            onDelete={(id, label) => setConfirmDelete({ type: 'seguro', id, label })}
            onGoReport={() => { setTab('relatorios'); setReportView('seguros'); }} />
        )}
        {tab === 'vendas' && (
          <VendasView data={data}
            onSaveMes={(v) => withSave(() => db.saveVendaMes(v))}
            onGoReport={() => { setTab('relatorios'); setReportView('vendas'); }} />
        )}
        {tab === 'relatorios' && <ReportsHome view={reportView} setView={setReportView} data={data} />}
      </main>

      {['empresas', 'bancos', 'contas', 'emprestimos', 'seguros'].includes(tab) && (
        <button className="fab" onClick={() => setModal({ type: tab === 'empresas' ? 'empresa' : tab === 'bancos' ? 'banco' : tab === 'contas' ? 'conta' : tab === 'emprestimos' ? 'emprestimo' : 'seguro', item: null })}>
          <Plus size={24} />
        </button>
      )}

      <nav className="tabbar">
        {[
          { key: 'dashboard', label: 'Geral', icon: LayoutDashboard },
          { key: 'empresas', label: 'Empresas', icon: Building2 },
          { key: 'bancos', label: 'Bancos', icon: Landmark },
          { key: 'contas', label: 'A pagar', icon: CalendarClock },
          { key: 'emprestimos', label: 'Empréstimos', icon: HandCoins },
          { key: 'taxas', label: 'Taxas', icon: Percent },
          { key: 'seguros', label: 'Seguros', icon: Shield },
          { key: 'vendas', label: 'Vendas', icon: ShoppingCart },
          { key: 'relatorios', label: 'Relatórios', icon: FileBarChart },
        ].map(({ key, label, icon: Icon }) => (
          <button key={key} className={`tab ${tab === key ? 'active' : ''}`} onClick={() => { setTab(key); setReportView(null); setTaxasSub(null); }}>
            <Icon size={19} strokeWidth={tab === key ? 2.2 : 1.6} /><span>{label}</span>
          </button>
        ))}
      </nav>

      {modal?.type === 'empresa' && <EmpresaForm item={modal.item} onClose={() => setModal(null)} onSave={(v) => withSave(async () => { await db.saveEmpresa(v); setModal(null); })} />}
      {modal?.type === 'banco' && <BancoForm item={modal.item} empresas={data.empresas} onClose={() => setModal(null)} onSave={(v) => withSave(async () => { await db.saveBanco(v); setModal(null); })} />}
      {modal?.type === 'conta' && <ContaForm item={modal.item} empresas={data.empresas} onClose={() => setModal(null)} onSave={(v) => withSave(async () => { await db.saveConta(v); setModal(null); })} />}
      {modal?.type === 'emprestimo' && <EmprestimoForm item={modal.item} empresas={data.empresas} onClose={() => setModal(null)} onSave={(v) => withSave(async () => { await db.saveEmprestimo(v, modal.item?.parcelas || []); setModal(null); })} />}
      {modal?.type === 'transferencia' && <TransferForm bancos={data.bancos} onClose={() => setModal(null)} onSave={(v) => withSave(async () => { await db.criarTransferencia(v, data.bancos); setModal(null); })} />}
      {modal?.type === 'cartao' && <CartaoTaxaForm item={modal.item} onClose={() => setModal(null)} onSave={(v) => withSave(async () => { await db.saveTaxaCartao(v); setModal(null); })} />}
      {modal?.type === 'boleto' && <BoletoTaxaForm item={modal.item} onClose={() => setModal(null)} onSave={(v) => withSave(async () => { await db.saveTaxaBoleto(v); setModal(null); })} />}
      {modal?.type === 'seguro' && <SeguroForm item={modal.item} empresas={data.empresas} onClose={() => setModal(null)} onSave={(v) => withSave(async () => { await db.saveSeguro(v); setModal(null); })} />}

      {confirmDelete && (
        <Modal title="Excluir registro" onClose={() => setConfirmDelete(null)}>
          <p style={{ color: 'var(--text-muted)', fontSize: 14, margin: 0 }}>Tem certeza que deseja excluir <strong style={{ color: 'var(--text)' }}>{confirmDelete.label}</strong>? Essa ação não pode ser desfeita.</p>
          <div style={{ display: 'flex', gap: 10 }}>
            <GhostButton full onClick={() => setConfirmDelete(null)}>Cancelar</GhostButton>
            <PrimaryButton full tone="danger" onClick={() => withSave(async () => {
              const { type, id } = confirmDelete;
              if (type === 'empresa') await db.deleteEmpresa(id);
              if (type === 'banco') await db.deleteBanco(id);
              if (type === 'conta') await db.deleteConta(id);
              if (type === 'emprestimo') await db.deleteEmprestimo(id);
              if (type === 'cartao') await db.deleteTaxaCartao(id);
              if (type === 'boleto') await db.deleteTaxaBoleto(id);
              if (type === 'seguro') await db.deleteSeguro(id);
              setConfirmDelete(null);
            })}>Excluir</PrimaryButton>
          </div>
        </Modal>
      )}

      {saving && <div className="saving-overlay">Salvando…</div>}

      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');
        body { font-family:'Inter',sans-serif; }
      `}</style>
      <style jsx>{`
        .app { min-height:100vh; background:var(--bg); display:flex; flex-direction:column; max-width:520px; margin:0 auto; position:relative; }
        .header { padding:20px 20px 18px; border-bottom:1px solid var(--border); }
        .eyebrow { font-size:11px; letter-spacing:.08em; text-transform:uppercase; color:var(--text-dim); margin-bottom:4px; }
        .saldo-total { font-family:'IBM Plex Mono',monospace; font-size:32px; font-weight:600; color:var(--text); }
        .header-meta { margin-top:8px; display:flex; align-items:center; gap:8px; font-size:12px; color:var(--text-dim); }
        .refresh-btn { margin-left:auto; background:none; border:1px solid var(--border); color:var(--text-muted); border-radius:6px; padding:4px 6px; cursor:pointer; display:flex; }
        .content { flex:1; overflow-y:auto; padding-bottom:70px; }
        .fab { position:fixed; right:calc(50% - 260px + 20px); bottom:80px; width:52px; height:52px; border-radius:50%; background:var(--accent); color:var(--accent-contrast); border:none; display:flex; align-items:center; justify-content:center; box-shadow:0 6px 18px var(--accent-shadow); cursor:pointer; z-index:10; }
        @media (max-width: 560px) { .fab { right:20px; } }
        .tabbar { position:fixed; bottom:0; left:0; right:0; max-width:520px; margin:0 auto; display:flex; background:var(--surface-alt); border-top:1px solid var(--border); padding:6px 2px calc(6px + env(safe-area-inset-bottom)); z-index:20; overflow-x:auto; }
        .tab { flex:1; min-width:58px; display:flex; flex-direction:column; align-items:center; gap:3px; background:none; border:none; color:var(--text-dim); padding:6px 1px; cursor:pointer; }
        .tab span { font-size:9.5px; }
        .tab.active { color:var(--accent); }
        .saving-overlay { position:fixed; top:12px; left:50%; transform:translateX(-50%); background:var(--surface); border:1px solid var(--border); color:var(--accent); font-size:12.5px; padding:8px 14px; border-radius:20px; z-index:60; }
      `}</style>
    </div>
  );
}

function CenterMsg({ text, error }) {
  return (
    <div style={{ minHeight: '100vh', background: 'var(--bg)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: error ? 'var(--danger)' : 'var(--text-dim)', fontFamily: 'Inter,sans-serif', padding: 24, textAlign: 'center' }}>
      {text}
    </div>
  );
}

// ---------- Dashboard ----------

function Dashboard({ data, onGoReports }) {
  const proximasContas = data.contas.filter((c) => c.status === 'pendente').sort((a, b) => a.data_vencimento.localeCompare(b.data_vencimento)).slice(0, 5);
  const totalEmprestimos = data.emprestimos.reduce((s, e) => s + (Number(e.valor_total) || 0), 0);
  const totalPendente = data.contas.filter((c) => c.status === 'pendente').reduce((s, c) => s + (Number(c.valor) || 0), 0);
  const saldoPorEmpresa = data.empresas.map((emp) => ({
    name: emp.nome.length > 10 ? emp.nome.slice(0, 9) + '…' : emp.nome,
    valor: data.bancos.filter((b) => b.empresa_id === emp.id).reduce((s, b) => s + (Number(b.saldo) || 0), 0),
  }));
  const segurosVencendo = (data.seguros || [])
    .filter((s) => statusSeguro(s.vigencia_fim) !== 'vigente')
    .sort((a, b) => a.vigencia_fim.localeCompare(b.vigencia_fim))
    .slice(0, 5);

  if (!data.empresas.length && !data.bancos.length) {
    return <EmptyState icon={Building2} text="Cadastre sua primeira empresa para começar." />;
  }

  return (
    <div className="dash">
      <div className="cards-row">
        <div className="stat-card"><TrendingDown size={16} color="var(--warning)" /><div className="stat-label">A pagar (pendente)</div><div className="stat-value">{fmtBRL(totalPendente)}</div></div>
        <div className="stat-card"><HandCoins size={16} color="var(--danger)" /><div className="stat-label">Empréstimos ativos</div><div className="stat-value">{fmtBRL(totalEmprestimos)}</div></div>
      </div>

      {saldoPorEmpresa.length > 0 && (
        <>
          <SectionTitle>Saldos por empresa</SectionTitle>
          <div className="chart-box">
            <ResponsiveContainer width="100%" height={Math.max(120, saldoPorEmpresa.length * 42)}>
              <BarChart data={saldoPorEmpresa} layout="vertical" margin={{ top: 4, right: 12, bottom: 4, left: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" horizontal={false} />
                <XAxis type="number" tick={{ fill: 'var(--text-dim)', fontSize: 10 }} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} axisLine={false} tickLine={false} />
                <YAxis type="category" dataKey="name" tick={{ fill: 'var(--text-muted)', fontSize: 11 }} axisLine={false} tickLine={false} width={70} />
                <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, fontSize: 12 }} labelStyle={{ color: 'var(--text)' }} />
                <Bar dataKey="valor" fill="var(--accent)" radius={[0, 4, 4, 0]} barSize={16} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </>
      )}

      <SectionTitle>Próximos vencimentos</SectionTitle>
      {proximasContas.length === 0 ? <p className="muted">Nenhuma conta pendente.</p> : (
        <div className="list">
          {proximasContas.map((c) => {
            const dias = daysUntil(c.data_vencimento);
            return (
              <div key={c.id} className="row">
                <div>
                  <div>{c.descricao}</div>
                  <div className="row-sub" style={{ color: dias <= 3 ? 'var(--danger)' : 'var(--text-dim)' }}>
                    {fmtDate(c.data_vencimento)} {dias === 0 ? '· vence hoje' : dias < 0 ? `· atrasada ${Math.abs(dias)}d` : `· em ${dias}d`}
                  </div>
                </div>
                <span className="mono">{fmtBRL(c.valor)}</span>
              </div>
            );
          })}
        </div>
      )}

      {segurosVencendo.length > 0 && (
        <>
          <SectionTitle>Seguros vencendo</SectionTitle>
          <div className="list">
            {segurosVencendo.map((s) => {
              const dias = daysUntil(s.vigencia_fim);
              return (
                <div key={s.id} className="row">
                  <div>
                    <div><Shield size={12} style={{ marginRight: 5, verticalAlign: -1 }} />{s.objeto}</div>
                    <div className="row-sub" style={{ color: dias < 0 ? 'var(--danger)' : 'var(--warning)' }}>
                      {fmtDate(s.vigencia_fim)} {dias < 0 ? `· vencido há ${Math.abs(dias)}d` : `· vence em ${dias}d`}
                    </div>
                  </div>
                  <SeguroStatusBadge vigenciaFim={s.vigencia_fim} />
                </div>
              );
            })}
          </div>
        </>
      )}

      <button className="reports-link" onClick={onGoReports}><FileBarChart size={16} /> Ver relatórios completos <ChevronRight size={15} /></button>

      <style jsx>{`
        .dash { padding:16px 20px 24px; display:flex; flex-direction:column; gap:4px; }
        .cards-row { display:flex; gap:10px; margin-bottom:8px; }
        .stat-card { flex:1; background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:14px; display:flex; flex-direction:column; gap:6px; box-shadow:0 1px 2px rgba(27,39,51,0.04); }
        .stat-label { font-size:12px; color:var(--text-muted); }
        .stat-value { font-family:'IBM Plex Mono',monospace; font-size:18px; color:var(--text); font-weight:600; }
        .muted { color:var(--text-dim); font-size:14px; padding:8px 0 16px; }
        .reports-link { margin-top:18px; display:flex; align-items:center; justify-content:center; gap:8px; width:100%; background:var(--surface); border:1px solid var(--border); color:var(--accent); font-weight:600; font-size:13px; padding:13px; border-radius:12px; cursor:pointer; }
      `}</style>
      <ViewStyle />
    </div>
  );
}

// ---------- Empresas ----------

function EmpresasView({ data, onEdit, onDelete }) {
  if (!data.empresas.length) return <EmptyState icon={Building2} text="Nenhuma empresa cadastrada. Toque em + para adicionar." />;
  return (
    <div className="view">
      {data.empresas.map((emp) => {
        const saldo = data.bancos.filter((b) => b.empresa_id === emp.id).reduce((s, b) => s + (Number(b.saldo) || 0), 0);
        const pendentes = data.contas.filter((c) => c.empresa_id === emp.id && c.status === 'pendente').length;
        return (
          <div key={emp.id} className="card">
            <div className="card-head">
              <div className="card-title"><Building2 size={16} /> {emp.nome}</div>
              <div className="card-actions">
                <button className="icon-btn" onClick={() => onEdit(emp)}><Pencil size={16} /></button>
                <button className="icon-btn danger" onClick={() => onDelete(emp.id, emp.nome)}><Trash2 size={16} /></button>
              </div>
            </div>
            <div className="card-grid">
              <div><span className="k">Saldo</span><span className="v mono">{fmtBRL(saldo)}</span></div>
              <div><span className="k">Contas pendentes</span><span className="v">{pendentes}</span></div>
            </div>
          </div>
        );
      })}
      <ViewStyle />
    </div>
  );
}

// ---------- Bancos ----------

function BancosView({ data, onEdit, onDelete, onNewTransfer, onDeleteTransfer }) {
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const bancoNome = (id) => data.bancos.find((b) => b.id === id)?.nome_banco || '—';
  const transfers = [...data.transferencias].sort((a, b) => b.data.localeCompare(a.data));
  return (
    <div className="view">
      {data.bancos.length >= 2 && (
        <button className="transfer-btn" onClick={onNewTransfer}><ArrowLeftRight size={16} /> Nova transferência entre bancos</button>
      )}
      {data.bancos.length === 0 ? <EmptyState icon={Landmark} text="Nenhum banco cadastrado. Toque em + para adicionar." /> : data.bancos.map((b) => (
        <div key={b.id} className="card">
          <div className="card-head">
            <div className="card-title"><Landmark size={16} /> {b.nome_banco}</div>
            <div className="card-actions">
              <button className="icon-btn" onClick={() => onEdit({ id: b.id, empresaId: b.empresa_id, nomeBanco: b.nome_banco, saldo: b.saldo })}><Pencil size={16} /></button>
              <button className="icon-btn danger" onClick={() => onDelete(b.id, b.nome_banco)}><Trash2 size={16} /></button>
            </div>
          </div>
          <div className="card-grid">
            <div><span className="k">Empresa</span><span className="v">{empresaNome(b.empresa_id)}</span></div>
            <div><span className="k">Saldo</span><span className="v mono">{fmtBRL(b.saldo)}</span></div>
          </div>
        </div>
      ))}
      {transfers.length > 0 && (
        <>
          <SectionTitle>Transferências recentes</SectionTitle>
          {transfers.map((t) => (
            <div key={t.id} className="card">
              <div className="transfer-row">
                <div>
                  <div className="transfer-path"><ArrowLeftRight size={13} /> {bancoNome(t.banco_origem_id)} → {bancoNome(t.banco_destino_id)}</div>
                  <div className="row-sub">{fmtDate(t.data)}{t.descricao ? ` · ${t.descricao}` : ''}</div>
                </div>
                <div className="transfer-actions">
                  <span className="mono">{fmtBRL(t.valor)}</span>
                  <button className="icon-btn danger" onClick={() => onDeleteTransfer(t)}><Trash2 size={15} /></button>
                </div>
              </div>
            </div>
          ))}
        </>
      )}
      <style jsx>{`
        .transfer-btn { display:flex; align-items:center; justify-content:center; gap:8px; width:100%; background:var(--accent-soft-10); border:1px solid var(--border); color:var(--accent); font-weight:600; font-size:13.5px; padding:12px; border-radius:12px; cursor:pointer; margin-bottom:4px; }
        .transfer-row { display:flex; justify-content:space-between; align-items:center; gap:10px; }
        .transfer-path { display:flex; align-items:center; gap:6px; font-size:13.5px; color:var(--text); font-weight:600; }
        .transfer-actions { display:flex; align-items:center; gap:8px; }
      `}</style>
      <ViewStyle />
    </div>
  );
}

// ---------- Contas a pagar ----------

function ContasView({ data, onDarBaixa, onReverterBaixa, onEdit, onDelete }) {
  const [baixaAlvo, setBaixaAlvo] = useState(null); // conta pendente que está recebendo baixa
  const [reverterAlvo, setReverterAlvo] = useState(null); // conta paga a reverter

  if (!data.contas.length) return <EmptyState icon={CalendarClock} text="Nenhuma conta cadastrada. Toque em + para adicionar." />;
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const sorted = [...data.contas].sort((a, b) => a.data_vencimento.localeCompare(b.data_vencimento));
  return (
    <div className="view">
      {sorted.map((c) => {
        const dias = daysUntil(c.data_vencimento);
        const atrasada = c.status === 'pendente' && dias < 0;
        return (
          <div key={c.id} className="card">
            <div className="card-head">
              <div className="card-title"><CalendarClock size={16} /> {c.descricao}</div>
              <div className="card-actions">
                <button className="icon-btn" onClick={() => onEdit({ id: c.id, empresaId: c.empresa_id, descricao: c.descricao, valor: c.valor, dataVencimento: c.data_vencimento, status: c.status, dataPagamento: c.data_pagamento })}><Pencil size={16} /></button>
                <button className="icon-btn danger" onClick={() => onDelete(c.id, c.descricao)}><Trash2 size={16} /></button>
              </div>
            </div>
            <div className="card-grid">
              <div><span className="k">Empresa</span><span className="v">{empresaNome(c.empresa_id)}</span></div>
              <div><span className="k">Vencimento</span><span className="v" style={{ color: atrasada ? 'var(--danger)' : undefined }}>{fmtDate(c.data_vencimento)}</span></div>
              <div><span className="k">Valor</span><span className="v mono">{fmtBRL(c.valor)}</span></div>
            </div>
            <button className={`status-btn ${c.status}`} onClick={() => c.status === 'pendente' ? setBaixaAlvo(c) : setReverterAlvo(c)}>
              {c.status === 'pago' ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />}
              {c.status === 'pago' ? `Baixa em ${fmtDate(c.data_pagamento)}` : atrasada ? 'Dar baixa · atrasada' : 'Dar baixa'}
            </button>
          </div>
        );
      })}
      <style jsx>{`
        .status-btn { margin-top:10px; display:inline-flex; align-items:center; gap:6px; border:none; border-radius:8px; padding:7px 12px; font-size:12px; font-weight:600; cursor:pointer; }
        .status-btn.pendente { background:var(--warning-soft-15); color:var(--warning); }
        .status-btn.pago { background:var(--success-soft-15); color:var(--success); }
      `}</style>
      <ViewStyle />

      {baixaAlvo && (
        <DarBaixaModal
          conta={baixaAlvo}
          onClose={() => setBaixaAlvo(null)}
          onConfirm={(data) => { onDarBaixa(baixaAlvo, data); setBaixaAlvo(null); }}
        />
      )}
      {reverterAlvo && (
        <Modal title="Reverter baixa" onClose={() => setReverterAlvo(null)}>
          <p style={{ color: 'var(--text-muted)', fontSize: 14, margin: 0 }}>
            Isso volta <strong style={{ color: 'var(--text)' }}>{reverterAlvo.descricao}</strong> para o status "pendente" e remove a data de baixa. Deseja continuar?
          </p>
          <div style={{ display: 'flex', gap: 10 }}>
            <GhostButton full onClick={() => setReverterAlvo(null)}>Cancelar</GhostButton>
            <PrimaryButton full onClick={() => { onReverterBaixa(reverterAlvo); setReverterAlvo(null); }}>Reverter</PrimaryButton>
          </div>
        </Modal>
      )}
    </div>
  );
}

function DarBaixaModal({ conta, onClose, onConfirm }) {
  const [data, setData] = useState(todayISO());
  return (
    <Modal title="Dar baixa na conta" onClose={onClose}>
      <p style={{ color: 'var(--text-muted)', fontSize: 14, margin: 0 }}>
        Confirmar pagamento de <strong style={{ color: 'var(--text)' }}>{conta.descricao}</strong> · <span className="mono">{fmtBRL(conta.valor)}</span>
      </p>
      <Field label="Data do pagamento">
        <input className="input" type="date" value={data} onChange={(e) => setData(e.target.value)} />
      </Field>
      <PrimaryButton full onClick={() => onConfirm(data)}>Confirmar baixa</PrimaryButton>
      <style jsx>{inputCss}</style>
    </Modal>
  );
}

// ---------- Empréstimos ----------

function EmprestimosView({ data, onEdit, onDelete, onToggleParcela, onGoReport }) {
  const [expandedId, setExpandedId] = useState(null);
  const [filtroTipo, setFiltroTipo] = useState('todos'); // todos | banco | empresa
  if (!data.emprestimos.length) return <EmptyState icon={HandCoins} text="Nenhum empréstimo cadastrado. Toque em + para adicionar." />;
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const lista = data.emprestimos.filter((e) => filtroTipo === 'todos' || e.tipo === filtroTipo);
  return (
    <div className="view">
      <div className="seg-filter">
        <button className={filtroTipo === 'todos' ? 'active' : ''} onClick={() => setFiltroTipo('todos')}>Todos</button>
        <button className={filtroTipo === 'banco' ? 'active' : ''} onClick={() => setFiltroTipo('banco')}>Bancários</button>
        <button className={filtroTipo === 'empresa' ? 'active' : ''} onClick={() => setFiltroTipo('empresa')}>Entre empresas</button>
      </div>
      {lista.length === 0 ? <EmptyState icon={HandCoins} text="Nenhum empréstimo nessa categoria." /> : lista.map((e) => {
        const parcelas = e.parcelas || [];
        const pagas = parcelas.filter((p) => p.status === 'pago').length;
        const expanded = expandedId === e.id;
        const isBanco = (e.tipo || 'banco') === 'banco';
        return (
          <div key={e.id} className="card">
            <div className="card-head">
              <div className="card-title">
                <HandCoins size={16} /> {isBanco ? e.credor : `${empresaNome(e.empresa_credora_id)} → ${empresaNome(e.empresa_id)}`}
                <span className={`badge ${isBanco ? '' : 'badge-alt'}`}>{isBanco ? 'Bancário' : 'Intercompany'}</span>
              </div>
              <div className="card-actions">
                <button className="icon-btn" onClick={() => onEdit({ id: e.id, empresaId: e.empresa_id, credor: e.credor, valorTotal: e.valor_total, valorParcela: e.valor_parcela, numParcelas: e.num_parcelas, dataInicio: e.data_inicio, parcelas: e.parcelas, tipo: e.tipo || 'banco', linhaCredito: e.linha_credito, empresaCredoraId: e.empresa_credora_id, temCarencia: e.tem_carencia, prazoCarencia: e.prazo_carencia })}><Pencil size={16} /></button>
                <button className="icon-btn danger" onClick={() => onDelete(e.id, isBanco ? e.credor : `Intercompany ${empresaNome(e.empresa_id)}`)}><Trash2 size={16} /></button>
              </div>
            </div>
            <div className="card-grid">
              <div><span className="k">{isBanco ? 'Empresa devedora' : 'Empresa devedora'}</span><span className="v">{empresaNome(e.empresa_id)}</span></div>
              {isBanco ? (
                <div><span className="k">Linha de crédito</span><span className="v">{e.linha_credito || '—'}</span></div>
              ) : (
                <div><span className="k">Empresa credora</span><span className="v">{empresaNome(e.empresa_credora_id)}</span></div>
              )}
              <div><span className="k">Valor total</span><span className="v mono">{fmtBRL(e.valor_total)}</span></div>
              <div><span className="k">Parcela</span><span className="v mono">{fmtBRL(e.valor_parcela)}</span></div>
              <div><span className="k">Parcelas pagas</span><span className="v">{pagas} / {parcelas.length}</span></div>
              <div><span className="k">Início</span><span className="v">{fmtDate(e.data_inicio)}</span></div>
              <div><span className="k">Carência</span><span className="v">{e.tem_carencia ? `${e.prazo_carencia} ${e.prazo_carencia === 1 ? 'mês' : 'meses'}` : 'Sem carência'}</span></div>
            </div>
            {parcelas.length > 0 && (
              <button className="expand-btn" onClick={() => setExpandedId(expanded ? null : e.id)}>
                {expanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />} {expanded ? 'Ocultar parcelas' : 'Ver parcelas'}
              </button>
            )}
            {expanded && (
              <div className="parcelas-list">
                {parcelas.map((p) => {
                  const dias = daysUntil(p.data_vencimento);
                  const atrasada = p.status === 'pendente' && dias < 0;
                  return (
                    <button key={p.id} className={`parcela-row ${p.status}`} onClick={() => onToggleParcela(p)}>
                      <span className="parcela-num">{p.numero}ª</span>
                      <span className="parcela-date" style={{ color: atrasada ? 'var(--danger)' : undefined }}>{p.status === 'pago' ? `Pago ${fmtDate(p.data_pagamento)}` : fmtDate(p.data_vencimento)}</span>
                      <span className="parcela-valor mono">{fmtBRL(p.valor)}</span>
                      {p.status === 'pago' ? <CheckCircle2 size={15} color="var(--success)" /> : <AlertCircle size={15} color={atrasada ? '#B3413E' : '#B8860B'} />}
                    </button>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}

      <div className="reports-links">
        <button className="reports-link" onClick={() => onGoReport('emprestimos-empresa')}><FileBarChart size={16} /> Relatório: empréstimos por empresa <ChevronRight size={15} /></button>
        <button className="reports-link" onClick={() => onGoReport('parcelas-emprestimo')}><FileBarChart size={16} /> Relatório: parcelas de empréstimos <ChevronRight size={15} /></button>
      </div>

      <style jsx>{`
        .expand-btn { margin-top:10px; display:inline-flex; align-items:center; gap:6px; background:none; border:1px solid var(--border); border-radius:8px; padding:7px 12px; font-size:12px; font-weight:600; color:var(--text-muted); cursor:pointer; }
        .parcelas-list { margin-top:10px; display:flex; flex-direction:column; gap:6px; }
        .parcela-row { display:flex; align-items:center; gap:10px; background:var(--surface-alt); border:1px solid var(--border-strong); border-radius:9px; padding:9px 10px; font-size:12.5px; color:var(--text); cursor:pointer; width:100%; text-align:left; }
        .parcela-num { color:var(--text-dim); width:26px; flex-shrink:0; }
        .parcela-date { flex:1; }
        .parcela-valor { color:var(--text-muted); }
        .parcela-row.pago { opacity:0.65; }
        .seg-filter { display:flex; gap:6px; margin-bottom:4px; }
        .seg-filter button { flex:1; background:var(--surface); border:1px solid var(--border); color:var(--text-muted); font-size:12.5px; font-weight:600; padding:9px 8px; border-radius:8px; cursor:pointer; }
        .seg-filter button.active { background:var(--accent); border-color:var(--accent); color:var(--accent-contrast); }
        .badge-alt { background:var(--accent-soft-10); border-color:var(--accent); color:var(--accent); }
        .reports-links { display:flex; flex-direction:column; gap:8px; margin-top:8px; }
        .reports-link { display:flex; align-items:center; justify-content:center; gap:8px; width:100%; background:var(--surface); border:1px solid var(--border); color:var(--accent); font-weight:600; font-size:13px; padding:13px; border-radius:10px; cursor:pointer; }
      `}</style>
      <ViewStyle />
    </div>
  );
}

// ---------- Seguros ----------

function diasParaVencimento(iso) { return daysUntil(iso); }

function statusSeguro(vigenciaFim) {
  const dias = daysUntil(vigenciaFim);
  if (dias < 0) return 'vencido';
  if (dias <= 30) return 'vencendo';
  return 'vigente';
}

function SeguroStatusBadge({ vigenciaFim }) {
  const status = statusSeguro(vigenciaFim);
  const dias = daysUntil(vigenciaFim);
  const label = status === 'vencido' ? `Vencido há ${Math.abs(dias)}d` : status === 'vencendo' ? `Vence em ${dias}d` : 'Vigente';
  return (
    <span className={`seguro-badge ${status}`}>
      {status === 'vigente' ? <CheckCircle2 size={13} /> : <AlertCircle size={13} />} {label}
      <style jsx>{`
        .seguro-badge { display:inline-flex; align-items:center; gap:5px; font-size:12px; font-weight:600; padding:4px 9px; border-radius:7px; }
        .seguro-badge.vigente { background:var(--success-soft-15); color:var(--success); }
        .seguro-badge.vencendo { background:var(--warning-soft-15); color:var(--warning); }
        .seguro-badge.vencido { background:var(--danger-soft-12); color:var(--danger); }
      `}</style>
    </span>
  );
}

function SegurosView({ data, onEdit, onDelete, onGoReport }) {
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const formaLabel = (v) => FORMAS_PAGAMENTO_SEGURO.find((f) => f.value === v)?.label || v;
  const lista = [...(data.seguros || [])].sort((a, b) => a.vigencia_fim.localeCompare(b.vigencia_fim));

  if (!lista.length) return <EmptyState icon={Shield} text="Nenhum seguro cadastrado. Toque em + para adicionar." />;

  return (
    <div className="view">
      {lista.map((s) => (
        <div key={s.id} className="card">
          <div className="card-head">
            <div className="card-title"><Shield size={16} /> {s.objeto}</div>
            <div className="card-actions">
              <button className="icon-btn" onClick={() => onEdit({
                id: s.id, empresaId: s.empresa_id, objeto: s.objeto, vigenciaInicio: s.vigencia_inicio, vigenciaFim: s.vigencia_fim,
                principalCondutor: s.principal_condutor, seguradora: s.seguradora, formaPagamento: s.forma_pagamento,
                numParcelas: s.num_parcelas, valorParcela: s.valor_parcela, valorTotal: s.valor_total,
              })}><Pencil size={16} /></button>
              <button className="icon-btn danger" onClick={() => onDelete(s.id, s.objeto)}><Trash2 size={16} /></button>
            </div>
          </div>
          <div className="card-grid">
            <div><span className="k">Empresa</span><span className="v">{empresaNome(s.empresa_id)}</span></div>
            <div><span className="k">Seguradora</span><span className="v">{s.seguradora}</span></div>
            <div><span className="k">Vigência</span><span className="v">{fmtDate(s.vigencia_inicio)} – {fmtDate(s.vigencia_fim)}</span></div>
            <div><span className="k">Principal condutor</span><span className="v">{s.principal_condutor || '—'}</span></div>
            <div><span className="k">Pagamento</span><span className="v">{formaLabel(s.forma_pagamento)}</span></div>
            <div><span className="k">Parcelas</span><span className="v">{s.num_parcelas ? `${s.num_parcelas}x de ${fmtBRL(s.valor_parcela)}` : '—'}</span></div>
            <div><span className="k">Valor total</span><span className="v mono">{fmtBRL(s.valor_total)}</span></div>
          </div>
          <div style={{ marginTop: 10 }}><SeguroStatusBadge vigenciaFim={s.vigencia_fim} /></div>
        </div>
      ))}

      <button className="reports-link" onClick={onGoReport}><FileBarChart size={16} /> Relatório de seguros <ChevronRight size={15} /></button>

      <style jsx>{`
        .reports-link { display:flex; align-items:center; justify-content:center; gap:8px; width:100%; background:var(--surface); border:1px solid var(--border); color:var(--accent); font-weight:600; font-size:13px; padding:13px; border-radius:10px; cursor:pointer; margin-top:4px; }
      `}</style>
      <ViewStyle />
    </div>
  );
}

function SeguroForm({ item, empresas, onClose, onSave }) {
  const [empresaId, setEmpresaId] = useState(item?.empresaId || empresas[0]?.id || '');
  const [objeto, setObjeto] = useState(item?.objeto || '');
  const [seguradora, setSeguradora] = useState(item?.seguradora || '');
  const [principalCondutor, setPrincipalCondutor] = useState(item?.principalCondutor || '');
  const [vigenciaInicio, setVigenciaInicio] = useState(item?.vigenciaInicio || todayISO());
  const [vigenciaFim, setVigenciaFim] = useState(item?.vigenciaFim || todayISO());
  const [formaPagamento, setFormaPagamento] = useState(item?.formaPagamento || 'boleto');
  const [numParcelas, setNumParcelas] = useState(item?.numParcelas ?? '');
  const [valorParcela, setValorParcela] = useState(item?.valorParcela ?? '');
  const [valorTotal, setValorTotal] = useState(item?.valorTotal ?? '');

  const canSave = empresaId && objeto.trim() && seguradora.trim() && vigenciaInicio && vigenciaFim && valorTotal;

  return (
    <Modal title={item ? 'Editar seguro' : 'Novo seguro'} onClose={onClose}>
      {empresas.length === 0 ? <p className="muted2">Cadastre uma empresa antes.</p> : (
        <>
          <Field label="Empresa"><select className="select" value={empresaId} onChange={(e) => setEmpresaId(e.target.value)}>{empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}</select></Field>
          <Field label="Objeto do seguro"><input className="input" value={objeto} onChange={(e) => setObjeto(e.target.value)} placeholder="Ex: Veículo Fusca ABC-1234, Prédio sede..." /></Field>
          <Field label="Seguradora"><input className="input" value={seguradora} onChange={(e) => setSeguradora(e.target.value)} placeholder="Ex: Porto Seguro, Bradesco Seguros..." /></Field>
          <Field label="Principal condutor (opcional)"><input className="input" value={principalCondutor} onChange={(e) => setPrincipalCondutor(e.target.value)} placeholder="Nome do condutor" /></Field>
          <Field label="Início da vigência"><input className="input" type="date" value={vigenciaInicio} onChange={(e) => setVigenciaInicio(e.target.value)} /></Field>
          <Field label="Fim da vigência"><input className="input" type="date" value={vigenciaFim} onChange={(e) => setVigenciaFim(e.target.value)} /></Field>
          <Field label="Forma de pagamento">
            <div className="tipo-toggle">
              {FORMAS_PAGAMENTO_SEGURO.map((f) => (
                <button type="button" key={f.value} className={formaPagamento === f.value ? 'active' : ''} onClick={() => setFormaPagamento(f.value)}>{f.label}</button>
              ))}
            </div>
          </Field>
          <Field label="Quantidade de parcelas"><input className="input" type="number" value={numParcelas} onChange={(e) => setNumParcelas(e.target.value)} placeholder="Ex: 12" /></Field>
          <Field label="Valor da parcela (R$)"><DecimalMaskInput value={valorParcela} onChange={setValorParcela} /></Field>
          <Field label="Valor total do seguro (R$)"><DecimalMaskInput value={valorTotal} onChange={setValorTotal} /></Field>
          <PrimaryButton full onClick={() => canSave && onSave({
            id: item?.id, empresaId, objeto: objeto.trim(), seguradora: seguradora.trim(), principalCondutor: principalCondutor.trim(),
            vigenciaInicio, vigenciaFim, formaPagamento, numParcelas: parseInt(numParcelas, 10) || null,
            valorParcela: parseFloat(valorParcela) || null, valorTotal: parseFloat(valorTotal) || 0,
          })}>{item ? 'Salvar alterações' : 'Salvar seguro'}</PrimaryButton>
        </>
      )}
      <style jsx>{inputCss}</style>
      <style jsx>{`
        .tipo-toggle { display:flex; gap:6px; }
        .tipo-toggle button { flex:1; background:var(--surface); border:1px solid var(--border); color:var(--text-muted); font-size:12.5px; font-weight:600; padding:9px 8px; border-radius:8px; cursor:pointer; }
        .tipo-toggle button.active { background:var(--accent); border-color:var(--accent); color:var(--accent-contrast); }
      `}</style>
    </Modal>
  );
}

// ---------- Vendas ----------

function setorLabel(v) { return SETORES_VENDA.find((s) => s.value === v)?.label || v; }

function VendasView({ data, onSaveMes, onGoReport }) {
  const [empresaId, setEmpresaId] = useState(data.empresas[0]?.id || '');
  const [setor, setSetor] = useState(SETORES_VENDA[0].value);
  const [ano, setAno] = useState(new Date().getFullYear());
  const [editMes, setEditMes] = useState(null); // número do mês (1-12) sendo editado

  if (!data.empresas.length) return <EmptyState icon={ShoppingCart} text="Cadastre uma empresa antes de lançar vendas." />;

  const registros = (data.vendas || []).filter((v) => v.empresa_id === empresaId && v.setor === setor && v.ano === ano);
  const porMes = (mes) => registros.find((v) => v.mes === mes);
  const totalAno = registros.reduce((s, v) => s + (Number(v.valor) || 0), 0);

  const chartData = MESES.map((nome, i) => ({ name: nome.slice(0, 3), valor: Number(porMes(i + 1)?.valor) || 0 }));

  return (
    <div className="view">
      <div className="vendas-filtros">
        <select className="select" value={empresaId} onChange={(e) => setEmpresaId(e.target.value)}>
          {data.empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}
        </select>
        <select className="select" value={ano} onChange={(e) => setAno(parseInt(e.target.value, 10))}>
          {anosDisponiveis().map((y) => <option key={y} value={y}>{y}</option>)}
        </select>
      </div>

      <div className="seg-filter setor-filter">
        {SETORES_VENDA.map((s) => (
          <button key={s.value} className={setor === s.value ? 'active' : ''} onClick={() => setSetor(s.value)}>{s.label}</button>
        ))}
      </div>

      <div className="stat-card total-ano">
        <div className="stat-label">Total de {setorLabel(setor)} em {ano}</div>
        <div className="stat-value">{fmtBRL(totalAno)}</div>
      </div>

      <div className="chart-box">
        <ResponsiveContainer width="100%" height={160}>
          <BarChart data={chartData} margin={{ top: 4, right: 12, bottom: 4, left: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" vertical={false} />
            <XAxis dataKey="name" tick={{ fill: 'var(--text-dim)', fontSize: 10 }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fill: 'var(--text-dim)', fontSize: 10 }} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} axisLine={false} tickLine={false} width={34} />
            <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, fontSize: 12 }} labelStyle={{ color: 'var(--text)' }} />
            <Bar dataKey="valor" fill="var(--accent)" radius={[4, 4, 0, 0]} barSize={16} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <SectionTitle>Lançamentos mensais</SectionTitle>
      <div className="meses-list">
        {MESES.map((nome, i) => {
          const mes = i + 1;
          const reg = porMes(mes);
          return (
            <button key={mes} className="mes-row" onClick={() => setEditMes(mes)}>
              <span className="mes-nome">{nome}</span>
              <span className="mes-valor mono">{reg ? fmtBRL(reg.valor) : '—'}</span>
              {reg?.observacoes && <span className="mes-obs-dot" title={reg.observacoes} />}
              <Pencil size={14} color="var(--text-dim)" />
            </button>
          );
        })}
      </div>

      <button className="reports-link" onClick={onGoReport}><FileBarChart size={16} /> Relatório de vendas <ChevronRight size={15} /></button>

      {editMes && (
        <VendaMesForm
          empresaId={empresaId} setor={setor} ano={ano} mes={editMes}
          registro={porMes(editMes)}
          onClose={() => setEditMes(null)}
          onSave={(v) => { onSaveMes(v); setEditMes(null); }}
        />
      )}

      <style jsx>{`
        .vendas-filtros { display:flex; gap:8px; margin-bottom:10px; }
        .seg-filter { display:flex; gap:6px; margin-bottom:14px; flex-wrap:wrap; }
        .seg-filter button { flex:1 1 30%; min-width:90px; background:var(--surface); border:1px solid var(--border); color:var(--text-muted); font-size:12px; font-weight:600; padding:9px 6px; border-radius:8px; cursor:pointer; }
        .seg-filter button.active { background:var(--accent); border-color:var(--accent); color:var(--accent-contrast); }
        .total-ano { margin-bottom:12px; }
        .meses-list { display:flex; flex-direction:column; gap:8px; margin-bottom:8px; }
        .mes-row { display:flex; align-items:center; gap:10px; background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:12px 14px; cursor:pointer; text-align:left; box-shadow:0 1px 2px rgba(27,39,51,0.04); }
        .mes-nome { flex:1; font-size:13.5px; color:var(--text); font-weight:500; }
        .mes-valor { font-size:13.5px; color:var(--text-muted); }
        .mes-obs-dot { width:6px; height:6px; border-radius:50%; background:var(--accent); flex-shrink:0; }
        .reports-link { display:flex; align-items:center; justify-content:center; gap:8px; width:100%; background:var(--surface); border:1px solid var(--border); color:var(--accent); font-weight:600; font-size:13px; padding:13px; border-radius:10px; cursor:pointer; margin-top:8px; }
      `}</style>
      <style jsx>{inputCss}</style>
      <ViewStyle />
    </div>
  );
}

function VendaMesForm({ empresaId, setor, ano, mes, registro, onClose, onSave }) {
  const [valor, setValor] = useState(registro?.valor ?? '');
  const [observacoes, setObservacoes] = useState(registro?.observacoes || '');

  return (
    <Modal title={`${MESES[mes - 1]} de ${ano} · ${setorLabel(setor)}`} onClose={onClose}>
      <Field label="Valor total do mês (R$)"><DecimalMaskInput value={valor} onChange={setValor} /></Field>
      <Field label="Observações (opcional)">
        <textarea className="input textarea" value={observacoes} onChange={(e) => setObservacoes(e.target.value)} placeholder="Alguma observação sobre este mês..." rows={3} />
      </Field>
      <PrimaryButton full onClick={() => onSave({ empresaId, setor, ano, mes, valor: parseFloat(valor) || 0, observacoes: observacoes.trim() })}>Salvar lançamento</PrimaryButton>
      <style jsx>{inputCss}</style>
      <style jsx>{`.textarea { resize:vertical; font-family:'Inter',sans-serif; }`}</style>
    </Modal>
  );
}

// ---------- Taxas ----------

function TaxasHome({ data, sub, setSub, onNewCartao, onEditCartao, onDeleteCartao, onNewBoleto, onEditBoleto, onDeleteBoleto }) {
  if (sub === 'cartoes') return <CartoesView data={data} onEdit={onEditCartao} onNew={onNewCartao} onDelete={onDeleteCartao} onBack={() => setSub(null)} />;
  if (sub === 'boletos') return <BoletosView data={data} onEdit={onEditBoleto} onNew={onNewBoleto} onDelete={onDeleteBoleto} onBack={() => setSub(null)} />;
  const options = [
    { key: 'cartoes', label: 'Taxas de cartão de crédito', desc: 'Pix, débito, crédito à vista e parcelado por bandeira', icon: CreditCard, count: data.taxasCartao.length },
    { key: 'boletos', label: 'Taxas de boleto', desc: 'Emissão, baixa, protesto e antecipação', icon: Receipt, count: data.taxasBoleto.length },
  ];
  return (
    <div className="reports-home">
      {options.map(({ key, label, desc, icon: Icon, count }) => (
        <button key={key} className="report-option" onClick={() => setSub(key)}>
          <div className="report-icon"><Icon size={18} /></div>
          <div className="report-text"><div className="report-label">{label}</div><div className="report-desc">{desc}{count > 0 ? ` · ${count} cadastrada${count === 1 ? '' : 's'}` : ''}</div></div>
          <ChevronRight size={18} color="var(--text-dim)" />
        </button>
      ))}
      <style jsx>{`
        .reports-home { padding:16px 20px 24px; display:flex; flex-direction:column; gap:10px; }
        .report-option { display:flex; align-items:center; gap:12px; background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:14px; cursor:pointer; text-align:left; box-shadow:0 1px 2px rgba(27,39,51,0.04); }
        .report-icon { width:38px; height:38px; border-radius:10px; background:var(--accent-soft-12); color:var(--accent); display:flex; align-items:center; justify-content:center; flex-shrink:0; }
        .report-text { flex:1; }
        .report-label { font-family:'Space Grotesk',sans-serif; font-size:14.5px; color:var(--text); font-weight:600; }
        .report-desc { font-size:12px; color:var(--text-dim); margin-top:2px; }
      `}</style>
    </div>
  );
}

function ReportHeader({ title, onBack }) {
  return (
    <div className="rh">
      <button className="back" onClick={onBack}><ArrowLeft size={18} /></button>
      <h2>{title}</h2>
      <style jsx>{`
        .rh { display:flex; align-items:center; gap:10px; padding:16px 20px 4px; }
        .back { background:none; border:none; color:var(--text-muted); cursor:pointer; padding:4px; display:flex; }
        h2 { font-family:'Space Grotesk',sans-serif; font-size:17px; color:var(--text); font-weight:600; margin:0; }
      `}</style>
    </div>
  );
}

function CartoesView({ data, onEdit, onNew, onDelete, onBack }) {
  const [bandeiraFiltro, setBandeiraFiltro] = useState('todas');
  const [modalidade, setModalidade] = useState('avista');
  const [expandedId, setExpandedId] = useState(null);
  const lista = data.taxasCartao.filter((t) => bandeiraFiltro === 'todas' || t.bandeira === bandeiraFiltro);
  const valorModalidade = (t) => modalidade === 'pix' ? t.taxa_pix : modalidade === 'debito' ? t.taxa_debito : modalidade === 'avista' ? t.taxa_credito_avista : t.parcelas?.[modalidade];
  const comparativo = [...lista].filter((t) => valorModalidade(t) != null && valorModalidade(t) !== '')
    .map((t) => ({ name: t.administradora, valor: Number(valorModalidade(t)) || 0 })).sort((a, b) => a.valor - b.valor);

  const doExport = () => exportToExcel(lista, [
    { header: 'Administradora', get: (t) => t.administradora, width: 22 },
    { header: 'Bandeira', get: (t) => t.bandeira, width: 16 },
    { header: 'Pix', get: (t) => Number(t.taxa_pix) || 0, width: 10 },
    { header: 'Débito', get: (t) => Number(t.taxa_debito) || 0, width: 10 },
    { header: 'Crédito à vista', get: (t) => Number(t.taxa_credito_avista) || 0, width: 14 },
    ...PARCELAS_RANGE.map((n) => ({ header: `${n}x`, get: (t) => Number(t.parcelas?.[n]) || 0, width: 8 })),
  ], 'Taxas de Cartão', 'taxas-cartao');

  return (
    <div className="report-body">
      <ReportHeader title="Taxas de cartão de crédito" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10 }}>
        <select className="select" value={bandeiraFiltro} onChange={(e) => setBandeiraFiltro(e.target.value)}>
          <option value="todas">Todas as bandeiras</option>
          {BANDEIRAS.map((b) => <option key={b} value={b}>{b}</option>)}
        </select>
        <button className="add-inline-btn" onClick={onNew}><Plus size={15} /> Nova taxa</button>
      </div>
      {lista.length === 0 ? <EmptyState icon={CreditCard} text="Nenhuma taxa cadastrada." /> : (
        <>
          <div className="view" style={{ padding: '0 20px 8px' }}>
            {lista.map((t) => {
              const expanded = expandedId === t.id;
              return (
                <div key={t.id} className="card">
                  <div className="card-head">
                    <div className="card-title"><CreditCard size={16} /> {t.administradora} <span className="badge">{t.bandeira}</span></div>
                    <div className="card-actions">
                      <button className="icon-btn" onClick={() => onEdit({ id: t.id, administradora: t.administradora, bandeira: t.bandeira, taxaPix: t.taxa_pix, taxaDebito: t.taxa_debito, taxaCreditoAvista: t.taxa_credito_avista, parcelas: t.parcelas })}><Pencil size={16} /></button>
                      <button className="icon-btn danger" onClick={() => onDelete(t.id, `${t.administradora} (${t.bandeira})`)}><Trash2 size={16} /></button>
                    </div>
                  </div>
                  <div className="card-grid">
                    <div><span className="k">Pix</span><span className="v mono">{fmtPct(t.taxa_pix)}</span></div>
                    <div><span className="k">Débito</span><span className="v mono">{fmtPct(t.taxa_debito)}</span></div>
                    <div><span className="k">Crédito à vista</span><span className="v mono">{fmtPct(t.taxa_credito_avista)}</span></div>
                  </div>
                  <button className="expand-btn2" onClick={() => setExpandedId(expanded ? null : t.id)}>
                    {expanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />} {expanded ? 'Ocultar parcelado' : 'Ver crédito parcelado (2x-12x)'}
                  </button>
                  {expanded && (
                    <div className="parcelas-taxa-grid">
                      {PARCELAS_RANGE.map((n) => (
                        <div key={n} className="pti"><span className="k">{n}x</span><span className="v mono">{fmtPct(t.parcelas?.[n])}</span></div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
          <SectionTitle>Comparativo entre administradoras</SectionTitle>
          <div className="report-filters" style={{ marginBottom: 8 }}>
            <select className="select" value={modalidade} onChange={(e) => setModalidade(e.target.value)}>
              <option value="pix">Pix</option>
              <option value="debito">Débito</option>
              <option value="avista">Crédito à vista</option>
              {PARCELAS_RANGE.map((n) => <option key={n} value={n}>Crédito {n}x</option>)}
            </select>
          </div>
          {comparativo.length === 0 ? <p className="muted2">Nenhuma taxa cadastrada para essa modalidade.</p> : (
            <div className="chart-box" style={{ margin: '0 20px' }}>
              <ResponsiveContainer width="100%" height={Math.max(120, comparativo.length * 40)}>
                <BarChart data={comparativo} layout="vertical" margin={{ top: 4, right: 30, bottom: 4, left: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" horizontal={false} />
                  <XAxis type="number" tick={{ fill: 'var(--text-dim)', fontSize: 10 }} tickFormatter={(v) => `${v}%`} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" tick={{ fill: 'var(--text-muted)', fontSize: 11 }} axisLine={false} tickLine={false} width={90} />
                  <Tooltip formatter={(v) => fmtPct(v)} contentStyle={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, fontSize: 12 }} labelStyle={{ color: 'var(--text)' }} />
                  <Bar dataKey="valor" fill="var(--accent)" radius={[0, 4, 4, 0]} barSize={16} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
          <div style={{ padding: '14px 20px 4px' }}><ExportButton onClick={doExport} disabled={lista.length === 0} label="Exportar comparativo (Excel)" /></div>
        </>
      )}
      <style jsx>{`
        .expand-btn2 { margin-top:10px; display:inline-flex; align-items:center; gap:6px; background:none; border:1px solid var(--border); border-radius:8px; padding:7px 12px; font-size:12px; font-weight:600; color:var(--text-muted); cursor:pointer; }
        .parcelas-taxa-grid { margin-top:10px; display:grid; grid-template-columns:repeat(4,1fr); gap:8px; }
        .pti { background:var(--surface-alt); border:1px solid var(--border-strong); border-radius:9px; padding:8px; display:flex; flex-direction:column; gap:2px; }
        .pti .k { font-size:10px; } .pti .v { font-size:12.5px; }
        .muted2 { color:var(--text-dim); font-size:14px; padding:0 20px 8px; }
      `}</style>
      <ViewStyle />
    </div>
  );
}

function BoletosView({ data, onEdit, onNew, onDelete, onBack }) {
  const [modalidade, setModalidade] = useState('emissao');
  const valorModalidade = (t) => modalidade === 'emissao' ? t.taxa_emissao : modalidade === 'baixa' ? t.taxa_baixa : modalidade === 'protesto' ? t.taxa_protesto : t.taxa_antecipacao;
  const comparativo = [...data.taxasBoleto].filter((t) => valorModalidade(t) != null && valorModalidade(t) !== '')
    .map((t) => ({ name: t.administradora, valor: Number(valorModalidade(t)) || 0 })).sort((a, b) => a.valor - b.valor);
  const isAntecipacao = modalidade === 'antecipacao';

  const doExport = () => exportToExcel(data.taxasBoleto, [
    { header: 'Administradora', get: (t) => t.administradora, width: 24 },
    { header: 'Emissão (R$)', get: (t) => Number(t.taxa_emissao) || 0, width: 14 },
    { header: 'Baixa (R$)', get: (t) => Number(t.taxa_baixa) || 0, width: 14 },
    { header: 'Protesto (R$)', get: (t) => Number(t.taxa_protesto) || 0, width: 14 },
    { header: 'Antecipação (% a.m.)', get: (t) => Number(t.taxa_antecipacao) || 0, width: 18 },
  ], 'Taxas de Boleto', 'taxas-boleto');

  return (
    <div className="report-body">
      <ReportHeader title="Taxas de boleto" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10, justifyContent: 'flex-end' }}>
        <button className="add-inline-btn" onClick={onNew}><Plus size={15} /> Nova taxa</button>
      </div>
      {data.taxasBoleto.length === 0 ? <EmptyState icon={Receipt} text="Nenhuma taxa de boleto cadastrada." /> : (
        <>
          <div className="view" style={{ padding: '0 20px 8px' }}>
            {data.taxasBoleto.map((t) => (
              <div key={t.id} className="card">
                <div className="card-head">
                  <div className="card-title"><Receipt size={16} /> {t.administradora}</div>
                  <div className="card-actions">
                    <button className="icon-btn" onClick={() => onEdit({ id: t.id, administradora: t.administradora, taxaEmissao: t.taxa_emissao, taxaBaixa: t.taxa_baixa, taxaProtesto: t.taxa_protesto, taxaAntecipacao: t.taxa_antecipacao })}><Pencil size={16} /></button>
                    <button className="icon-btn danger" onClick={() => onDelete(t.id, t.administradora)}><Trash2 size={16} /></button>
                  </div>
                </div>
                <div className="card-grid">
                  <div><span className="k">Emissão</span><span className="v mono">{fmtBRL(t.taxa_emissao)}</span></div>
                  <div><span className="k">Baixa</span><span className="v mono">{fmtBRL(t.taxa_baixa)}</span></div>
                  <div><span className="k">Protesto</span><span className="v mono">{fmtBRL(t.taxa_protesto)}</span></div>
                  <div><span className="k">Antecipação</span><span className="v mono">{fmtPct(t.taxa_antecipacao)} a.m.</span></div>
                </div>
              </div>
            ))}
          </div>
          <SectionTitle>Comparativo entre administradoras</SectionTitle>
          <div className="report-filters" style={{ marginBottom: 8 }}>
            <select className="select" value={modalidade} onChange={(e) => setModalidade(e.target.value)}>
              <option value="emissao">Taxa de emissão</option>
              <option value="baixa">Taxa de baixa</option>
              <option value="protesto">Taxa de protesto</option>
              <option value="antecipacao">Taxa de antecipação (% a.m.)</option>
            </select>
          </div>
          {comparativo.length === 0 ? <p className="muted2">Nenhuma taxa cadastrada para essa modalidade.</p> : (
            <div className="chart-box" style={{ margin: '0 20px' }}>
              <ResponsiveContainer width="100%" height={Math.max(120, comparativo.length * 40)}>
                <BarChart data={comparativo} layout="vertical" margin={{ top: 4, right: 30, bottom: 4, left: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" horizontal={false} />
                  <XAxis type="number" tick={{ fill: 'var(--text-dim)', fontSize: 10 }} tickFormatter={(v) => isAntecipacao ? `${v}%` : `R$${v}`} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" tick={{ fill: 'var(--text-muted)', fontSize: 11 }} axisLine={false} tickLine={false} width={90} />
                  <Tooltip formatter={(v) => isAntecipacao ? fmtPct(v) : fmtBRL(v)} contentStyle={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, fontSize: 12 }} labelStyle={{ color: 'var(--text)' }} />
                  <Bar dataKey="valor" fill="var(--accent)" radius={[0, 4, 4, 0]} barSize={16} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
          <div style={{ padding: '14px 20px 4px' }}><ExportButton onClick={doExport} disabled={data.taxasBoleto.length === 0} label="Exportar comparativo (Excel)" /></div>
        </>
      )}
      <style jsx>{`.muted2 { color:var(--text-dim); font-size:14px; padding:0 20px 8px; }`}</style>
      <ViewStyle />
    </div>
  );
}

// ---------- Reports ----------

function ReportsHome({ view, setView, data }) {
  if (view === 'bancos') return <ReportSaldosBancos data={data} onBack={() => setView(null)} />;
  if (view === 'empresas') return <ReportSaldosEmpresas data={data} onBack={() => setView(null)} />;
  if (view === 'contas') return <ReportContasPagar data={data} onBack={() => setView(null)} />;
  if (view === 'pagamentos') return <ReportPagamentos data={data} onBack={() => setView(null)} />;
  if (view === 'emprestimos-empresa') return <ReportEmprestimosPorEmpresa data={data} onBack={() => setView(null)} />;
  if (view === 'parcelas-emprestimo') return <ReportParcelasEmprestimo data={data} onBack={() => setView(null)} />;
  if (view === 'seguros') return <ReportSeguros data={data} onBack={() => setView(null)} />;
  if (view === 'vendas') return <ReportVendas data={data} onBack={() => setView(null)} />;
  const options = [
    { key: 'bancos', label: 'Saldos por banco', desc: 'Total consolidado em cada conta bancária', icon: Landmark },
    { key: 'empresas', label: 'Saldos por empresa', desc: 'Total consolidado por empresa do grupo', icon: Building2 },
    { key: 'contas', label: 'Contas a pagar', desc: 'Pendentes e atrasadas, filtráveis por empresa', icon: CalendarClock },
    { key: 'pagamentos', label: 'Pagamentos realizados', desc: 'Histórico de contas já pagas (baixas)', icon: CheckCircle2 },
    { key: 'emprestimos-empresa', label: 'Empréstimos por empresa', desc: 'Bancários e entre empresas, agrupados por devedora', icon: HandCoins },
    { key: 'parcelas-emprestimo', label: 'Parcelas de empréstimos', desc: 'Pendentes e pagas, filtráveis por tipo e empresa', icon: CalendarClock },
    { key: 'seguros', label: 'Seguros', desc: 'Vigências, vencimentos e valores por empresa', icon: Shield },
    { key: 'vendas', label: 'Vendas por setor', desc: 'Totais mensais e anuais desde 2018, por setor e empresa', icon: ShoppingCart },
  ];
  return (
    <div className="reports-home">
      {options.map(({ key, label, desc, icon: Icon }) => (
        <button key={key} className="report-option" onClick={() => setView(key)}>
          <div className="report-icon"><Icon size={18} /></div>
          <div className="report-text"><div className="report-label">{label}</div><div className="report-desc">{desc}</div></div>
          <ChevronRight size={18} color="var(--text-dim)" />
        </button>
      ))}
      <style jsx>{`
        .reports-home { padding:16px 20px 24px; display:flex; flex-direction:column; gap:10px; }
        .report-option { display:flex; align-items:center; gap:12px; background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:14px; cursor:pointer; text-align:left; box-shadow:0 1px 2px rgba(27,39,51,0.04); }
        .report-icon { width:38px; height:38px; border-radius:10px; background:var(--accent-soft-12); color:var(--accent); display:flex; align-items:center; justify-content:center; flex-shrink:0; }
        .report-text { flex:1; }
        .report-label { font-family:'Space Grotesk',sans-serif; font-size:14.5px; color:var(--text); font-weight:600; }
        .report-desc { font-size:12px; color:var(--text-dim); margin-top:2px; }
      `}</style>
    </div>
  );
}

function ExportRow({ total, count, exportExcel, exportPdf }) {
  return (
    <div className="rtotal">
      {count != null ? `Total (${count}): ` : 'Total: '}<span className="mono">{fmtBRL(total)}</span>
      <div className="rbuttons">
        <ExportButton disabled={!exportExcel} onClick={exportExcel} label="Excel" />
        <ExportButton disabled={!exportPdf} onClick={exportPdf} label="PDF" />
      </div>
    </div>
  );
}

function ReportSaldosBancos({ data, onBack }) {
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const total = data.bancos.reduce((s, b) => s + (Number(b.saldo) || 0), 0);
  const sorted = [...data.bancos].sort((a, b) => (Number(b.saldo) || 0) - (Number(a.saldo) || 0));
  const columns = [
    { header: 'Banco', get: (b) => b.nome_banco, width: 24 },
    { header: 'Saldo', get: (b) => Number(b.saldo) || 0, width: 16 },
  ];
  const groups = groupRows(sorted, (b) => b.empresa_id, (b) => empresaNome(b.empresa_id))
    .sort((a, b) => a.label.localeCompare(b.label));

  return (
    <div className="report-body">
      <ReportHeader title="Saldos por banco" onBack={onBack} />
      <ExportRow
        total={total}
        exportExcel={sorted.length ? () => exportGroupedExcel({ groups: groups.map((g) => ({ ...g, label: `Empresa: ${g.label}` })), columns, valueColIndex: 1, sheetName: 'Saldos por Banco', fileName: 'saldos-por-banco', reportTitle: 'Saldos por Banco' }) : null}
        exportPdf={sorted.length ? () => exportToPDF({ title: 'Saldos por Banco', groups: groups.map((g) => ({ ...g, label: `Empresa: ${g.label}` })), columns, valueColIndex: 1, fileName: 'saldos-por-banco' }) : null}
      />
      {sorted.length === 0 ? <EmptyState icon={Landmark} text="Nenhum banco cadastrado." /> : (
        <div className="grouped-report">
          {groups.map((g) => {
            const subtotal = g.rows.reduce((s, b) => s + (Number(b.saldo) || 0), 0);
            return (
              <div key={g.label} className="group-block">
                <div className="group-head"><Building2 size={14} /> {g.label}</div>
                <table className="rtable"><thead><tr><th>Banco</th><th className="right">Saldo</th></tr></thead>
                  <tbody>
                    {g.rows.map((b) => <tr key={b.id}><td>{b.nome_banco}</td><td className="right mono">{fmtBRL(b.saldo)}</td></tr>)}
                    <tr className="subtotal-row"><td>Subtotal</td><td className="right mono">{fmtBRL(subtotal)}</td></tr>
                  </tbody>
                </table>
              </div>
            );
          })}
        </div>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function ReportSaldosEmpresas({ data, onBack }) {
  const total = data.bancos.reduce((s, b) => s + (Number(b.saldo) || 0), 0);
  const rows = data.empresas.map((emp) => ({ nome: emp.nome, saldo: data.bancos.filter((b) => b.empresa_id === emp.id).reduce((s, b) => s + (Number(b.saldo) || 0), 0), numBancos: data.bancos.filter((b) => b.empresa_id === emp.id).length })).sort((a, b) => b.saldo - a.saldo);
  const columns = [
    { header: 'Empresa', get: (r) => r.nome, width: 24 },
    { header: 'Bancos', get: (r) => r.numBancos, width: 12 },
    { header: 'Saldo', get: (r) => r.saldo, width: 16 },
  ];
  const groups = [{ label: 'Todas as empresas', rows }];

  return (
    <div className="report-body">
      <ReportHeader title="Saldos por empresa" onBack={onBack} />
      <ExportRow
        total={total}
        exportExcel={rows.length ? () => exportGroupedExcel({ groups, columns, valueColIndex: 2, sheetName: 'Saldos por Empresa', fileName: 'saldos-por-empresa', reportTitle: 'Saldos por Empresa' }) : null}
        exportPdf={rows.length ? () => exportToPDF({ title: 'Saldos por Empresa', groups, columns, valueColIndex: 2, fileName: 'saldos-por-empresa' }) : null}
      />
      {rows.length === 0 ? <EmptyState icon={Building2} text="Nenhuma empresa cadastrada." /> : (
        <table className="rtable"><thead><tr><th>Empresa</th><th>Bancos</th><th className="right">Saldo</th></tr></thead>
          <tbody>{rows.map((r) => <tr key={r.nome}><td>{r.nome}</td><td className="dim">{r.numBancos}</td><td className="right mono">{fmtBRL(r.saldo)}</td></tr>)}</tbody>
        </table>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function PeriodoFiltro({ de, ate, onDe, onAte }) {
  return (
    <div className="periodo-filtro">
      <label><span>De</span><input className="input" type="date" value={de} onChange={(e) => onDe(e.target.value)} /></label>
      <label><span>Até</span><input className="input" type="date" value={ate} onChange={(e) => onAte(e.target.value)} /></label>
      <style jsx>{`
        .periodo-filtro { display:flex; gap:8px; flex:1 1 100%; }
        label { display:flex; flex-direction:column; gap:4px; flex:1; }
        span { font-size:11px; color:var(--text-dim); }
      `}</style>
    </div>
  );
}

function ReportContasPagar({ data, onBack }) {
  const [empresaFiltro, setEmpresaFiltro] = useState('todas');
  const [statusFiltro, setStatusFiltro] = useState('pendente');
  const [periodoDe, setPeriodoDe] = useState('');
  const [periodoAte, setPeriodoAte] = useState('');
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const filtradas = useMemo(() => data.contas
    .filter((c) => empresaFiltro === 'todas' || c.empresa_id === empresaFiltro)
    .filter((c) => statusFiltro === 'todas' || c.status === statusFiltro)
    .filter((c) => !periodoDe || c.data_vencimento >= periodoDe)
    .filter((c) => !periodoAte || c.data_vencimento <= periodoAte)
    .sort((a, b) => a.data_vencimento.localeCompare(b.data_vencimento)), [data.contas, empresaFiltro, statusFiltro, periodoDe, periodoAte]);
  const total = filtradas.reduce((s, c) => s + (Number(c.valor) || 0), 0);
  const columns = [
    { header: 'Descrição', get: (c) => c.descricao, width: 28 },
    { header: 'Vencimento', get: (c) => fmtDate(c.data_vencimento), width: 14 },
    { header: 'Status', get: (c) => c.status === 'pago' ? 'Pago' : (daysUntil(c.data_vencimento) < 0 ? 'Atrasada' : 'Pendente'), width: 12 },
    { header: 'Valor', get: (c) => Number(c.valor) || 0, width: 16 },
  ];
  const groups = groupRows(filtradas, (c) => c.empresa_id, (c) => `Empresa: ${empresaNome(c.empresa_id)}`);

  return (
    <div className="report-body">
      <ReportHeader title="Contas a pagar" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10 }}>
        <select className="select" value={empresaFiltro} onChange={(e) => setEmpresaFiltro(e.target.value)}>
          <option value="todas">Todas as empresas</option>
          {data.empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}
        </select>
        <select className="select" value={statusFiltro} onChange={(e) => setStatusFiltro(e.target.value)}>
          <option value="pendente">Pendentes</option><option value="pago">Pagas</option><option value="todas">Todas</option>
        </select>
        <PeriodoFiltro de={periodoDe} ate={periodoAte} onDe={setPeriodoDe} onAte={setPeriodoAte} />
      </div>
      <ExportRow
        total={total} count={filtradas.length}
        exportExcel={filtradas.length ? () => exportGroupedExcel({ groups, columns, valueColIndex: 3, sheetName: 'Contas a Pagar', fileName: 'contas-a-pagar', reportTitle: 'Contas a Pagar' }) : null}
        exportPdf={filtradas.length ? () => exportToPDF({ title: 'Contas a Pagar', groups, columns, valueColIndex: 3, fileName: 'contas-a-pagar' }) : null}
      />
      {filtradas.length === 0 ? <EmptyState icon={CalendarClock} text="Nenhuma conta encontrada." /> : (
        <table className="rtable"><thead><tr><th>Descrição</th><th>Empresa</th><th>Vencimento</th><th className="right">Valor</th></tr></thead>
          <tbody>{filtradas.map((c) => { const atrasada = c.status === 'pendente' && daysUntil(c.data_vencimento) < 0; return (
            <tr key={c.id}><td>{c.descricao}</td><td className="dim">{empresaNome(c.empresa_id)}</td><td style={{ color: atrasada ? 'var(--danger)' : undefined }}>{fmtDate(c.data_vencimento)}</td><td className="right mono">{fmtBRL(c.valor)}</td></tr>
          ); })}</tbody>
        </table>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function ReportPagamentos({ data, onBack }) {
  const [empresaFiltro, setEmpresaFiltro] = useState('todas');
  const [periodoDe, setPeriodoDe] = useState('');
  const [periodoAte, setPeriodoAte] = useState('');
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const pagas = useMemo(() => data.contas
    .filter((c) => c.status === 'pago')
    .filter((c) => empresaFiltro === 'todas' || c.empresa_id === empresaFiltro)
    .filter((c) => !periodoDe || (c.data_pagamento || '') >= periodoDe)
    .filter((c) => !periodoAte || (c.data_pagamento || '') <= periodoAte)
    .sort((a, b) => (b.data_pagamento || '').localeCompare(a.data_pagamento || '')), [data.contas, empresaFiltro, periodoDe, periodoAte]);
  const total = pagas.reduce((s, c) => s + (Number(c.valor) || 0), 0);
  const columns = [
    { header: 'Descrição', get: (c) => c.descricao, width: 28 },
    { header: 'Baixa em', get: (c) => fmtDate(c.data_pagamento), width: 14 },
    { header: 'Valor', get: (c) => Number(c.valor) || 0, width: 16 },
  ];
  const groups = groupRows(pagas, (c) => c.empresa_id, (c) => `Empresa: ${empresaNome(c.empresa_id)}`);

  return (
    <div className="report-body">
      <ReportHeader title="Contas pagas (baixas)" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10 }}>
        <select className="select" value={empresaFiltro} onChange={(e) => setEmpresaFiltro(e.target.value)}>
          <option value="todas">Todas as empresas</option>
          {data.empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}
        </select>
        <PeriodoFiltro de={periodoDe} ate={periodoAte} onDe={setPeriodoDe} onAte={setPeriodoAte} />
      </div>
      <ExportRow
        total={total} count={pagas.length}
        exportExcel={pagas.length ? () => exportGroupedExcel({ groups, columns, valueColIndex: 2, sheetName: 'Contas Pagas', fileName: 'contas-pagas', reportTitle: 'Contas Pagas (Baixas)' }) : null}
        exportPdf={pagas.length ? () => exportToPDF({ title: 'Contas Pagas (Baixas)', groups, columns, valueColIndex: 2, fileName: 'contas-pagas' }) : null}
      />
      {pagas.length === 0 ? <EmptyState icon={CheckCircle2} text="Nenhuma baixa registrada ainda." /> : (
        <table className="rtable"><thead><tr><th>Descrição</th><th>Empresa</th><th>Baixa em</th><th className="right">Valor</th></tr></thead>
          <tbody>{pagas.map((c) => <tr key={c.id}><td>{c.descricao}</td><td className="dim">{empresaNome(c.empresa_id)}</td><td>{fmtDate(c.data_pagamento)}</td><td className="right mono">{fmtBRL(c.valor)}</td></tr>)}</tbody>
        </table>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function dataFinalContrato(e) {
  const parcelas = e.parcelas || [];
  if (parcelas.length) return parcelas[parcelas.length - 1].data_vencimento;
  // fallback caso as parcelas ainda não tenham sido carregadas
  // "num_parcelas" já representa o prazo total do contrato em meses (inclui a carência)
  return addMonthsLocal(e.data_inicio, (Number(e.num_parcelas) || 1) - 1);
}

function ReportEmprestimosPorEmpresa({ data, onBack }) {
  const [tipoFiltro, setTipoFiltro] = useState('todos');
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const lista = data.emprestimos.filter((e) => tipoFiltro === 'todos' || e.tipo === tipoFiltro);
  const total = lista.reduce((s, e) => s + (Number(e.valor_total) || 0), 0);
  const columns = [
    { header: 'Tipo', get: (e) => (e.tipo || 'banco') === 'banco' ? 'Bancário' : 'Intercompany', width: 14 },
    { header: 'Credor / Empresa credora', get: (e) => (e.tipo || 'banco') === 'banco' ? e.credor : empresaNome(e.empresa_credora_id), width: 24 },
    { header: 'Linha de crédito', get: (e) => e.linha_credito || '—', width: 20 },
    { header: 'Valor da parcela', get: (e) => Number(e.valor_parcela) || 0, width: 16 },
    { header: 'Parcelas pagas', get: (e) => `${(e.parcelas || []).filter((p) => p.status === 'pago').length}/${(e.parcelas || []).length}`, width: 14 },
    { header: 'Carência', get: (e) => e.tem_carencia ? `${e.prazo_carencia} ${e.prazo_carencia === 1 ? 'mês' : 'meses'}` : 'Sem carência', width: 16 },
    { header: 'Data final do contrato', get: (e) => fmtDate(dataFinalContrato(e)), width: 18 },
    { header: 'Valor total', get: (e) => Number(e.valor_total) || 0, width: 16 },
  ];
  const groups = groupRows(lista, (e) => e.empresa_id, (e) => `Empresa devedora: ${empresaNome(e.empresa_id)}`);

  return (
    <div className="report-body">
      <ReportHeader title="Empréstimos por empresa" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10 }}>
        <select className="select" value={tipoFiltro} onChange={(e) => setTipoFiltro(e.target.value)}>
          <option value="todos">Todos os tipos</option>
          <option value="banco">Somente bancários</option>
          <option value="empresa">Somente entre empresas</option>
        </select>
      </div>
      <ExportRow
        total={total} count={lista.length}
        exportExcel={lista.length ? () => exportGroupedExcel({ groups, columns, valueColIndex: 7, sheetName: 'Empréstimos', fileName: 'emprestimos-por-empresa', reportTitle: 'Empréstimos por Empresa' }) : null}
        exportPdf={lista.length ? () => exportToPDF({ title: 'Empréstimos por Empresa', groups, columns, valueColIndex: 7, fileName: 'emprestimos-por-empresa' }) : null}
      />
      {lista.length === 0 ? <EmptyState icon={HandCoins} text="Nenhum empréstimo encontrado." /> : (
        <div className="grouped-report">
          {groups.map((g) => {
            const subtotal = g.rows.reduce((s, e) => s + (Number(e.valor_total) || 0), 0);
            return (
              <div key={g.label} className="group-block">
                <div className="group-head"><Building2 size={14} /> {g.label}</div>
                <table className="rtable"><thead><tr><th>Tipo</th><th>Credor/Credora</th><th className="right">Parcela</th><th>Parcelas</th><th>Carência</th><th>Final do contrato</th><th className="right">Valor total</th></tr></thead>
                  <tbody>
                    {g.rows.map((e) => (
                      <tr key={e.id}>
                        <td>{(e.tipo || 'banco') === 'banco' ? 'Bancário' : 'Intercompany'}</td>
                        <td className="dim">{(e.tipo || 'banco') === 'banco' ? e.credor : empresaNome(e.empresa_credora_id)}</td>
                        <td className="right mono">{fmtBRL(e.valor_parcela)}</td>
                        <td className="dim">{(e.parcelas || []).filter((p) => p.status === 'pago').length}/{(e.parcelas || []).length}</td>
                        <td className="dim">{e.tem_carencia ? `${e.prazo_carencia} ${e.prazo_carencia === 1 ? 'mês' : 'meses'}` : '—'}</td>
                        <td className="dim">{fmtDate(dataFinalContrato(e))}</td>
                        <td className="right mono">{fmtBRL(e.valor_total)}</td>
                      </tr>
                    ))}
                    <tr className="subtotal-row"><td colSpan={6}>Subtotal {g.label.replace('Empresa devedora: ', '')}</td><td className="right mono">{fmtBRL(subtotal)}</td></tr>
                  </tbody>
                </table>
              </div>
            );
          })}
          <div className="grand-total-row">
            <span>Total geral</span>
            <span className="mono">{fmtBRL(total)}</span>
          </div>
        </div>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function ReportParcelasEmprestimo({ data, onBack }) {
  const [empresaFiltro, setEmpresaFiltro] = useState('todas');
  const [tipoFiltro, setTipoFiltro] = useState('todos');
  const [statusFiltro, setStatusFiltro] = useState('pendente');
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';

  const linhas = useMemo(() => {
    const out = [];
    data.emprestimos
      .filter((e) => empresaFiltro === 'todas' || e.empresa_id === empresaFiltro)
      .filter((e) => tipoFiltro === 'todos' || e.tipo === tipoFiltro)
      .forEach((e) => {
        (e.parcelas || [])
          .filter((p) => statusFiltro === 'todas' || p.status === statusFiltro)
          .forEach((p) => out.push({ ...p, emprestimo: e }));
      });
    return out.sort((a, b) => a.data_vencimento.localeCompare(b.data_vencimento));
  }, [data.emprestimos, empresaFiltro, tipoFiltro, statusFiltro]);

  const total = linhas.reduce((s, p) => s + (Number(p.valor) || 0), 0);
  const columns = [
    { header: 'Credor/Credora', get: (p) => (p.emprestimo.tipo || 'banco') === 'banco' ? p.emprestimo.credor : empresaNome(p.emprestimo.empresa_credora_id), width: 24 },
    { header: 'Parcela', get: (p) => `${p.numero}ª`, width: 10 },
    { header: 'Vencimento', get: (p) => fmtDate(p.data_vencimento), width: 14 },
    { header: 'Status', get: (p) => p.status === 'pago' ? 'Pago' : (daysUntil(p.data_vencimento) < 0 ? 'Atrasada' : 'Pendente'), width: 12 },
    { header: 'Valor', get: (p) => Number(p.valor) || 0, width: 16 },
  ];
  const groups = groupRows(linhas, (p) => p.emprestimo.empresa_id, (p) => `Empresa devedora: ${empresaNome(p.emprestimo.empresa_id)}`);

  return (
    <div className="report-body">
      <ReportHeader title="Parcelas de empréstimos" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10 }}>
        <select className="select" value={empresaFiltro} onChange={(e) => setEmpresaFiltro(e.target.value)}>
          <option value="todas">Todas as empresas</option>
          {data.empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}
        </select>
        <select className="select" value={tipoFiltro} onChange={(e) => setTipoFiltro(e.target.value)}>
          <option value="todos">Todos os tipos</option>
          <option value="banco">Bancários</option>
          <option value="empresa">Entre empresas</option>
        </select>
        <select className="select" value={statusFiltro} onChange={(e) => setStatusFiltro(e.target.value)}>
          <option value="pendente">Pendentes</option><option value="pago">Pagas</option><option value="todas">Todas</option>
        </select>
      </div>
      <ExportRow
        total={total} count={linhas.length}
        exportExcel={linhas.length ? () => exportGroupedExcel({ groups, columns, valueColIndex: 4, sheetName: 'Parcelas', fileName: 'parcelas-emprestimos', reportTitle: 'Parcelas de Empréstimos' }) : null}
        exportPdf={linhas.length ? () => exportToPDF({ title: 'Parcelas de Empréstimos', groups, columns, valueColIndex: 4, fileName: 'parcelas-emprestimos' }) : null}
      />
      {linhas.length === 0 ? <EmptyState icon={CalendarClock} text="Nenhuma parcela encontrada." /> : (
        <div className="grouped-report">
          {groups.map((g) => {
            const subtotal = g.rows.reduce((s, p) => s + (Number(p.valor) || 0), 0);
            return (
              <div key={g.label} className="group-block">
                <div className="group-head"><Building2 size={14} /> {g.label}</div>
                <table className="rtable"><thead><tr><th>Credor/Credora</th><th>Parcela</th><th>Vencimento</th><th className="right">Valor</th></tr></thead>
                  <tbody>
                    {g.rows.map((p) => {
                      const atrasada = p.status === 'pendente' && daysUntil(p.data_vencimento) < 0;
                      return (
                        <tr key={p.id}>
                          <td className="dim">{(p.emprestimo.tipo || 'banco') === 'banco' ? p.emprestimo.credor : empresaNome(p.emprestimo.empresa_credora_id)}</td>
                          <td>{p.numero}ª</td>
                          <td style={{ color: atrasada ? 'var(--danger)' : undefined }}>{fmtDate(p.data_vencimento)}</td>
                          <td className="right mono">{fmtBRL(p.valor)}</td>
                        </tr>
                      );
                    })}
                    <tr className="subtotal-row"><td colSpan={3}>Subtotal {g.label.replace('Empresa devedora: ', '')}</td><td className="right mono">{fmtBRL(subtotal)}</td></tr>
                  </tbody>
                </table>
              </div>
            );
          })}
          <div className="grand-total-row">
            <span>Total geral</span>
            <span className="mono">{fmtBRL(total)}</span>
          </div>
        </div>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function ReportSeguros({ data, onBack }) {
  const [empresaFiltro, setEmpresaFiltro] = useState('todas');
  const [statusFiltro, setStatusFiltro] = useState('todos');
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const formaLabel = (v) => FORMAS_PAGAMENTO_SEGURO.find((f) => f.value === v)?.label || v;

  const lista = useMemo(() => (data.seguros || [])
    .filter((s) => empresaFiltro === 'todas' || s.empresa_id === empresaFiltro)
    .filter((s) => statusFiltro === 'todos' || statusSeguro(s.vigencia_fim) === statusFiltro)
    .sort((a, b) => a.vigencia_fim.localeCompare(b.vigencia_fim)), [data.seguros, empresaFiltro, statusFiltro]);

  const total = lista.reduce((s, x) => s + (Number(x.valor_total) || 0), 0);
  const columns = [
    { header: 'Objeto', get: (s) => s.objeto, width: 24 },
    { header: 'Seguradora', get: (s) => s.seguradora, width: 20 },
    { header: 'Vigência início', get: (s) => fmtDate(s.vigencia_inicio), width: 14 },
    { header: 'Vigência fim', get: (s) => fmtDate(s.vigencia_fim), width: 14 },
    { header: 'Condutor', get: (s) => s.principal_condutor || '—', width: 18 },
    { header: 'Pagamento', get: (s) => formaLabel(s.forma_pagamento), width: 14 },
    { header: 'Valor total', get: (s) => Number(s.valor_total) || 0, width: 16 },
  ];
  const groups = groupRows(lista, (s) => s.empresa_id, (s) => `Empresa: ${empresaNome(s.empresa_id)}`);

  return (
    <div className="report-body">
      <ReportHeader title="Seguros" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10 }}>
        <select className="select" value={empresaFiltro} onChange={(e) => setEmpresaFiltro(e.target.value)}>
          <option value="todas">Todas as empresas</option>
          {data.empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}
        </select>
        <select className="select" value={statusFiltro} onChange={(e) => setStatusFiltro(e.target.value)}>
          <option value="todos">Todos os status</option>
          <option value="vigente">Vigentes</option>
          <option value="vencendo">Vencendo em 30 dias</option>
          <option value="vencido">Vencidos</option>
        </select>
      </div>
      <ExportRow
        total={total} count={lista.length}
        exportExcel={lista.length ? () => exportGroupedExcel({ groups, columns, valueColIndex: 6, sheetName: 'Seguros', fileName: 'seguros', reportTitle: 'Seguros' }) : null}
        exportPdf={lista.length ? () => exportToPDF({ title: 'Seguros', groups, columns, valueColIndex: 6, fileName: 'seguros' }) : null}
      />
      {lista.length === 0 ? <EmptyState icon={Shield} text="Nenhum seguro encontrado." /> : (
        <div className="grouped-report">
          {groups.map((g) => {
            const subtotal = g.rows.reduce((s, x) => s + (Number(x.valor_total) || 0), 0);
            return (
              <div key={g.label} className="group-block">
                <div className="group-head"><Building2 size={14} /> {g.label}</div>
                <table className="rtable"><thead><tr><th>Objeto</th><th>Seguradora</th><th>Vigência</th><th>Status</th><th className="right">Valor total</th></tr></thead>
                  <tbody>
                    {g.rows.map((s) => (
                      <tr key={s.id}>
                        <td>{s.objeto}</td>
                        <td className="dim">{s.seguradora}</td>
                        <td className="dim">{fmtDate(s.vigencia_inicio)} – {fmtDate(s.vigencia_fim)}</td>
                        <td><SeguroStatusBadge vigenciaFim={s.vigencia_fim} /></td>
                        <td className="right mono">{fmtBRL(s.valor_total)}</td>
                      </tr>
                    ))}
                    <tr className="subtotal-row"><td colSpan={4}>Subtotal {g.label.replace('Empresa: ', '')}</td><td className="right mono">{fmtBRL(subtotal)}</td></tr>
                  </tbody>
                </table>
              </div>
            );
          })}
          <div className="grand-total-row">
            <span>Total geral</span>
            <span className="mono">{fmtBRL(total)}</span>
          </div>
        </div>
      )}
      <style jsx>{inputCss}</style>
      <RptStyle /><ViewStyle />
    </div>
  );
}

function ReportVendas({ data, onBack }) {
  const [empresaFiltro, setEmpresaFiltro] = useState('todas');
  const [setorFiltro, setSetorFiltro] = useState('todos');
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';

  // Agrega os lançamentos mensais em totais por empresa + setor + ano
  const linhas = useMemo(() => {
    const map = new Map();
    (data.vendas || [])
      .filter((v) => empresaFiltro === 'todas' || v.empresa_id === empresaFiltro)
      .filter((v) => setorFiltro === 'todos' || v.setor === setorFiltro)
      .forEach((v) => {
        const key = `${v.empresa_id}|${v.setor}|${v.ano}`;
        if (!map.has(key)) map.set(key, { empresa_id: v.empresa_id, setor: v.setor, ano: v.ano, total: 0 });
        map.get(key).total += Number(v.valor) || 0;
      });
    return Array.from(map.values()).sort((a, b) => b.ano - a.ano || a.setor.localeCompare(b.setor));
  }, [data.vendas, empresaFiltro, setorFiltro]);

  const total = linhas.reduce((s, l) => s + l.total, 0);
  const columns = [
    { header: 'Setor', get: (l) => setorLabel(l.setor), width: 22 },
    { header: 'Ano', get: (l) => l.ano, width: 10 },
    { header: 'Total do ano', get: (l) => l.total, width: 16 },
  ];
  const groups = groupRows(linhas, (l) => l.empresa_id, (l) => `Empresa: ${empresaNome(l.empresa_id)}`);

  return (
    <div className="report-body">
      <ReportHeader title="Vendas por setor" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10 }}>
        <select className="select" value={empresaFiltro} onChange={(e) => setEmpresaFiltro(e.target.value)}>
          <option value="todas">Todas as empresas</option>
          {data.empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}
        </select>
        <select className="select" value={setorFiltro} onChange={(e) => setSetorFiltro(e.target.value)}>
          <option value="todos">Todos os setores</option>
          {SETORES_VENDA.map((s) => <option key={s.value} value={s.value}>{s.label}</option>)}
        </select>
      </div>
      <ExportRow
        total={total} count={linhas.length}
        exportExcel={linhas.length ? () => exportGroupedExcel({ groups, columns, valueColIndex: 2, sheetName: 'Vendas', fileName: 'vendas-por-setor', reportTitle: 'Vendas por Setor' }) : null}
        exportPdf={linhas.length ? () => exportToPDF({ title: 'Vendas por Setor', groups, columns, valueColIndex: 2, fileName: 'vendas-por-setor' }) : null}
      />
      {linhas.length === 0 ? <EmptyState icon={ShoppingCart} text="Nenhum lançamento de venda encontrado." /> : (
        <div className="grouped-report">
          {groups.map((g) => {
            const subtotal = g.rows.reduce((s, l) => s + l.total, 0);
            return (
              <div key={g.label} className="group-block">
                <div className="group-head"><Building2 size={14} /> {g.label}</div>
                <table className="rtable"><thead><tr><th>Setor</th><th>Ano</th><th className="right">Total</th></tr></thead>
                  <tbody>
                    {g.rows.map((l) => (
                      <tr key={`${l.setor}-${l.ano}`}>
                        <td>{setorLabel(l.setor)}</td>
                        <td className="dim">{l.ano}</td>
                        <td className="right mono">{fmtBRL(l.total)}</td>
                      </tr>
                    ))}
                    <tr className="subtotal-row"><td colSpan={2}>Subtotal {g.label.replace('Empresa: ', '')}</td><td className="right mono">{fmtBRL(subtotal)}</td></tr>
                  </tbody>
                </table>
              </div>
            );
          })}
          <div className="grand-total-row">
            <span>Total geral</span>
            <span className="mono">{fmtBRL(total)}</span>
          </div>
        </div>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function RptStyle() {
  return <style jsx global>{`
    .rtotal { display:flex; align-items:center; flex-wrap:wrap; gap:10px; margin:12px 20px 14px; font-size:13px; color:var(--text-muted); }
    .rtotal .mono { color:var(--text); font-size:15px; font-weight:600; }
    .rbuttons { display:flex; gap:8px; margin-left:auto; }
    .grouped-report { padding:0 20px 8px; display:flex; flex-direction:column; gap:18px; }
    .group-block { display:flex; flex-direction:column; gap:6px; }
    .group-head { display:flex; align-items:center; gap:6px; font-family:'Space Grotesk',sans-serif; font-size:13.5px; color:var(--accent); font-weight:600; }
    .group-block .rtable { width:100%; margin:0; }
    .subtotal-row td { font-weight:600; color:var(--text); border-top:1px solid var(--border-strong); border-bottom:none; }
    .grand-total-row { display:flex; justify-content:space-between; align-items:center; margin-top:4px; padding:12px 16px; background:var(--accent); border-radius:10px; font-size:14px; font-weight:700; color:var(--accent-contrast); }
    .grand-total-row .mono { font-size:16px; }
  `}</style>;
}

// ---------- Forms ----------

function EmpresaForm({ item, onClose, onSave }) {
  const [nome, setNome] = useState(item?.nome || '');
  return (
    <Modal title={item ? 'Editar empresa' : 'Nova empresa'} onClose={onClose}>
      <Field label="Nome da empresa"><input className="input" value={nome} onChange={(e) => setNome(e.target.value)} placeholder="Ex: Matriz Ltda" autoFocus /></Field>
      <PrimaryButton full onClick={() => nome.trim() && onSave({ id: item?.id, nome: nome.trim() })}>{item ? 'Salvar alterações' : 'Salvar empresa'}</PrimaryButton>
      <style jsx>{inputCss}</style>
    </Modal>
  );
}

function BancoForm({ item, empresas, onClose, onSave }) {
  const [empresaId, setEmpresaId] = useState(item?.empresaId || empresas[0]?.id || '');
  const [nomeBanco, setNomeBanco] = useState(item?.nomeBanco || '');
  const [saldo, setSaldo] = useState(item?.saldo ?? '');
  const canSave = empresaId && nomeBanco.trim();
  return (
    <Modal title={item ? 'Editar banco' : 'Novo banco'} onClose={onClose}>
      {empresas.length === 0 ? <p className="muted2">Cadastre uma empresa antes.</p> : (
        <>
          <Field label="Empresa"><select className="select" value={empresaId} onChange={(e) => setEmpresaId(e.target.value)}>{empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}</select></Field>
          <Field label="Banco"><input className="input" value={nomeBanco} onChange={(e) => setNomeBanco(e.target.value)} placeholder="Ex: Itaú, Bradesco, Nubank..." /></Field>
          <Field label="Saldo atual (R$)"><DecimalMaskInput value={saldo} onChange={setSaldo} /></Field>
          <PrimaryButton full onClick={() => canSave && onSave({ id: item?.id, empresaId, nomeBanco: nomeBanco.trim(), saldo: parseFloat(saldo) || 0 })}>{item ? 'Salvar alterações' : 'Salvar banco'}</PrimaryButton>
        </>
      )}
      <style jsx>{inputCss}</style>
    </Modal>
  );
}

function ContaForm({ item, empresas, onClose, onSave }) {
  const [empresaId, setEmpresaId] = useState(item?.empresaId || empresas[0]?.id || '');
  const [descricao, setDescricao] = useState(item?.descricao || '');
  const [valor, setValor] = useState(item?.valor ?? '');
  const [dataVencimento, setDataVencimento] = useState(item?.dataVencimento || todayISO());
  const canSave = empresaId && descricao.trim() && valor && dataVencimento;
  return (
    <Modal title={item ? 'Editar conta a pagar' : 'Nova conta a pagar'} onClose={onClose}>
      {empresas.length === 0 ? <p className="muted2">Cadastre uma empresa antes.</p> : (
        <>
          <Field label="Empresa"><select className="select" value={empresaId} onChange={(e) => setEmpresaId(e.target.value)}>{empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}</select></Field>
          <Field label="Descrição"><input className="input" value={descricao} onChange={(e) => setDescricao(e.target.value)} placeholder="Ex: Aluguel, fornecedor X..." /></Field>
          <Field label="Valor (R$)"><DecimalMaskInput value={valor} onChange={setValor} /></Field>
          <Field label="Data de vencimento"><input className="input" type="date" value={dataVencimento} onChange={(e) => setDataVencimento(e.target.value)} /></Field>
          <PrimaryButton full onClick={() => canSave && onSave({ id: item?.id, empresaId, descricao: descricao.trim(), valor: parseFloat(valor) || 0, dataVencimento, status: item?.status || 'pendente', dataPagamento: item?.dataPagamento || null })}>{item ? 'Salvar alterações' : 'Salvar conta'}</PrimaryButton>
        </>
      )}
      <style jsx>{inputCss}</style>
    </Modal>
  );
}

function EmprestimoForm({ item, empresas, onClose, onSave }) {
  const [tipo, setTipo] = useState(item?.tipo || 'banco');
  const [empresaId, setEmpresaId] = useState(item?.empresaId || empresas[0]?.id || '');
  const [credor, setCredor] = useState(item?.credor || '');
  const [linhaCredito, setLinhaCredito] = useState(item?.linhaCredito || LINHAS_CREDITO[0]);
  const [empresaCredoraId, setEmpresaCredoraId] = useState(item?.empresaCredoraId || '');
  const [valorTotal, setValorTotal] = useState(item?.valorTotal ?? '');
  const [valorParcela, setValorParcela] = useState(item?.valorParcela ?? '');
  const [numParcelas, setNumParcelas] = useState(item?.numParcelas ?? '');
  const [dataInicio, setDataInicio] = useState(item?.dataInicio || todayISO());
  const [temCarencia, setTemCarencia] = useState(item?.temCarencia || false);
  const [prazoCarencia, setPrazoCarencia] = useState(item?.prazoCarencia ?? '');

  const empresasCredoras = empresas.filter((emp) => emp.id !== empresaId);
  const canSave = empresaId && valorTotal && (tipo === 'banco' ? credor.trim() : empresaCredoraId);

  const handleSave = () => onSave({
    id: item?.id, empresaId, tipo,
    credor: tipo === 'banco' ? credor.trim() : '',
    linhaCredito: tipo === 'banco' ? linhaCredito : null,
    empresaCredoraId: tipo === 'empresa' ? empresaCredoraId : null,
    valorTotal: parseFloat(valorTotal) || 0,
    valorParcela: parseFloat(valorParcela) || 0,
    numParcelas: parseInt(numParcelas) || 0,
    dataInicio,
    temCarencia,
    prazoCarencia: temCarencia ? (parseInt(prazoCarencia, 10) || 0) : 0,
  });

  return (
    <Modal title={item ? 'Editar empréstimo' : 'Novo empréstimo'} onClose={onClose}>
      {empresas.length === 0 ? <p className="muted2">Cadastre uma empresa antes.</p> : (
        <>
          <Field label="Tipo de empréstimo">
            <div className="tipo-toggle">
              <button type="button" className={tipo === 'banco' ? 'active' : ''} onClick={() => setTipo('banco')}>Bancário</button>
              <button type="button" className={tipo === 'empresa' ? 'active' : ''} onClick={() => setTipo('empresa')}>Entre empresas do grupo</button>
            </div>
          </Field>
          <Field label="Empresa devedora"><select className="select" value={empresaId} onChange={(e) => setEmpresaId(e.target.value)}>{empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}</select></Field>

          {tipo === 'banco' ? (
            <>
              <Field label="Credor / instituição"><input className="input" value={credor} onChange={(e) => setCredor(e.target.value)} placeholder="Ex: Banco do Brasil" /></Field>
              <Field label="Linha de crédito"><select className="select" value={linhaCredito} onChange={(e) => setLinhaCredito(e.target.value)}>{LINHAS_CREDITO.map((l) => <option key={l} value={l}>{l}</option>)}</select></Field>
            </>
          ) : (
            <Field label="Empresa credora">
              {empresasCredoras.length === 0 ? <p className="muted2">Cadastre outra empresa para poder registrar um empréstimo entre empresas do grupo.</p> : (
                <select className="select" value={empresaCredoraId} onChange={(e) => setEmpresaCredoraId(e.target.value)}>
                  <option value="">Selecione...</option>
                  {empresasCredoras.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}
                </select>
              )}
            </Field>
          )}

          <Field label="Valor total (R$)"><DecimalMaskInput value={valorTotal} onChange={setValorTotal} /></Field>
          <Field label="Valor da parcela (R$)"><DecimalMaskInput value={valorParcela} onChange={setValorParcela} /></Field>
          <Field label="Prazo total do contrato (meses)"><input className="input" type="number" value={numParcelas} onChange={(e) => setNumParcelas(e.target.value)} placeholder="Ex: 36" /></Field>
          <Field label="Data de início">
            <input className="input" type="date" value={dataInicio} onChange={(e) => setDataInicio(e.target.value)} />
          </Field>

          <Field label="Carência">
            <label className="carencia-check">
              <input type="checkbox" checked={temCarencia} onChange={(e) => setTemCarencia(e.target.checked)} />
              <span>Este empréstimo tem período de carência</span>
            </label>
          </Field>
          {temCarencia && (
            <Field label="Prazo de carência (meses)">
              <input className="input" type="number" min="1" value={prazoCarencia} onChange={(e) => setPrazoCarencia(e.target.value)} placeholder="Ex: 6" />
            </Field>
          )}
          {numParcelas && dataInicio && (() => {
            const totalMeses = parseInt(numParcelas, 10) || 0;
            const carenciaMeses = temCarencia ? (parseInt(prazoCarencia, 10) || 0) : 0;
            const qtdAmortizacao = Math.max(totalMeses - carenciaMeses, 0);
            if (!qtdAmortizacao) return carenciaMeses >= totalMeses && totalMeses > 0 ? (
              <p className="carencia-hint" style={{ color: 'var(--danger)' }}>A carência não pode ser igual ou maior que o prazo total do contrato.</p>
            ) : null;
            const primeiraParcela = addMonthsLocal(dataInicio, carenciaMeses);
            const ultimaParcela = addMonthsLocal(dataInicio, totalMeses - 1);
            return (
              <p className="carencia-hint">
                {carenciaMeses > 0 ? `Carência de ${carenciaMeses} ${carenciaMeses === 1 ? 'mês' : 'meses'}, depois ` : ''}
                {qtdAmortizacao} {qtdAmortizacao === 1 ? 'parcela' : 'parcelas'} de amortização: 1ª em {fmtDate(primeiraParcela)}, última em {fmtDate(ultimaParcela)}.
              </p>
            );
          })()}

          <PrimaryButton full onClick={() => canSave && handleSave()}>{item ? 'Salvar alterações' : 'Salvar empréstimo'}</PrimaryButton>
        </>
      )}
      <style jsx>{inputCss}</style>
      <style jsx>{`
        .tipo-toggle { display:flex; gap:6px; }
        .tipo-toggle button { flex:1; background:var(--surface); border:1px solid var(--border); color:var(--text-muted); font-size:12.5px; font-weight:600; padding:9px 8px; border-radius:8px; cursor:pointer; }
        .tipo-toggle button.active { background:var(--accent); border-color:var(--accent); color:var(--accent-contrast); }
        .carencia-check { display:flex; align-items:center; gap:8px; font-size:13.5px; color:var(--text); cursor:pointer; }
        .carencia-check input { width:16px; height:16px; accent-color:var(--accent); }
        .carencia-hint { font-size:12.5px; color:var(--text-muted); margin:-6px 0 0; }
      `}</style>
    </Modal>
  );
}

function TransferForm({ bancos, onClose, onSave }) {
  const [bancoOrigemId, setBancoOrigemId] = useState(bancos[0]?.id || '');
  const [bancoDestinoId, setBancoDestinoId] = useState(bancos[1]?.id || bancos[0]?.id || '');
  const [valor, setValor] = useState('');
  const [data, setDataCampo] = useState(todayISO());
  const [descricao, setDescricao] = useState('');
  const [erro, setErro] = useState('');
  const handleSave = () => {
    if (bancoOrigemId === bancoDestinoId) { setErro('Escolha bancos de origem e destino diferentes.'); return; }
    const v = parseFloat(valor);
    if (!v || v <= 0) { setErro('Informe um valor válido.'); return; }
    setErro('');
    onSave({ bancoOrigemId, bancoDestinoId, valor: v, data, descricao: descricao.trim() });
  };
  return (
    <Modal title="Nova transferência" onClose={onClose}>
      <Field label="Banco de origem"><select className="select" value={bancoOrigemId} onChange={(e) => setBancoOrigemId(e.target.value)}>{bancos.map((b) => <option key={b.id} value={b.id}>{b.nome_banco} · {fmtBRL(b.saldo)}</option>)}</select></Field>
      <Field label="Banco de destino"><select className="select" value={bancoDestinoId} onChange={(e) => setBancoDestinoId(e.target.value)}>{bancos.map((b) => <option key={b.id} value={b.id}>{b.nome_banco} · {fmtBRL(b.saldo)}</option>)}</select></Field>
      <Field label="Valor (R$)"><DecimalMaskInput value={valor} onChange={setValor} /></Field>
      <Field label="Data"><input className="input" type="date" value={data} onChange={(e) => setDataCampo(e.target.value)} /></Field>
      <Field label="Descrição (opcional)"><input className="input" value={descricao} onChange={(e) => setDescricao(e.target.value)} placeholder="Ex: Reforço de caixa" /></Field>
      {erro && <p style={{ color: 'var(--danger)', fontSize: 13, margin: 0 }}>{erro}</p>}
      <PrimaryButton full onClick={handleSave}>Transferir</PrimaryButton>
      <style jsx>{inputCss}</style>
    </Modal>
  );
}

function CartaoTaxaForm({ item, onClose, onSave }) {
  const [administradora, setAdministradora] = useState(item?.administradora || '');
  const [bandeira, setBandeira] = useState(item?.bandeira || BANDEIRAS[0]);
  const [taxaPix, setTaxaPix] = useState(item?.taxaPix ?? '');
  const [taxaDebito, setTaxaDebito] = useState(item?.taxaDebito ?? '');
  const [taxaCreditoAvista, setTaxaCreditoAvista] = useState(item?.taxaCreditoAvista ?? '');
  const [parcelas, setParcelas] = useState(() => { const base = {}; PARCELAS_RANGE.forEach((n) => { base[n] = item?.parcelas?.[n] ?? ''; }); return base; });
  const canSave = administradora.trim() && bandeira;
  const handleSave = () => {
    const parcelasNum = {};
    PARCELAS_RANGE.forEach((n) => { parcelasNum[n] = parcelas[n] === '' ? null : parseFloat(parcelas[n]); });
    onSave({ id: item?.id, administradora: administradora.trim(), bandeira, taxaPix: taxaPix === '' ? null : parseFloat(taxaPix), taxaDebito: taxaDebito === '' ? null : parseFloat(taxaDebito), taxaCreditoAvista: taxaCreditoAvista === '' ? null : parseFloat(taxaCreditoAvista), parcelas: parcelasNum });
  };
  return (
    <Modal title={item ? 'Editar taxa de cartão' : 'Nova taxa de cartão'} onClose={onClose}>
      <Field label="Administradora / operadora"><input className="input" value={administradora} onChange={(e) => setAdministradora(e.target.value)} placeholder="Ex: Stone, Cielo, Rede..." autoFocus /></Field>
      <Field label="Bandeira"><select className="select" value={bandeira} onChange={(e) => setBandeira(e.target.value)}>{BANDEIRAS.map((b) => <option key={b} value={b}>{b}</option>)}</select></Field>
      <Field label="Taxa Pix (%)"><DecimalMaskInput value={taxaPix} onChange={setTaxaPix} /></Field>
      <Field label="Taxa Débito (%)"><DecimalMaskInput value={taxaDebito} onChange={setTaxaDebito} /></Field>
      <Field label="Taxa Crédito à vista (%)"><DecimalMaskInput value={taxaCreditoAvista} onChange={setTaxaCreditoAvista} /></Field>
      <Field label="Taxas de crédito parcelado (%)">
        <div className="pgrid">
          {PARCELAS_RANGE.map((n) => (
            <div key={n} className="pitem">
              <span>{n}x</span>
              <DecimalMaskInput value={parcelas[n]} onChange={(v) => setParcelas((p) => ({ ...p, [n]: v }))} />
            </div>
          ))}
        </div>
      </Field>
      <PrimaryButton full onClick={() => canSave && handleSave()}>{item ? 'Salvar alterações' : 'Salvar taxa'}</PrimaryButton>
      <style jsx>{inputCss}</style>
      <style jsx>{`
        .pgrid { display:grid; grid-template-columns:repeat(3,1fr); gap:8px; }
        .pitem { display:flex; flex-direction:column; gap:4px; }
        .pitem span { font-size:11px; color:var(--text-dim); }
        .pitem :global(.input) { padding:8px 9px; font-size:13px; }
      `}</style>
    </Modal>
  );
}

function BoletoTaxaForm({ item, onClose, onSave }) {
  const [administradora, setAdministradora] = useState(item?.administradora || '');
  const [taxaEmissao, setTaxaEmissao] = useState(item?.taxaEmissao ?? '');
  const [taxaBaixa, setTaxaBaixa] = useState(item?.taxaBaixa ?? '');
  const [taxaProtesto, setTaxaProtesto] = useState(item?.taxaProtesto ?? '');
  const [taxaAntecipacao, setTaxaAntecipacao] = useState(item?.taxaAntecipacao ?? '');
  const canSave = administradora.trim();
  return (
    <Modal title={item ? 'Editar taxa de boleto' : 'Nova taxa de boleto'} onClose={onClose}>
      <Field label="Administradora / banco"><input className="input" value={administradora} onChange={(e) => setAdministradora(e.target.value)} placeholder="Ex: Itaú, Bradesco, Asaas..." autoFocus /></Field>
      <Field label="Taxa de emissão (R$)"><DecimalMaskInput value={taxaEmissao} onChange={setTaxaEmissao} /></Field>
      <Field label="Taxa de baixa (R$)"><DecimalMaskInput value={taxaBaixa} onChange={setTaxaBaixa} /></Field>
      <Field label="Taxa de protesto (R$)"><DecimalMaskInput value={taxaProtesto} onChange={setTaxaProtesto} /></Field>
      <Field label="Taxa de antecipação (% ao mês)"><DecimalMaskInput value={taxaAntecipacao} onChange={setTaxaAntecipacao} /></Field>
      <PrimaryButton full onClick={() => canSave && onSave({ id: item?.id, administradora: administradora.trim(), taxaEmissao: taxaEmissao === '' ? null : parseFloat(taxaEmissao), taxaBaixa: taxaBaixa === '' ? null : parseFloat(taxaBaixa), taxaProtesto: taxaProtesto === '' ? null : parseFloat(taxaProtesto), taxaAntecipacao: taxaAntecipacao === '' ? null : parseFloat(taxaAntecipacao) })}>{item ? 'Salvar alterações' : 'Salvar taxa'}</PrimaryButton>
      <style jsx>{inputCss}</style>
    </Modal>
  );
}
