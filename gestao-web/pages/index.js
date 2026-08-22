import React, { useEffect, useState, useMemo } from 'react';
import {
  Building2, Landmark, CalendarClock, HandCoins, Plus, X, Trash2, Pencil,
  LayoutDashboard, TrendingDown, AlertCircle, CheckCircle2, FileBarChart,
  ChevronRight, ArrowLeft, Download, ArrowLeftRight, ChevronDown, ChevronUp,
  CreditCard, Receipt, Percent, RefreshCw,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, Cell } from 'recharts';
import * as db from '../lib/data';

const BANDEIRAS = ['Visa', 'Mastercard', 'Elo', 'American Express', 'Hipercard', 'Outra'];
const PARCELAS_RANGE = Array.from({ length: 11 }, (_, i) => i + 2);

const fmtBRL = (v) => (Number(v) || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const fmtPct = (v) => (v === null || v === undefined || v === '') ? '—' : `${Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
const fmtDate = (iso) => { if (!iso) return '—'; const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };
const todayISO = () => new Date().toISOString().slice(0, 10);
const daysUntil = (iso) => Math.round((new Date(iso) - new Date(todayISO())) / 86400000);

// ---------- Reusable UI ----------

function Field({ label, children }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
      <style jsx>{`
        .field { display:flex; flex-direction:column; gap:6px; }
        span { font-size:12px; letter-spacing:.04em; text-transform:uppercase; color:#8891A0; }
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
        .modal { background:#1C222C; border:1px solid #2A3140; border-radius:16px; width:100%; max-width:460px; max-height:88vh; overflow-y:auto; padding:20px; }
        .head { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; }
        .head h3 { font-family:'Space Grotesk',sans-serif; font-size:18px; color:#E8EAED; margin:0; font-weight:600; }
        .body { display:flex; flex-direction:column; gap:14px; }
        .icon-btn { background:none; border:none; color:#8891A0; cursor:pointer; padding:4px; }
      `}</style>
    </div>
  );
}

const inputCss = `
  .input, .select { background:#12161D; border:1px solid #2A3140; border-radius:10px; padding:11px 12px; color:#E8EAED; font-size:15px; outline:none; width:100%; box-sizing:border-box; }
  .input:focus, .select:focus { border-color:#4FD1AE; }
`;

function PrimaryButton({ children, onClick, full, tone }) {
  return (
    <button onClick={onClick} className={`btn ${full ? 'full' : ''}`}>
      {children}
      <style jsx>{`
        .btn { background:${tone === 'danger' ? '#E2596B' : '#4FD1AE'}; color:#0E1116; border:none; border-radius:10px; padding:12px 18px; font-weight:600; font-size:14px; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; gap:6px; }
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
        .btn { background:transparent; color:#8891A0; border:1px solid #2A3140; border-radius:10px; padding:11px 18px; font-weight:600; font-size:14px; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; gap:6px; }
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
        .empty { display:flex; flex-direction:column; align-items:center; gap:10px; padding:48px 20px; color:#5A6272; text-align:center; }
        p { margin:0; font-size:14px; }
      `}</style>
    </div>
  );
}

function SectionTitle({ children }) {
  return (
    <h2 className="st">
      {children}
      <style jsx>{`.st { font-family:'Space Grotesk',sans-serif; font-size:14px; letter-spacing:.02em; color:#8891A0; margin:20px 0 10px; font-weight:600; text-transform:uppercase; }`}</style>
    </h2>
  );
}

function ExportButton({ onClick, disabled, label = 'Exportar Excel' }) {
  return (
    <button className="exp" onClick={onClick} disabled={disabled}>
      <Download size={14} /> {label}
      <style jsx>{`
        .exp { display:inline-flex; align-items:center; gap:6px; background:none; border:1px solid #2A3140; color:#4FD1AE; font-size:12.5px; font-weight:600; padding:8px 12px; border-radius:9px; cursor:pointer; }
        .exp:disabled { color:#5A6272; }
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

// ---------- Shared card styles ----------
function ViewStyle() {
  return (
    <style jsx global>{`
      .view { padding:16px 20px 24px; display:flex; flex-direction:column; gap:12px; }
      .card { background:#1C222C; border:1px solid #2A3140; border-radius:14px; padding:14px 16px; }
      .card-head { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; gap:8px; }
      .card-title { display:flex; align-items:center; gap:8px; font-family:'Space Grotesk',sans-serif; font-size:15px; color:#E8EAED; font-weight:600; }
      .card-actions { display:flex; gap:2px; flex-shrink:0; }
      .card-grid { display:grid; grid-template-columns:1fr 1fr; gap:10px 14px; }
      .k { display:block; font-size:11px; color:#5A6272; text-transform:uppercase; letter-spacing:.03em; margin-bottom:2px; }
      .v { display:block; font-size:14px; color:#E8EAED; }
      .mono { font-family:'IBM Plex Mono',monospace; }
      .icon-btn { background:none; border:none; color:#5A6272; cursor:pointer; padding:5px; border-radius:6px; display:flex; }
      .icon-btn:hover { color:#4FD1AE; }
      .icon-btn.danger:hover { color:#E2596B; }
      .row { display:flex; justify-content:space-between; align-items:center; padding:12px 14px; background:#1C222C; border:1px solid #2A3140; border-radius:12px; margin-bottom:8px; font-size:14px; color:#E8EAED; }
      .row-sub { font-size:12px; margin-top:2px; }
      .badge { display:inline-block; background:#12161D; border:1px solid #2A3140; color:#8891A0; font-size:10.5px; font-weight:600; padding:2px 7px; border-radius:6px; margin-left:6px; }
      .add-inline-btn { display:inline-flex; align-items:center; gap:5px; background:#4FD1AE; color:#0E1116; border:none; border-radius:9px; padding:9px 12px; font-size:12.5px; font-weight:700; cursor:pointer; white-space:nowrap; }
      .chart-box { background:#1C222C; border:1px solid #2A3140; border-radius:14px; padding:12px 8px 4px; margin-bottom:4px; }
      .report-filters { display:flex; gap:8px; padding:4px 20px 0; flex-wrap:wrap; }
      .rtable { width:calc(100% - 40px); margin:0 20px; border-collapse:collapse; font-size:13px; }
      .rtable thead th { text-align:left; font-size:10.5px; text-transform:uppercase; color:#5A6272; padding:0 8px 8px 0; border-bottom:1px solid #2A3140; }
      .rtable tbody td { padding:10px 8px 10px 0; border-bottom:1px solid #1E2430; color:#E8EAED; }
      .rtable .dim { color:#8891A0; }
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
            onToggle={(c) => withSave(() => db.toggleContaStatus(c))}
            onEdit={(it) => setModal({ type: 'conta', item: it })}
            onDelete={(id, label) => setConfirmDelete({ type: 'conta', id, label })} />
        )}
        {tab === 'emprestimos' && (
          <EmprestimosView data={data}
            onEdit={(it) => setModal({ type: 'emprestimo', item: it })}
            onDelete={(id, label) => setConfirmDelete({ type: 'emprestimo', id, label })}
            onToggleParcela={(p) => withSave(() => db.toggleParcela(p))} />
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
        {tab === 'relatorios' && <ReportsHome view={reportView} setView={setReportView} data={data} />}
      </main>

      {['empresas', 'bancos', 'contas', 'emprestimos'].includes(tab) && (
        <button className="fab" onClick={() => setModal({ type: tab === 'empresas' ? 'empresa' : tab === 'bancos' ? 'banco' : tab === 'contas' ? 'conta' : 'emprestimo', item: null })}>
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

      {confirmDelete && (
        <Modal title="Excluir registro" onClose={() => setConfirmDelete(null)}>
          <p style={{ color: '#B7BEC9', fontSize: 14, margin: 0 }}>Tem certeza que deseja excluir <strong style={{ color: '#E8EAED' }}>{confirmDelete.label}</strong>? Essa ação não pode ser desfeita.</p>
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
        .app { min-height:100vh; background:#0E1116; display:flex; flex-direction:column; max-width:520px; margin:0 auto; position:relative; }
        .header { padding:20px 20px 18px; border-bottom:1px solid #1E2430; }
        .eyebrow { font-size:11px; letter-spacing:.08em; text-transform:uppercase; color:#5A6272; margin-bottom:4px; }
        .saldo-total { font-family:'IBM Plex Mono',monospace; font-size:32px; font-weight:600; color:#E8EAED; }
        .header-meta { margin-top:8px; display:flex; align-items:center; gap:8px; font-size:12px; color:#5A6272; }
        .refresh-btn { margin-left:auto; background:none; border:1px solid #2A3140; color:#8891A0; border-radius:6px; padding:4px 6px; cursor:pointer; display:flex; }
        .content { flex:1; overflow-y:auto; padding-bottom:70px; }
        .fab { position:fixed; right:calc(50% - 260px + 20px); bottom:80px; width:52px; height:52px; border-radius:50%; background:#4FD1AE; color:#0E1116; border:none; display:flex; align-items:center; justify-content:center; box-shadow:0 6px 18px rgba(79,209,174,0.35); cursor:pointer; z-index:10; }
        @media (max-width: 560px) { .fab { right:20px; } }
        .tabbar { position:fixed; bottom:0; left:0; right:0; max-width:520px; margin:0 auto; display:flex; background:#12161D; border-top:1px solid #1E2430; padding:6px 2px calc(6px + env(safe-area-inset-bottom)); z-index:20; overflow-x:auto; }
        .tab { flex:1; min-width:58px; display:flex; flex-direction:column; align-items:center; gap:3px; background:none; border:none; color:#5A6272; padding:6px 1px; cursor:pointer; }
        .tab span { font-size:9.5px; }
        .tab.active { color:#4FD1AE; }
        .saving-overlay { position:fixed; top:12px; left:50%; transform:translateX(-50%); background:#1C222C; border:1px solid #2A3140; color:#4FD1AE; font-size:12.5px; padding:8px 14px; border-radius:20px; z-index:60; }
      `}</style>
    </div>
  );
}

function CenterMsg({ text, error }) {
  return (
    <div style={{ minHeight: '100vh', background: '#0E1116', display: 'flex', alignItems: 'center', justifyContent: 'center', color: error ? '#E2596B' : '#5A6272', fontFamily: 'Inter,sans-serif', padding: 24, textAlign: 'center' }}>
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

  if (!data.empresas.length && !data.bancos.length) {
    return <EmptyState icon={Building2} text="Cadastre sua primeira empresa para começar." />;
  }

  return (
    <div className="dash">
      <div className="cards-row">
        <div className="stat-card"><TrendingDown size={16} color="#E8A33D" /><div className="stat-label">A pagar (pendente)</div><div className="stat-value">{fmtBRL(totalPendente)}</div></div>
        <div className="stat-card"><HandCoins size={16} color="#E2596B" /><div className="stat-label">Empréstimos ativos</div><div className="stat-value">{fmtBRL(totalEmprestimos)}</div></div>
      </div>

      {saldoPorEmpresa.length > 0 && (
        <>
          <SectionTitle>Saldos por empresa</SectionTitle>
          <div className="chart-box">
            <ResponsiveContainer width="100%" height={Math.max(120, saldoPorEmpresa.length * 42)}>
              <BarChart data={saldoPorEmpresa} layout="vertical" margin={{ top: 4, right: 12, bottom: 4, left: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#1E2430" horizontal={false} />
                <XAxis type="number" tick={{ fill: '#5A6272', fontSize: 10 }} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} axisLine={false} tickLine={false} />
                <YAxis type="category" dataKey="name" tick={{ fill: '#8891A0', fontSize: 11 }} axisLine={false} tickLine={false} width={70} />
                <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ background: '#1C222C', border: '1px solid #2A3140', borderRadius: 8, fontSize: 12 }} labelStyle={{ color: '#E8EAED' }} />
                <Bar dataKey="valor" fill="#4FD1AE" radius={[0, 4, 4, 0]} barSize={16} />
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
                  <div className="row-sub" style={{ color: dias <= 3 ? '#E2596B' : '#5A6272' }}>
                    {fmtDate(c.data_vencimento)} {dias === 0 ? '· vence hoje' : dias < 0 ? `· atrasada ${Math.abs(dias)}d` : `· em ${dias}d`}
                  </div>
                </div>
                <span className="mono">{fmtBRL(c.valor)}</span>
              </div>
            );
          })}
        </div>
      )}

      <button className="reports-link" onClick={onGoReports}><FileBarChart size={16} /> Ver relatórios completos <ChevronRight size={15} /></button>

      <style jsx>{`
        .dash { padding:16px 20px 24px; display:flex; flex-direction:column; gap:4px; }
        .cards-row { display:flex; gap:10px; margin-bottom:8px; }
        .stat-card { flex:1; background:#1C222C; border:1px solid #2A3140; border-radius:14px; padding:14px; display:flex; flex-direction:column; gap:6px; }
        .stat-label { font-size:12px; color:#8891A0; }
        .stat-value { font-family:'IBM Plex Mono',monospace; font-size:18px; color:#E8EAED; font-weight:600; }
        .muted { color:#5A6272; font-size:14px; padding:8px 0 16px; }
        .reports-link { margin-top:18px; display:flex; align-items:center; justify-content:center; gap:8px; width:100%; background:#1C222C; border:1px solid #2A3140; color:#4FD1AE; font-weight:600; font-size:13px; padding:13px; border-radius:12px; cursor:pointer; }
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
        .transfer-btn { display:flex; align-items:center; justify-content:center; gap:8px; width:100%; background:rgba(79,209,174,0.1); border:1px solid #2A3140; color:#4FD1AE; font-weight:600; font-size:13.5px; padding:12px; border-radius:12px; cursor:pointer; margin-bottom:4px; }
        .transfer-row { display:flex; justify-content:space-between; align-items:center; gap:10px; }
        .transfer-path { display:flex; align-items:center; gap:6px; font-size:13.5px; color:#E8EAED; font-weight:600; }
        .transfer-actions { display:flex; align-items:center; gap:8px; }
      `}</style>
      <ViewStyle />
    </div>
  );
}

// ---------- Contas a pagar ----------

function ContasView({ data, onToggle, onEdit, onDelete }) {
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
              <div><span className="k">Vencimento</span><span className="v" style={{ color: atrasada ? '#E2596B' : undefined }}>{fmtDate(c.data_vencimento)}</span></div>
              <div><span className="k">Valor</span><span className="v mono">{fmtBRL(c.valor)}</span></div>
            </div>
            <button className={`status-btn ${c.status}`} onClick={() => onToggle(c)}>
              {c.status === 'pago' ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />}
              {c.status === 'pago' ? `Pago em ${fmtDate(c.data_pagamento)}` : atrasada ? 'Pendente · atrasada' : 'Pendente'}
            </button>
          </div>
        );
      })}
      <style jsx>{`
        .status-btn { margin-top:10px; display:inline-flex; align-items:center; gap:6px; border:none; border-radius:8px; padding:7px 12px; font-size:12px; font-weight:600; cursor:pointer; }
        .status-btn.pendente { background:rgba(232,163,61,0.15); color:#E8A33D; }
        .status-btn.pago { background:rgba(79,209,174,0.15); color:#4FD1AE; }
      `}</style>
      <ViewStyle />
    </div>
  );
}

// ---------- Empréstimos ----------

function EmprestimosView({ data, onEdit, onDelete, onToggleParcela }) {
  const [expandedId, setExpandedId] = useState(null);
  if (!data.emprestimos.length) return <EmptyState icon={HandCoins} text="Nenhum empréstimo cadastrado. Toque em + para adicionar." />;
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  return (
    <div className="view">
      {data.emprestimos.map((e) => {
        const parcelas = e.parcelas || [];
        const pagas = parcelas.filter((p) => p.status === 'pago').length;
        const expanded = expandedId === e.id;
        return (
          <div key={e.id} className="card">
            <div className="card-head">
              <div className="card-title"><HandCoins size={16} /> {e.credor}</div>
              <div className="card-actions">
                <button className="icon-btn" onClick={() => onEdit({ id: e.id, empresaId: e.empresa_id, credor: e.credor, valorTotal: e.valor_total, valorParcela: e.valor_parcela, numParcelas: e.num_parcelas, dataInicio: e.data_inicio, parcelas: e.parcelas })}><Pencil size={16} /></button>
                <button className="icon-btn danger" onClick={() => onDelete(e.id, e.credor)}><Trash2 size={16} /></button>
              </div>
            </div>
            <div className="card-grid">
              <div><span className="k">Empresa</span><span className="v">{empresaNome(e.empresa_id)}</span></div>
              <div><span className="k">Valor total</span><span className="v mono">{fmtBRL(e.valor_total)}</span></div>
              <div><span className="k">Parcela</span><span className="v mono">{fmtBRL(e.valor_parcela)}</span></div>
              <div><span className="k">Parcelas pagas</span><span className="v">{pagas} / {parcelas.length}</span></div>
              <div><span className="k">Início</span><span className="v">{fmtDate(e.data_inicio)}</span></div>
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
                      <span className="parcela-date" style={{ color: atrasada ? '#E2596B' : undefined }}>{p.status === 'pago' ? `Pago ${fmtDate(p.data_pagamento)}` : fmtDate(p.data_vencimento)}</span>
                      <span className="parcela-valor mono">{fmtBRL(p.valor)}</span>
                      {p.status === 'pago' ? <CheckCircle2 size={15} color="#4FD1AE" /> : <AlertCircle size={15} color={atrasada ? '#E2596B' : '#E8A33D'} />}
                    </button>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
      <style jsx>{`
        .expand-btn { margin-top:10px; display:inline-flex; align-items:center; gap:6px; background:none; border:1px solid #2A3140; border-radius:8px; padding:7px 12px; font-size:12px; font-weight:600; color:#8891A0; cursor:pointer; }
        .parcelas-list { margin-top:10px; display:flex; flex-direction:column; gap:6px; }
        .parcela-row { display:flex; align-items:center; gap:10px; background:#12161D; border:1px solid #232A38; border-radius:9px; padding:9px 10px; font-size:12.5px; color:#E8EAED; cursor:pointer; width:100%; text-align:left; }
        .parcela-num { color:#5A6272; width:26px; flex-shrink:0; }
        .parcela-date { flex:1; }
        .parcela-valor { color:#B7BEC9; }
        .parcela-row.pago { opacity:0.65; }
      `}</style>
      <ViewStyle />
    </div>
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
          <ChevronRight size={18} color="#5A6272" />
        </button>
      ))}
      <style jsx>{`
        .reports-home { padding:16px 20px 24px; display:flex; flex-direction:column; gap:10px; }
        .report-option { display:flex; align-items:center; gap:12px; background:#1C222C; border:1px solid #2A3140; border-radius:14px; padding:14px; cursor:pointer; text-align:left; }
        .report-icon { width:38px; height:38px; border-radius:10px; background:rgba(79,209,174,0.12); color:#4FD1AE; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
        .report-text { flex:1; }
        .report-label { font-family:'Space Grotesk',sans-serif; font-size:14.5px; color:#E8EAED; font-weight:600; }
        .report-desc { font-size:12px; color:#5A6272; margin-top:2px; }
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
        .back { background:none; border:none; color:#8891A0; cursor:pointer; padding:4px; display:flex; }
        h2 { font-family:'Space Grotesk',sans-serif; font-size:17px; color:#E8EAED; font-weight:600; margin:0; }
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
                  <CartesianGrid strokeDasharray="3 3" stroke="#1E2430" horizontal={false} />
                  <XAxis type="number" tick={{ fill: '#5A6272', fontSize: 10 }} tickFormatter={(v) => `${v}%`} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" tick={{ fill: '#8891A0', fontSize: 11 }} axisLine={false} tickLine={false} width={90} />
                  <Tooltip formatter={(v) => fmtPct(v)} contentStyle={{ background: '#1C222C', border: '1px solid #2A3140', borderRadius: 8, fontSize: 12 }} labelStyle={{ color: '#E8EAED' }} />
                  <Bar dataKey="valor" fill="#4FD1AE" radius={[0, 4, 4, 0]} barSize={16} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
          <div style={{ padding: '14px 20px 4px' }}><ExportButton onClick={doExport} disabled={lista.length === 0} label="Exportar comparativo (Excel)" /></div>
        </>
      )}
      <style jsx>{`
        .expand-btn2 { margin-top:10px; display:inline-flex; align-items:center; gap:6px; background:none; border:1px solid #2A3140; border-radius:8px; padding:7px 12px; font-size:12px; font-weight:600; color:#8891A0; cursor:pointer; }
        .parcelas-taxa-grid { margin-top:10px; display:grid; grid-template-columns:repeat(4,1fr); gap:8px; }
        .pti { background:#12161D; border:1px solid #232A38; border-radius:9px; padding:8px; display:flex; flex-direction:column; gap:2px; }
        .pti .k { font-size:10px; } .pti .v { font-size:12.5px; }
        .muted2 { color:#5A6272; font-size:14px; padding:0 20px 8px; }
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
                  <CartesianGrid strokeDasharray="3 3" stroke="#1E2430" horizontal={false} />
                  <XAxis type="number" tick={{ fill: '#5A6272', fontSize: 10 }} tickFormatter={(v) => isAntecipacao ? `${v}%` : `R$${v}`} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="name" tick={{ fill: '#8891A0', fontSize: 11 }} axisLine={false} tickLine={false} width={90} />
                  <Tooltip formatter={(v) => isAntecipacao ? fmtPct(v) : fmtBRL(v)} contentStyle={{ background: '#1C222C', border: '1px solid #2A3140', borderRadius: 8, fontSize: 12 }} labelStyle={{ color: '#E8EAED' }} />
                  <Bar dataKey="valor" fill="#4FD1AE" radius={[0, 4, 4, 0]} barSize={16} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
          <div style={{ padding: '14px 20px 4px' }}><ExportButton onClick={doExport} disabled={data.taxasBoleto.length === 0} label="Exportar comparativo (Excel)" /></div>
        </>
      )}
      <style jsx>{`.muted2 { color:#5A6272; font-size:14px; padding:0 20px 8px; }`}</style>
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
  const options = [
    { key: 'bancos', label: 'Saldos por banco', desc: 'Total consolidado em cada conta bancária', icon: Landmark },
    { key: 'empresas', label: 'Saldos por empresa', desc: 'Total consolidado por empresa do grupo', icon: Building2 },
    { key: 'contas', label: 'Contas a pagar', desc: 'Pendentes e atrasadas, filtráveis por empresa', icon: CalendarClock },
    { key: 'pagamentos', label: 'Pagamentos realizados', desc: 'Histórico de contas já pagas', icon: CheckCircle2 },
  ];
  return (
    <div className="reports-home">
      {options.map(({ key, label, desc, icon: Icon }) => (
        <button key={key} className="report-option" onClick={() => setView(key)}>
          <div className="report-icon"><Icon size={18} /></div>
          <div className="report-text"><div className="report-label">{label}</div><div className="report-desc">{desc}</div></div>
          <ChevronRight size={18} color="#5A6272" />
        </button>
      ))}
      <style jsx>{`
        .reports-home { padding:16px 20px 24px; display:flex; flex-direction:column; gap:10px; }
        .report-option { display:flex; align-items:center; gap:12px; background:#1C222C; border:1px solid #2A3140; border-radius:14px; padding:14px; cursor:pointer; text-align:left; }
        .report-icon { width:38px; height:38px; border-radius:10px; background:rgba(79,209,174,0.12); color:#4FD1AE; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
        .report-text { flex:1; }
        .report-label { font-family:'Space Grotesk',sans-serif; font-size:14.5px; color:#E8EAED; font-weight:600; }
        .report-desc { font-size:12px; color:#5A6272; margin-top:2px; }
      `}</style>
    </div>
  );
}

function ReportSaldosBancos({ data, onBack }) {
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const total = data.bancos.reduce((s, b) => s + (Number(b.saldo) || 0), 0);
  const sorted = [...data.bancos].sort((a, b) => (Number(b.saldo) || 0) - (Number(a.saldo) || 0));
  return (
    <div className="report-body">
      <ReportHeader title="Saldos por banco" onBack={onBack} />
      <div className="rtotal">Total: <span className="mono">{fmtBRL(total)}</span> <ExportButton disabled={!sorted.length} onClick={() => exportToExcel(sorted, [
        { header: 'Banco', get: (b) => b.nome_banco, width: 24 }, { header: 'Empresa', get: (b) => empresaNome(b.empresa_id), width: 24 }, { header: 'Saldo', get: (b) => Number(b.saldo) || 0, width: 16 },
      ], 'Saldos por Banco', 'saldos-por-banco')} /></div>
      {sorted.length === 0 ? <EmptyState icon={Landmark} text="Nenhum banco cadastrado." /> : (
        <table className="rtable"><thead><tr><th>Banco</th><th>Empresa</th><th className="right">Saldo</th></tr></thead>
          <tbody>{sorted.map((b) => <tr key={b.id}><td>{b.nome_banco}</td><td className="dim">{empresaNome(b.empresa_id)}</td><td className="right mono">{fmtBRL(b.saldo)}</td></tr>)}</tbody>
        </table>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function ReportSaldosEmpresas({ data, onBack }) {
  const total = data.bancos.reduce((s, b) => s + (Number(b.saldo) || 0), 0);
  const rows = data.empresas.map((emp) => ({ nome: emp.nome, saldo: data.bancos.filter((b) => b.empresa_id === emp.id).reduce((s, b) => s + (Number(b.saldo) || 0), 0), numBancos: data.bancos.filter((b) => b.empresa_id === emp.id).length })).sort((a, b) => b.saldo - a.saldo);
  return (
    <div className="report-body">
      <ReportHeader title="Saldos por empresa" onBack={onBack} />
      <div className="rtotal">Total: <span className="mono">{fmtBRL(total)}</span> <ExportButton disabled={!rows.length} onClick={() => exportToExcel(rows, [
        { header: 'Empresa', get: (r) => r.nome, width: 24 }, { header: 'Bancos', get: (r) => r.numBancos, width: 12 }, { header: 'Saldo', get: (r) => r.saldo, width: 16 },
      ], 'Saldos por Empresa', 'saldos-por-empresa')} /></div>
      {rows.length === 0 ? <EmptyState icon={Building2} text="Nenhuma empresa cadastrada." /> : (
        <table className="rtable"><thead><tr><th>Empresa</th><th>Bancos</th><th className="right">Saldo</th></tr></thead>
          <tbody>{rows.map((r) => <tr key={r.nome}><td>{r.nome}</td><td className="dim">{r.numBancos}</td><td className="right mono">{fmtBRL(r.saldo)}</td></tr>)}</tbody>
        </table>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function ReportContasPagar({ data, onBack }) {
  const [empresaFiltro, setEmpresaFiltro] = useState('todas');
  const [statusFiltro, setStatusFiltro] = useState('pendente');
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const filtradas = useMemo(() => data.contas.filter((c) => empresaFiltro === 'todas' || c.empresa_id === empresaFiltro).filter((c) => statusFiltro === 'todas' || c.status === statusFiltro).sort((a, b) => a.data_vencimento.localeCompare(b.data_vencimento)), [data.contas, empresaFiltro, statusFiltro]);
  const total = filtradas.reduce((s, c) => s + (Number(c.valor) || 0), 0);
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
      </div>
      <div className="rtotal">Total ({filtradas.length}): <span className="mono">{fmtBRL(total)}</span> <ExportButton disabled={!filtradas.length} onClick={() => exportToExcel(filtradas, [
        { header: 'Descrição', get: (c) => c.descricao, width: 28 }, { header: 'Empresa', get: (c) => empresaNome(c.empresa_id), width: 24 },
        { header: 'Vencimento', get: (c) => fmtDate(c.data_vencimento), width: 14 },
        { header: 'Status', get: (c) => c.status === 'pago' ? 'Pago' : (daysUntil(c.data_vencimento) < 0 ? 'Atrasada' : 'Pendente'), width: 12 },
        { header: 'Valor', get: (c) => Number(c.valor) || 0, width: 16 },
      ], 'Contas a Pagar', 'contas-a-pagar')} /></div>
      {filtradas.length === 0 ? <EmptyState icon={CalendarClock} text="Nenhuma conta encontrada." /> : (
        <table className="rtable"><thead><tr><th>Descrição</th><th>Empresa</th><th>Vencimento</th><th className="right">Valor</th></tr></thead>
          <tbody>{filtradas.map((c) => { const atrasada = c.status === 'pendente' && daysUntil(c.data_vencimento) < 0; return (
            <tr key={c.id}><td>{c.descricao}</td><td className="dim">{empresaNome(c.empresa_id)}</td><td style={{ color: atrasada ? '#E2596B' : undefined }}>{fmtDate(c.data_vencimento)}</td><td className="right mono">{fmtBRL(c.valor)}</td></tr>
          ); })}</tbody>
        </table>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function ReportPagamentos({ data, onBack }) {
  const [empresaFiltro, setEmpresaFiltro] = useState('todas');
  const empresaNome = (id) => data.empresas.find((e) => e.id === id)?.nome || '—';
  const pagas = useMemo(() => data.contas.filter((c) => c.status === 'pago').filter((c) => empresaFiltro === 'todas' || c.empresa_id === empresaFiltro).sort((a, b) => (b.data_pagamento || '').localeCompare(a.data_pagamento || '')), [data.contas, empresaFiltro]);
  const total = pagas.reduce((s, c) => s + (Number(c.valor) || 0), 0);
  return (
    <div className="report-body">
      <ReportHeader title="Pagamentos realizados" onBack={onBack} />
      <div className="report-filters" style={{ marginBottom: 10 }}>
        <select className="select" value={empresaFiltro} onChange={(e) => setEmpresaFiltro(e.target.value)}>
          <option value="todas">Todas as empresas</option>
          {data.empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}
        </select>
      </div>
      <div className="rtotal">Total pago ({pagas.length}): <span className="mono">{fmtBRL(total)}</span> <ExportButton disabled={!pagas.length} onClick={() => exportToExcel(pagas, [
        { header: 'Descrição', get: (c) => c.descricao, width: 28 }, { header: 'Empresa', get: (c) => empresaNome(c.empresa_id), width: 24 },
        { header: 'Pago em', get: (c) => fmtDate(c.data_pagamento), width: 14 }, { header: 'Valor', get: (c) => Number(c.valor) || 0, width: 16 },
      ], 'Pagamentos', 'pagamentos-realizados')} /></div>
      {pagas.length === 0 ? <EmptyState icon={CheckCircle2} text="Nenhum pagamento registrado ainda." /> : (
        <table className="rtable"><thead><tr><th>Descrição</th><th>Empresa</th><th>Pago em</th><th className="right">Valor</th></tr></thead>
          <tbody>{pagas.map((c) => <tr key={c.id}><td>{c.descricao}</td><td className="dim">{empresaNome(c.empresa_id)}</td><td>{fmtDate(c.data_pagamento)}</td><td className="right mono">{fmtBRL(c.valor)}</td></tr>)}</tbody>
        </table>
      )}
      <RptStyle /><ViewStyle />
    </div>
  );
}

function RptStyle() {
  return <style jsx global>{`.rtotal { display:flex; align-items:center; gap:10px; margin:12px 20px 14px; font-size:13px; color:#8891A0; } .rtotal .mono { color:#E8EAED; font-size:15px; font-weight:600; }`}</style>;
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
          <Field label="Saldo atual (R$)"><input className="input" type="number" step="0.01" value={saldo} onChange={(e) => setSaldo(e.target.value)} placeholder="0,00" /></Field>
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
          <Field label="Valor (R$)"><input className="input" type="number" step="0.01" value={valor} onChange={(e) => setValor(e.target.value)} placeholder="0,00" /></Field>
          <Field label="Data de vencimento"><input className="input" type="date" value={dataVencimento} onChange={(e) => setDataVencimento(e.target.value)} /></Field>
          <PrimaryButton full onClick={() => canSave && onSave({ id: item?.id, empresaId, descricao: descricao.trim(), valor: parseFloat(valor) || 0, dataVencimento, status: item?.status || 'pendente', dataPagamento: item?.dataPagamento || null })}>{item ? 'Salvar alterações' : 'Salvar conta'}</PrimaryButton>
        </>
      )}
      <style jsx>{inputCss}</style>
    </Modal>
  );
}

function EmprestimoForm({ item, empresas, onClose, onSave }) {
  const [empresaId, setEmpresaId] = useState(item?.empresaId || empresas[0]?.id || '');
  const [credor, setCredor] = useState(item?.credor || '');
  const [valorTotal, setValorTotal] = useState(item?.valorTotal ?? '');
  const [valorParcela, setValorParcela] = useState(item?.valorParcela ?? '');
  const [numParcelas, setNumParcelas] = useState(item?.numParcelas ?? '');
  const [dataInicio, setDataInicio] = useState(item?.dataInicio || todayISO());
  const canSave = empresaId && credor.trim() && valorTotal;
  return (
    <Modal title={item ? 'Editar empréstimo' : 'Novo empréstimo'} onClose={onClose}>
      {empresas.length === 0 ? <p className="muted2">Cadastre uma empresa antes.</p> : (
        <>
          <Field label="Empresa"><select className="select" value={empresaId} onChange={(e) => setEmpresaId(e.target.value)}>{empresas.map((emp) => <option key={emp.id} value={emp.id}>{emp.nome}</option>)}</select></Field>
          <Field label="Credor / instituição"><input className="input" value={credor} onChange={(e) => setCredor(e.target.value)} placeholder="Ex: Banco do Brasil" /></Field>
          <Field label="Valor total (R$)"><input className="input" type="number" step="0.01" value={valorTotal} onChange={(e) => setValorTotal(e.target.value)} placeholder="0,00" /></Field>
          <Field label="Valor da parcela (R$)"><input className="input" type="number" step="0.01" value={valorParcela} onChange={(e) => setValorParcela(e.target.value)} placeholder="0,00" /></Field>
          <Field label="Número de parcelas"><input className="input" type="number" value={numParcelas} onChange={(e) => setNumParcelas(e.target.value)} placeholder="Ex: 12" /></Field>
          <Field label="Data de início"><input className="input" type="date" value={dataInicio} onChange={(e) => setDataInicio(e.target.value)} /></Field>
          <PrimaryButton full onClick={() => canSave && onSave({ id: item?.id, empresaId, credor: credor.trim(), valorTotal: parseFloat(valorTotal) || 0, valorParcela: parseFloat(valorParcela) || 0, numParcelas: parseInt(numParcelas) || 0, dataInicio })}>{item ? 'Salvar alterações' : 'Salvar empréstimo'}</PrimaryButton>
        </>
      )}
      <style jsx>{inputCss}</style>
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
      <Field label="Valor (R$)"><input className="input" type="number" step="0.01" value={valor} onChange={(e) => setValor(e.target.value)} placeholder="0,00" /></Field>
      <Field label="Data"><input className="input" type="date" value={data} onChange={(e) => setDataCampo(e.target.value)} /></Field>
      <Field label="Descrição (opcional)"><input className="input" value={descricao} onChange={(e) => setDescricao(e.target.value)} placeholder="Ex: Reforço de caixa" /></Field>
      {erro && <p style={{ color: '#E2596B', fontSize: 13, margin: 0 }}>{erro}</p>}
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
      <Field label="Taxa Pix (%)"><input className="input" type="number" step="0.01" value={taxaPix} onChange={(e) => setTaxaPix(e.target.value)} placeholder="0,00" /></Field>
      <Field label="Taxa Débito (%)"><input className="input" type="number" step="0.01" value={taxaDebito} onChange={(e) => setTaxaDebito(e.target.value)} placeholder="0,00" /></Field>
      <Field label="Taxa Crédito à vista (%)"><input className="input" type="number" step="0.01" value={taxaCreditoAvista} onChange={(e) => setTaxaCreditoAvista(e.target.value)} placeholder="0,00" /></Field>
      <Field label="Taxas de crédito parcelado (%)">
        <div className="pgrid">
          {PARCELAS_RANGE.map((n) => (
            <div key={n} className="pitem">
              <span>{n}x</span>
              <input className="input" type="number" step="0.01" value={parcelas[n]} onChange={(e) => setParcelas((p) => ({ ...p, [n]: e.target.value }))} placeholder="0,00" />
            </div>
          ))}
        </div>
      </Field>
      <PrimaryButton full onClick={() => canSave && handleSave()}>{item ? 'Salvar alterações' : 'Salvar taxa'}</PrimaryButton>
      <style jsx>{inputCss}</style>
      <style jsx>{`
        .pgrid { display:grid; grid-template-columns:repeat(3,1fr); gap:8px; }
        .pitem { display:flex; flex-direction:column; gap:4px; }
        .pitem span { font-size:11px; color:#5A6272; }
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
      <Field label="Taxa de emissão (R$)"><input className="input" type="number" step="0.01" value={taxaEmissao} onChange={(e) => setTaxaEmissao(e.target.value)} placeholder="0,00" /></Field>
      <Field label="Taxa de baixa (R$)"><input className="input" type="number" step="0.01" value={taxaBaixa} onChange={(e) => setTaxaBaixa(e.target.value)} placeholder="0,00" /></Field>
      <Field label="Taxa de protesto (R$)"><input className="input" type="number" step="0.01" value={taxaProtesto} onChange={(e) => setTaxaProtesto(e.target.value)} placeholder="0,00" /></Field>
      <Field label="Taxa de antecipação (% ao mês)"><input className="input" type="number" step="0.01" value={taxaAntecipacao} onChange={(e) => setTaxaAntecipacao(e.target.value)} placeholder="0,00" /></Field>
      <PrimaryButton full onClick={() => canSave && onSave({ id: item?.id, administradora: administradora.trim(), taxaEmissao: taxaEmissao === '' ? null : parseFloat(taxaEmissao), taxaBaixa: taxaBaixa === '' ? null : parseFloat(taxaBaixa), taxaProtesto: taxaProtesto === '' ? null : parseFloat(taxaProtesto), taxaAntecipacao: taxaAntecipacao === '' ? null : parseFloat(taxaAntecipacao) })}>{item ? 'Salvar alterações' : 'Salvar taxa'}</PrimaryButton>
      <style jsx>{inputCss}</style>
    </Modal>
  );
}
