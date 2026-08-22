import { supabase } from './supabaseClient';

// ---------- Fetch all ----------
export async function fetchAll() {
  const [empresas, bancos, contas, emprestimos, parcelas, transferencias, taxasCartao, taxasBoleto] = await Promise.all([
    supabase.from('empresas').select('*').order('nome'),
    supabase.from('bancos').select('*'),
    supabase.from('contas').select('*'),
    supabase.from('emprestimos').select('*'),
    supabase.from('parcelas_emprestimo').select('*'),
    supabase.from('transferencias').select('*'),
    supabase.from('taxas_cartao').select('*'),
    supabase.from('taxas_boleto').select('*'),
  ]);
  const err = [empresas, bancos, contas, emprestimos, parcelas, transferencias, taxasCartao, taxasBoleto].find((r) => r.error);
  if (err) throw err.error;

  // attach parcelas to their emprestimo
  const emprestimosComParcelas = emprestimos.data.map((e) => ({
    ...e,
    parcelas: parcelas.data.filter((p) => p.emprestimo_id === e.id).sort((a, b) => a.numero - b.numero),
  }));

  return {
    empresas: empresas.data,
    bancos: bancos.data,
    contas: contas.data,
    emprestimos: emprestimosComParcelas,
    transferencias: transferencias.data,
    taxasCartao: taxasCartao.data,
    taxasBoleto: taxasBoleto.data,
  };
}

// ---------- Empresas ----------
export async function saveEmpresa({ id, nome }) {
  if (id) return supabase.from('empresas').update({ nome }).eq('id', id);
  return supabase.from('empresas').insert({ nome });
}
export async function deleteEmpresa(id) {
  return supabase.from('empresas').delete().eq('id', id);
}

// ---------- Bancos ----------
export async function saveBanco({ id, empresaId, nomeBanco, saldo }) {
  const row = { empresa_id: empresaId, nome_banco: nomeBanco, saldo };
  if (id) return supabase.from('bancos').update(row).eq('id', id);
  return supabase.from('bancos').insert(row);
}
export async function deleteBanco(id) {
  return supabase.from('bancos').delete().eq('id', id);
}

// ---------- Contas a pagar ----------
export async function saveConta({ id, empresaId, descricao, valor, dataVencimento, status, dataPagamento }) {
  const row = { empresa_id: empresaId, descricao, valor, data_vencimento: dataVencimento, status, data_pagamento: dataPagamento || null };
  if (id) return supabase.from('contas').update(row).eq('id', id);
  return supabase.from('contas').insert(row);
}
export async function deleteConta(id) {
  return supabase.from('contas').delete().eq('id', id);
}
export async function toggleContaStatus(conta) {
  const novoStatus = conta.status === 'pendente' ? 'pago' : 'pendente';
  return supabase.from('contas').update({
    status: novoStatus,
    data_pagamento: novoStatus === 'pago' ? new Date().toISOString().slice(0, 10) : null,
  }).eq('id', conta.id);
}

// ---------- Empréstimos + parcelas ----------
function addMonths(iso, n) {
  const [y, m, d] = iso.split('-').map(Number);
  const date = new Date(y, (m - 1) + n, d);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

export async function saveEmprestimo({ id, empresaId, credor, valorTotal, valorParcela, numParcelas, dataInicio }, oldParcelas = []) {
  const row = { empresa_id: empresaId, credor, valor_total: valorTotal, valor_parcela: valorParcela, num_parcelas: numParcelas, data_inicio: dataInicio };
  let emprestimoId = id;
  if (id) {
    await supabase.from('emprestimos').update(row).eq('id', id);
  } else {
    const { data, error } = await supabase.from('emprestimos').insert(row).select().single();
    if (error) throw error;
    emprestimoId = data.id;
  }

  // Regenerate installment schedule, preserving paid status by installment number
  await supabase.from('parcelas_emprestimo').delete().eq('emprestimo_id', emprestimoId);
  const novasParcelas = [];
  for (let i = 1; i <= Number(numParcelas); i++) {
    const antiga = oldParcelas.find((p) => p.numero === i);
    novasParcelas.push({
      emprestimo_id: emprestimoId,
      numero: i,
      data_vencimento: addMonths(dataInicio, i - 1),
      valor: valorParcela,
      status: antiga?.status || 'pendente',
      data_pagamento: antiga?.data_pagamento || null,
    });
  }
  if (novasParcelas.length) await supabase.from('parcelas_emprestimo').insert(novasParcelas);
  return emprestimoId;
}
export async function deleteEmprestimo(id) {
  return supabase.from('emprestimos').delete().eq('id', id);
}
export async function toggleParcela(parcela) {
  const novoStatus = parcela.status === 'pendente' ? 'pago' : 'pendente';
  return supabase.from('parcelas_emprestimo').update({
    status: novoStatus,
    data_pagamento: novoStatus === 'pago' ? new Date().toISOString().slice(0, 10) : null,
  }).eq('id', parcela.id);
}

// ---------- Transferências ----------
export async function criarTransferencia({ bancoOrigemId, bancoDestinoId, valor, data, descricao }, bancos) {
  const origem = bancos.find((b) => b.id === bancoOrigemId);
  const destino = bancos.find((b) => b.id === bancoDestinoId);
  await supabase.from('bancos').update({ saldo: Number(origem.saldo) - valor }).eq('id', bancoOrigemId);
  await supabase.from('bancos').update({ saldo: Number(destino.saldo) + valor }).eq('id', bancoDestinoId);
  return supabase.from('transferencias').insert({ banco_origem_id: bancoOrigemId, banco_destino_id: bancoDestinoId, valor, data, descricao: descricao || null });
}
export async function excluirTransferencia(transferencia, bancos) {
  const origem = bancos.find((b) => b.id === transferencia.banco_origem_id);
  const destino = bancos.find((b) => b.id === transferencia.banco_destino_id);
  if (origem) await supabase.from('bancos').update({ saldo: Number(origem.saldo) + Number(transferencia.valor) }).eq('id', origem.id);
  if (destino) await supabase.from('bancos').update({ saldo: Number(destino.saldo) - Number(transferencia.valor) }).eq('id', destino.id);
  return supabase.from('transferencias').delete().eq('id', transferencia.id);
}

// ---------- Taxas de cartão ----------
export async function saveTaxaCartao({ id, administradora, bandeira, taxaPix, taxaDebito, taxaCreditoAvista, parcelas }) {
  const row = { administradora, bandeira, taxa_pix: taxaPix, taxa_debito: taxaDebito, taxa_credito_avista: taxaCreditoAvista, parcelas };
  if (id) return supabase.from('taxas_cartao').update(row).eq('id', id);
  return supabase.from('taxas_cartao').insert(row);
}
export async function deleteTaxaCartao(id) {
  return supabase.from('taxas_cartao').delete().eq('id', id);
}

// ---------- Taxas de boleto ----------
export async function saveTaxaBoleto({ id, administradora, taxaEmissao, taxaBaixa, taxaProtesto, taxaAntecipacao }) {
  const row = { administradora, taxa_emissao: taxaEmissao, taxa_baixa: taxaBaixa, taxa_protesto: taxaProtesto, taxa_antecipacao: taxaAntecipacao };
  if (id) return supabase.from('taxas_boleto').update(row).eq('id', id);
  return supabase.from('taxas_boleto').insert(row);
}
export async function deleteTaxaBoleto(id) {
  return supabase.from('taxas_boleto').delete().eq('id', id);
}
