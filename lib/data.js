import { supabase } from './supabaseClient';

// ---------- Fetch all ----------
export async function fetchAll() {
  const [empresas, bancos, contas, emprestimos, parcelas, transferencias, taxasCartao, taxasBoleto, seguros, vendas] = await Promise.all([
    supabase.from('empresas').select('*').order('nome'),
    supabase.from('bancos').select('*'),
    supabase.from('contas').select('*'),
    supabase.from('emprestimos').select('*'),
    supabase.from('parcelas_emprestimo').select('*'),
    supabase.from('transferencias').select('*'),
    supabase.from('taxas_cartao').select('*'),
    supabase.from('taxas_boleto').select('*'),
    supabase.from('seguros').select('*'),
    supabase.from('vendas').select('*'),
  ]);
  const err = [empresas, bancos, contas, emprestimos, parcelas, transferencias, taxasCartao, taxasBoleto, seguros, vendas].find((r) => r.error);
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
    seguros: seguros.data,
    vendas: vendas.data,
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

// Dar baixa: marca como paga numa data específica (ou reverte para pendente)
export async function darBaixaConta(conta, dataPagamento) {
  return supabase.from('contas').update({ status: 'pago', data_pagamento: dataPagamento }).eq('id', conta.id);
}
export async function reverterBaixaConta(conta) {
  return supabase.from('contas').update({ status: 'pendente', data_pagamento: null }).eq('id', conta.id);
}

// ---------- Empréstimos + parcelas ----------
function addMonths(iso, n) {
  const [y, m, d] = iso.split('-').map(Number);
  const date = new Date(y, (m - 1) + n, d);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

export async function saveEmprestimo({ id, empresaId, credor, valorTotal, valorParcela, numParcelas, dataInicio, tipo, linhaCredito, empresaCredoraId, temCarencia, prazoCarencia }, oldParcelas = []) {
  const row = {
    empresa_id: empresaId, credor, valor_total: valorTotal, valor_parcela: valorParcela, num_parcelas: numParcelas, data_inicio: dataInicio,
    tipo: tipo || 'banco',
    linha_credito: tipo === 'banco' ? (linhaCredito || null) : null,
    empresa_credora_id: tipo === 'empresa' ? (empresaCredoraId || null) : null,
    tem_carencia: !!temCarencia,
    prazo_carencia: temCarencia ? (parseInt(prazoCarencia, 10) || 0) : 0,
  };
  let emprestimoId = id;
  if (id) {
    await supabase.from('emprestimos').update(row).eq('id', id);
  } else {
    const { data, error } = await supabase.from('emprestimos').insert(row).select().single();
    if (error) throw error;
    emprestimoId = data.id;
  }

  // Regenerate installment schedule, preserving paid status by installment number.
  // "numParcelas" representa o PRAZO TOTAL do contrato em meses (carência + amortização).
  // A carência é o período inicial sem pagamento; a amortização ocupa o restante do prazo.
  await supabase.from('parcelas_emprestimo').delete().eq('emprestimo_id', emprestimoId);
  const carenciaMeses = temCarencia ? (parseInt(prazoCarencia, 10) || 0) : 0;
  const prazoTotalMeses = Number(numParcelas) || 0;
  const parcelasAmortizacao = Math.max(prazoTotalMeses - carenciaMeses, 0);
  const novasParcelas = [];
  for (let i = 1; i <= parcelasAmortizacao; i++) {
    const antiga = oldParcelas.find((p) => p.numero === i);
    novasParcelas.push({
      emprestimo_id: emprestimoId,
      numero: i,
      data_vencimento: addMonths(dataInicio, carenciaMeses + i - 1),
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

// ---------- Seguros ----------
export async function saveSeguro({ id, empresaId, objeto, vigenciaInicio, vigenciaFim, principalCondutor, seguradora, formaPagamento, numParcelas, valorParcela, valorTotal }) {
  const row = {
    empresa_id: empresaId,
    objeto,
    vigencia_inicio: vigenciaInicio,
    vigencia_fim: vigenciaFim,
    principal_condutor: principalCondutor || null,
    seguradora,
    forma_pagamento: formaPagamento || 'boleto',
    num_parcelas: numParcelas || null,
    valor_parcela: valorParcela || null,
    valor_total: valorTotal,
  };
  if (id) return supabase.from('seguros').update(row).eq('id', id);
  return supabase.from('seguros').insert(row);
}
export async function deleteSeguro(id) {
  return supabase.from('seguros').delete().eq('id', id);
}

// ---------- Vendas ----------
// Um registro por (empresa, setor, ano, mês) — salvar de novo atualiza o existente.
export async function saveVendaMes({ empresaId, setor, ano, mes, valor, observacoes }) {
  return supabase.from('vendas').upsert(
    { empresa_id: empresaId, setor, ano, mes, valor: valor || 0, observacoes: observacoes || null },
    { onConflict: 'empresa_id,setor,ano,mes' }
  );
}
