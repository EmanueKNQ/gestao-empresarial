import { exigirAdmin } from '../../../lib/adminServer';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método não permitido.' });

  const check = await exigirAdmin(req);
  if (!check.ok) return res.status(check.status).json({ error: check.message });

  const { userId } = req.body || {};
  if (!userId) return res.status(400).json({ error: 'userId é obrigatório.' });
  if (userId === check.userId) return res.status(400).json({ error: 'Você não pode excluir o seu próprio usuário.' });

  const { admin } = check;
  const { error } = await admin.auth.admin.deleteUser(userId); // exclui de auth.users; perfis é apagado em cascata
  if (error) return res.status(400).json({ error: error.message });

  return res.status(200).json({ ok: true });
}
