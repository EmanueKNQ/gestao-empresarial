import { exigirAdmin } from '../../../lib/adminServer';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método não permitido.' });

  const check = await exigirAdmin(req);
  if (!check.ok) return res.status(check.status).json({ error: check.message });

  const { userId, novaSenha } = req.body || {};
  if (!userId || !novaSenha) return res.status(400).json({ error: 'userId e novaSenha são obrigatórios.' });
  if (novaSenha.length < 6) return res.status(400).json({ error: 'A senha precisa ter pelo menos 6 caracteres.' });

  const { admin } = check;
  const { error } = await admin.auth.admin.updateUserById(userId, { password: novaSenha });
  if (error) return res.status(400).json({ error: error.message });

  return res.status(200).json({ ok: true });
}
