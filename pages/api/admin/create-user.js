import { exigirAdmin } from '../../../lib/adminServer';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método não permitido.' });

  const check = await exigirAdmin(req);
  if (!check.ok) return res.status(check.status).json({ error: check.message });

  const { nome, email, senha, papel, abas } = req.body || {};
  if (!nome || !email || !senha) return res.status(400).json({ error: 'Nome, e-mail e senha são obrigatórios.' });
  if (senha.length < 6) return res.status(400).json({ error: 'A senha precisa ter pelo menos 6 caracteres.' });

  const { admin } = check;

  const { data: created, error: createErr } = await admin.auth.admin.createUser({
    email,
    password: senha,
    email_confirm: true,
  });
  if (createErr) return res.status(400).json({ error: createErr.message });

  const { error: perfilErr } = await admin.from('perfis').insert({
    id: created.user.id,
    nome,
    email,
    papel: papel === 'admin' ? 'admin' : 'usuario',
    abas: Array.isArray(abas) ? abas : [],
  });
  if (perfilErr) {
    // reverte a criação do usuário se não conseguiu salvar o perfil
    await admin.auth.admin.deleteUser(created.user.id);
    return res.status(400).json({ error: perfilErr.message });
  }

  return res.status(200).json({ ok: true });
}
