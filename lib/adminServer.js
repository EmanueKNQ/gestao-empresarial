import { createClient } from '@supabase/supabase-js';

// Cliente com a service_role key — só roda no servidor (rotas /api), nunca no navegador.
export function getAdminClient() {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!serviceKey) throw new Error('SUPABASE_SERVICE_ROLE_KEY não configurada no servidor.');
  return createClient(url, serviceKey, { auth: { autoRefreshToken: false, persistSession: false } });
}

// Confirma que quem está chamando a rota é um administrador autenticado.
// Lê o token enviado no header Authorization e verifica o papel no perfil.
export async function exigirAdmin(req) {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.replace('Bearer ', '');
  if (!token) return { ok: false, status: 401, message: 'Não autenticado.' };

  const admin = getAdminClient();
  const { data: userData, error: userErr } = await admin.auth.getUser(token);
  if (userErr || !userData?.user) return { ok: false, status: 401, message: 'Sessão inválida.' };

  const { data: perfil, error: perfilErr } = await admin.from('perfis').select('papel').eq('id', userData.user.id).single();
  if (perfilErr || perfil?.papel !== 'admin') return { ok: false, status: 403, message: 'Apenas administradores podem fazer isso.' };

  return { ok: true, admin, userId: userData.user.id };
}
