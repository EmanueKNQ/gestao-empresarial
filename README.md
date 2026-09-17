# Gestão Empresarial — App Web

App de gestão financeira (empresas, bancos, contas a pagar, empréstimos,
transferências, taxas de cartão e boleto) com Next.js + Supabase.

## Rodar localmente
1. `npm install`
2. Confirme que `.env.local` tem as credenciais do seu projeto Supabase.
3. `npm run dev` e acesse http://localhost:3000

## Atualização: novos campos de empréstimos

Antes de subir esta versão, rode `migracao-emprestimos.sql` no SQL Editor do
Supabase — ela adiciona os campos de tipo de empréstimo (bancário / entre
empresas), linha de crédito e empresa credora.

## Publicar no Vercel
1. Suba esta pasta para um repositório no GitHub (ou use "vercel deploy" via CLI).
2. Em vercel.com, "Add New Project" → importe o repositório.
3. Em "Environment Variables", adicione:
   - NEXT_PUBLIC_SUPABASE_URL
   - NEXT_PUBLIC_SUPABASE_ANON_KEY
   - SUPABASE_SERVICE_ROLE_KEY (Supabase → Project Settings → API → "service_role" secret.
     NUNCA coloque o prefixo NEXT_PUBLIC_ nessa variável — ela precisa ficar só no servidor.)
4. Deploy. Em ~1 minuto você recebe um link público (ex: seu-app.vercel.app).

## Login e controle de acesso
A partir desta versão o app exige login (e-mail e senha) e cada usuário só
enxerga as abas liberadas pelo administrador. Veja `migracao-acesso-usuarios.sql`
para criar seu primeiro usuário administrador.
