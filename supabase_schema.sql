create extension if not exists pgcrypto;

create table if not exists public.execucoes (
    id uuid primary key default gen_random_uuid(),
    nome_arquivo text not null,
    total_registros integer not null,
    total_transportadores integer not null,
    total_comando integer not null,
    total_ramal integer not null,
    status text not null default 'concluido',
    criado_em timestamptz not null default now()
);

-- A integração com Supabase fica desativada no primeiro MVP.
-- Ative autenticação e políticas RLS antes de gravar dados de usuários.
