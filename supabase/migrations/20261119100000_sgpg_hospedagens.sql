-- SGPG — módulo Hospedagens: as estadias na Casa da Fazenda Pao Grande (a única
-- acomodação é a Casa).
--
-- A casa recebe a família (no DRE da Casa: "Hospedagem Família" de cada núcleo),
-- locação (receita "Locação da Casa") e visitas ("Despesas com Visitas"). Até aqui nada
-- disso tinha registro próprio: não havia agenda do que vem pela frente nem histórico de
-- quem ficou, quando e quantos. Este módulo é o lugar: uma linha por estadia, com a
-- agenda futura e o histórico saindo da mesma tabela.
--
-- Mesma receita do módulo de exposições (migration 20261005150000): quem vê é
-- hub_can('sgpg'); quem grava é hub_sgpg_editor() (admin e sgpg_editores); autoria
-- posta pelo banco (sgpg_autoria). Repo público: nome de hóspede só no banco.

create table if not exists sgpg_hospedagem (
  id            bigint generated always as identity primary key,
  -- de onde vem a estadia; casa com as naturezas do DRE da Casa
  tipo          text not null default 'Família'
                check (tipo in ('Família','Locação','Visita','Outro')),
  -- quem se hospeda: pessoa, núcleo da família ou grupo, como a casa chama
  hospede       text not null,
  entrada       date not null,
  saida         date not null,
  adultos       int not null default 0 check (adultos >= 0),
  criancas      int not null default 0 check (criancas >= 0),
  -- manual, como nas exposições: Prevista (a confirmar), Confirmada, Realizada, Cancelada
  status        text not null default 'Prevista'
                check (status in ('Prevista','Confirmada','Realizada','Cancelada')),
  -- locação: o valor cobrado
  valor         numeric(12,2) check (valor is null or valor >= 0),
  observacao    text,
  criado_por    text,
  criado_em     timestamptz not null default now(),
  alterado_por  text,
  alterado_em   timestamptz not null default now(),
  check (saida >= entrada)
);
create index if not exists sgpg_hospedagem_entrada on sgpg_hospedagem (entrada);
create index if not exists sgpg_hospedagem_saida on sgpg_hospedagem (saida);

drop trigger if exists sgpg_hospedagem_autoria on sgpg_hospedagem;
create trigger sgpg_hospedagem_autoria before insert or update on sgpg_hospedagem
  for each row execute function public.sgpg_autoria();

alter table sgpg_hospedagem enable row level security;

drop policy if exists sgh_read on sgpg_hospedagem;
create policy sgh_read on sgpg_hospedagem for select to authenticated
  using ( public.hub_can('sgpg') );
drop policy if exists sgh_write on sgpg_hospedagem;
create policy sgh_write on sgpg_hospedagem for all to authenticated
  using ( public.hub_sgpg_editor() ) with check ( public.hub_sgpg_editor() );
