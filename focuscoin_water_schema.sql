-- ============================================================
-- FocusCoin 일일미션: 물 마시기 (challenge_water) 스키마 v1
-- 실행 방법: Supabase 대시보드 → SQL Editor → 전체 붙여넣기 → Run
-- 프로젝트: gjnwriqewsrwpbtxqbea
-- ============================================================

create extension if not exists pgcrypto;

-- 1) 오늘의 물 기록 (유저 x 날짜당 1행)
create table if not exists public.water_logs (
    id uuid primary key default gen_random_uuid(),
    user_id text not null,
    log_date date not null,
    cups integer not null default 0 check (cups >= 0 and cups <= 8),
    claimed boolean not null default false,
    last_added_at timestamptz,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    unique (user_id, log_date)
);
create index if not exists water_logs_user_date_idx
    on public.water_logs (user_id, log_date);

-- 2) 미션 완료 기록 (보상 적립 원장 — 영양제, 영단어 미션도 재사용 가능)
create table if not exists public.mission_completions (
    id uuid primary key default gen_random_uuid(),
    user_id text not null,
    mission_type text not null,
    completed_date date not null,
    reward_coins integer not null default 0,
    created_at timestamptz not null default now(),
    unique (user_id, mission_type, completed_date)
);

-- 3) 보안: 테이블 직접 접근 전면 차단 (조작은 아래 RPC 함수로만 가능)
alter table public.water_logs enable row level security;
alter table public.mission_completions enable row level security;
revoke all on public.water_logs from anon, authenticated;
revoke all on public.mission_completions from anon, authenticated;

-- 4) RPC: 오늘 상태 조회
create or replace function public.get_water_status(p_user_id text)
returns json
language plpgsql security definer set search_path = public
as $$
declare
    v_today date := (now() at time zone 'Asia/Seoul')::date;
    v_row public.water_logs%rowtype;
    v_remain integer := 0;
begin
    if p_user_id is null or length(trim(p_user_id)) = 0 then
        return json_build_object('ok', false, 'error', 'invalid_user');
    end if;
    select * into v_row from public.water_logs
     where user_id = p_user_id and log_date = v_today;
    if not found then
        return json_build_object('ok', true, 'cups', 0, 'claimed', false, 'cooldown_remain_sec', 0);
    end if;
    if v_row.last_added_at is not null then
        v_remain := greatest(0, 300 - floor(extract(epoch from (now() - v_row.last_added_at)))::integer);
    end if;
    return json_build_object('ok', true, 'cups', v_row.cups, 'claimed', v_row.claimed, 'cooldown_remain_sec', v_remain);
end;
$$;

-- 5) RPC: 물 등록 (서버에서 5분 쿨다운 강제 — 클라이언트 조작 불가)
create or replace function public.add_water_cups(p_user_id text, p_cups integer)
returns json
language plpgsql security definer set search_path = public
as $$
declare
    v_today date := (now() at time zone 'Asia/Seoul')::date;
    v_row public.water_logs%rowtype;
    v_remain integer;
begin
    if p_user_id is null or length(trim(p_user_id)) = 0 then
        return json_build_object('ok', false, 'error', 'invalid_user');
    end if;
    if p_cups is null or p_cups < 1 or p_cups > 4 then
        return json_build_object('ok', false, 'error', 'invalid_cups');
    end if;

    insert into public.water_logs (user_id, log_date, cups)
    values (p_user_id, v_today, 0)
    on conflict (user_id, log_date) do nothing;

    select * into v_row from public.water_logs
     where user_id = p_user_id and log_date = v_today
     for update;

    if v_row.claimed then
        return json_build_object('ok', false, 'error', 'already_claimed', 'cups', v_row.cups, 'claimed', true);
    end if;
    if v_row.cups >= 8 then
        return json_build_object('ok', false, 'error', 'already_full', 'cups', v_row.cups, 'claimed', v_row.claimed);
    end if;
    if v_row.last_added_at is not null
       and now() - v_row.last_added_at < interval '5 minutes' then
        v_remain := greatest(0, 300 - floor(extract(epoch from (now() - v_row.last_added_at)))::integer);
        return json_build_object('ok', false, 'error', 'cooldown', 'cooldown_remain_sec', v_remain, 'cups', v_row.cups, 'claimed', v_row.claimed);
    end if;

    update public.water_logs
       set cups = least(v_row.cups + p_cups, 8),
           last_added_at = now(),
           updated_at = now()
     where id = v_row.id
     returning * into v_row;

    return json_build_object('ok', true, 'cups', v_row.cups, 'claimed', v_row.claimed, 'cooldown_remain_sec', 300);
end;
$$;

-- 6) RPC: 보상 수령 (8잔 완료 + 하루 1회만)
create or replace function public.claim_water_reward(p_user_id text)
returns json
language plpgsql security definer set search_path = public
as $$
declare
    v_today date := (now() at time zone 'Asia/Seoul')::date;
    v_row public.water_logs%rowtype;
    v_count integer := 0;
begin
    if p_user_id is null or length(trim(p_user_id)) = 0 then
        return json_build_object('ok', false, 'error', 'invalid_user');
    end if;

    select * into v_row from public.water_logs
     where user_id = p_user_id and log_date = v_today
     for update;

    if not found or v_row.cups < 8 then
        return json_build_object('ok', false, 'error', 'not_full');
    end if;
    if v_row.claimed then
        return json_build_object('ok', false, 'error', 'duplicate');
    end if;

    insert into public.mission_completions (user_id, mission_type, completed_date, reward_coins)
    values (p_user_id, 'DAILY_WATER_CUPS', v_today, 10)
    on conflict (user_id, mission_type, completed_date) do nothing;
    get diagnostics v_count = row_count;

    update public.water_logs
       set claimed = true, updated_at = now()
     where id = v_row.id;

    if v_count = 0 then
        return json_build_object('ok', false, 'error', 'duplicate');
    end if;

    return json_build_object('ok', true, 'reward_coins', 10);
end;
$$;

-- 7) anon 키로는 RPC 함수만 호출 가능
grant execute on function public.get_water_status(text) to anon;
grant execute on function public.add_water_cups(text, integer) to anon;
grant execute on function public.claim_water_reward(text) to anon;
