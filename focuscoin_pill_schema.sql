-- ============================================================
-- FocusCoin 일일미션: 영양제 챙기기 (challenge_pill) 스키마 v1
-- 실행 방법: Supabase 대시보드 → SQL Editor → 전체 붙여넣기 → Run
-- 프로젝트: gjnwriqewsrwpbtxqbea
-- ============================================================
-- 의존성: mission_completions 테이블 (water_schema가 먼저 만들었어야 함)
--   안 만들었어도 아래 if not exists로 안전하게 생성됨
-- 시간대: Asia/Seoul 기준 (water_schema와 동일)
-- ============================================================

create extension if not exists pgcrypto;

-- ============================================================
-- 1) 사용자별 영양제 마스터
-- ============================================================
create table if not exists public.pill_items (
    id uuid primary key default gen_random_uuid(),
    user_id text not null,
    name text not null,
    dose text,
    time_slot text,                          -- '아침'/'점심'/'저녁'/'자기 전' 또는 null
    sort_order integer not null default 0,
    deleted_at timestamptz,                  -- soft delete (null = 활성)
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);
create index if not exists pill_items_user_idx
    on public.pill_items (user_id)
    where deleted_at is null;

-- ============================================================
-- 2) 영양제 × 날짜 챙김 기록 (존재 = 챙김 / 없음 = 미챙김)
-- ============================================================
create table if not exists public.pill_daily_logs (
    id uuid primary key default gen_random_uuid(),
    user_id text not null,
    pill_id uuid not null references public.pill_items(id) on delete cascade,
    log_date date not null,
    taken_at timestamptz not null default now(),
    unique (pill_id, log_date)               -- 같은 영양제 같은 날짜 = 1행
);
create index if not exists pill_daily_logs_user_date_idx
    on public.pill_daily_logs (user_id, log_date);

-- ============================================================
-- 3) 미션 완료 공용 원장 (water_schema가 이미 정의했을 수 있음)
--    재사용 안전 (if not exists)
-- ============================================================
create table if not exists public.mission_completions (
    id uuid primary key default gen_random_uuid(),
    user_id text not null,
    mission_type text not null,              -- 'DAILY_WATER_CUPS' / 'DAILY_PILL_TAKEN' / ...
    completed_date date not null,
    reward_coins integer not null default 0,
    created_at timestamptz not null default now(),
    unique (user_id, mission_type, completed_date)
);

-- ============================================================
-- 4) 보안: 테이블 직접 접근 차단 (RPC 함수로만 조작 가능)
-- ============================================================
alter table public.pill_items enable row level security;
alter table public.pill_daily_logs enable row level security;
alter table public.mission_completions enable row level security;
revoke all on public.pill_items from anon, authenticated;
revoke all on public.pill_daily_logs from anon, authenticated;
revoke all on public.mission_completions from anon, authenticated;

-- ============================================================
-- 5) RPC: 오늘 상태 조회
--    응답: { ok, items, taken_ids, claimed }
-- ============================================================
create or replace function public.get_pill_status(p_user_id text)
returns json
language plpgsql security definer set search_path = public
as $$
declare
    v_today date := (now() at time zone 'Asia/Seoul')::date;
    v_items json;
    v_taken_ids json;
    v_claimed boolean := false;
begin
    if p_user_id is null or length(trim(p_user_id)) = 0 then
        return json_build_object('ok', false, 'error', 'invalid_user');
    end if;

    -- 활성 영양제 목록 (sort_order, created_at 순)
    select coalesce(
        json_agg(
            json_build_object(
                'id', id,
                'name', name,
                'dose', dose,
                'time_slot', time_slot,
                'sort_order', sort_order
            )
            order by sort_order, created_at
        ),
        '[]'::json
    ) into v_items
    from public.pill_items
    where user_id = p_user_id and deleted_at is null;

    -- 오늘 체크한 pill_id 목록
    select coalesce(json_agg(pill_id), '[]'::json) into v_taken_ids
    from public.pill_daily_logs
    where user_id = p_user_id and log_date = v_today;

    -- 오늘 완료 여부
    select exists(
        select 1 from public.mission_completions
        where user_id = p_user_id
          and mission_type = 'DAILY_PILL_TAKEN'
          and completed_date = v_today
    ) into v_claimed;

    return json_build_object(
        'ok', true,
        'items', v_items,
        'taken_ids', v_taken_ids,
        'claimed', v_claimed
    );
end;
$$;

-- ============================================================
-- 6) RPC: 영양제 추가
--    name 필수 (1~20자), dose/time_slot 선택
-- ============================================================
create or replace function public.add_pill_item(
    p_user_id text,
    p_name text,
    p_dose text default null,
    p_time_slot text default null
)
returns json
language plpgsql security definer set search_path = public
as $$
declare
    v_id uuid;
    v_sort integer;
begin
    if p_user_id is null or length(trim(p_user_id)) = 0 then
        return json_build_object('ok', false, 'error', 'invalid_user');
    end if;
    if p_name is null or length(trim(p_name)) = 0 or length(p_name) > 20 then
        return json_build_object('ok', false, 'error', 'invalid_name');
    end if;
    if p_dose is not null and length(p_dose) > 20 then
        return json_build_object('ok', false, 'error', 'invalid_dose');
    end if;
    if p_time_slot is not null
       and p_time_slot not in ('아침','점심','저녁','자기 전') then
        return json_build_object('ok', false, 'error', 'invalid_time_slot');
    end if;

    -- 정렬 순서: 사용자별 마지막 + 1
    select coalesce(max(sort_order), 0) + 1 into v_sort
    from public.pill_items
    where user_id = p_user_id and deleted_at is null;

    insert into public.pill_items (user_id, name, dose, time_slot, sort_order)
    values (
        p_user_id,
        trim(p_name),
        nullif(trim(coalesce(p_dose, '')), ''),
        p_time_slot,
        v_sort
    )
    returning id into v_id;

    return json_build_object('ok', true, 'id', v_id);
end;
$$;

-- ============================================================
-- 7) RPC: 영양제 삭제 (soft delete)
--    IDOR 방지: 본인 소유 영양제만 삭제 가능
-- ============================================================
create or replace function public.delete_pill_item(
    p_user_id text,
    p_pill_id uuid
)
returns json
language plpgsql security definer set search_path = public
as $$
declare
    v_today date := (now() at time zone 'Asia/Seoul')::date;
    v_owner text;
begin
    if p_user_id is null or length(trim(p_user_id)) = 0 then
        return json_build_object('ok', false, 'error', 'invalid_user');
    end if;

    select user_id into v_owner
    from public.pill_items
    where id = p_pill_id and deleted_at is null;

    if v_owner is null then
        return json_build_object('ok', false, 'error', 'not_found');
    end if;
    if v_owner <> p_user_id then
        return json_build_object('ok', false, 'error', 'forbidden');
    end if;

    update public.pill_items
       set deleted_at = now(), updated_at = now()
     where id = p_pill_id;

    -- 오늘 체크 기록도 정리 (이전 날짜는 통계 유지)
    delete from public.pill_daily_logs
     where pill_id = p_pill_id and log_date = v_today;

    return json_build_object('ok', true);
end;
$$;

-- ============================================================
-- 8) RPC: 오늘 챙김 토글
--    p_taken=true → INSERT (멱등)
--    p_taken=false → DELETE
--    이미 완료(claimed) 상태면 변경 불가
-- ============================================================
create or replace function public.set_pill_taken(
    p_user_id text,
    p_pill_id uuid,
    p_taken boolean
)
returns json
language plpgsql security definer set search_path = public
as $$
declare
    v_today date := (now() at time zone 'Asia/Seoul')::date;
    v_owner text;
    v_claimed boolean;
begin
    if p_user_id is null or length(trim(p_user_id)) = 0 then
        return json_build_object('ok', false, 'error', 'invalid_user');
    end if;

    -- IDOR 방지
    select user_id into v_owner
    from public.pill_items
    where id = p_pill_id and deleted_at is null;

    if v_owner is null then
        return json_build_object('ok', false, 'error', 'not_found');
    end if;
    if v_owner <> p_user_id then
        return json_build_object('ok', false, 'error', 'forbidden');
    end if;

    -- 오늘 이미 완료 처리됐으면 변경 잠금
    select exists(
        select 1 from public.mission_completions
        where user_id = p_user_id
          and mission_type = 'DAILY_PILL_TAKEN'
          and completed_date = v_today
    ) into v_claimed;

    if v_claimed then
        return json_build_object('ok', false, 'error', 'already_claimed');
    end if;

    if p_taken then
        insert into public.pill_daily_logs (user_id, pill_id, log_date)
        values (p_user_id, p_pill_id, v_today)
        on conflict (pill_id, log_date) do nothing;
    else
        delete from public.pill_daily_logs
         where pill_id = p_pill_id and log_date = v_today;
    end if;

    return json_build_object('ok', true);
end;
$$;

-- ============================================================
-- 9) RPC: 오늘 완료 처리
--    서버에서 전부 챙겼는지 검증 (클라이언트 조작 불가)
--    하루 1회만 (mission_completions unique 제약)
-- ============================================================
create or replace function public.claim_pill_complete(p_user_id text)
returns json
language plpgsql security definer set search_path = public
as $$
declare
    v_today date := (now() at time zone 'Asia/Seoul')::date;
    v_total integer;
    v_taken integer;
    v_count integer := 0;
begin
    if p_user_id is null or length(trim(p_user_id)) = 0 then
        return json_build_object('ok', false, 'error', 'invalid_user');
    end if;

    -- 전체 활성 영양제 수
    select count(*) into v_total
    from public.pill_items
    where user_id = p_user_id and deleted_at is null;

    if v_total = 0 then
        return json_build_object('ok', false, 'error', 'no_pills');
    end if;

    -- 오늘 챙긴 수
    select count(*) into v_taken
    from public.pill_daily_logs
    where user_id = p_user_id and log_date = v_today;

    if v_taken < v_total then
        return json_build_object(
            'ok', false, 'error', 'not_complete',
            'taken', v_taken, 'total', v_total
        );
    end if;

    -- 완료 기록 INSERT (캐시 제거됐으므로 reward_coins=0)
    insert into public.mission_completions (
        user_id, mission_type, completed_date, reward_coins
    )
    values (p_user_id, 'DAILY_PILL_TAKEN', v_today, 0)
    on conflict (user_id, mission_type, completed_date) do nothing;
    get diagnostics v_count = row_count;

    if v_count = 0 then
        return json_build_object('ok', false, 'error', 'duplicate');
    end if;

    return json_build_object('ok', true);
end;
$$;

-- ============================================================
-- 10) anon 키로는 RPC 함수만 호출 가능 (테이블 직접 접근은 RLS+revoke로 차단)
-- ============================================================
grant execute on function public.get_pill_status(text) to anon;
grant execute on function public.add_pill_item(text, text, text, text) to anon;
grant execute on function public.delete_pill_item(text, uuid) to anon;
grant execute on function public.set_pill_taken(text, uuid, boolean) to anon;
grant execute on function public.claim_pill_complete(text) to anon;
