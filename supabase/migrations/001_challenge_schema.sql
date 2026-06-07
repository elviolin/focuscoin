-- ═══════════════════════════════════════════════════════════
-- FocusCoin 일일미션 스키마 v1
-- 챌린지 설정 / 단어 풀 / 기기별 진행 기록 + streak 계산
-- ═══════════════════════════════════════════════════════════

-- 1. 챌린지 설정 (캐시 금액, 통과 기준 등 — DB 값만 바꾸면 화면 전체 반영)
create table if not exists focuscoin_challenge_config (
    id text primary key,                       -- 'english' | 'pill' | 'water'
    title text not null,
    reward_cash int not null default 10,       -- +10캐시
    pass_threshold int not null default 3,     -- 5문제 중 3문제 통과
    questions_per_test int not null default 5,
    words_per_day int not null default 5,
    ad_seconds int not null default 5,
    active boolean not null default true,
    updated_at timestamptz not null default now()
);

-- 2. 영단어 풀
create table if not exists focuscoin_words (
    id uuid primary key default gen_random_uuid(),
    en text not null,
    kr text not null,
    pos text not null default '',
    ex text not null default '',
    ex_kr text not null default '',
    active boolean not null default true,
    created_at timestamptz not null default now(),
    unique (en)
);

-- 3. 기기별 일일 진행 기록
create table if not exists focuscoin_challenge_progress (
    id uuid primary key default gen_random_uuid(),
    device_id text not null,                   -- 브라우저 localStorage 자동 발급 ID
    challenge_id text not null default 'english',
    date date not null default (now() at time zone 'Asia/Seoul')::date,
    studied_word_ids uuid[] not null default '{}',
    test_score int,
    test_passed boolean not null default false,
    claimed boolean not null default false,
    claimed_at timestamptz,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    unique (device_id, challenge_id, date)
);

create index if not exists idx_fc_progress_device
    on focuscoin_challenge_progress (device_id, challenge_id, date desc);

-- 4. 오늘의 단어 5개 (날짜 기준 결정적 선택 — 같은 날은 모두 같은 단어)
create or replace function focuscoin_get_daily_words(p_date date default (now() at time zone 'Asia/Seoul')::date)
returns setof focuscoin_words
language sql stable
as $$
    select w.*
    from focuscoin_words w
    where w.active
    order by md5(w.id::text || p_date::text)
    limit (select words_per_day from focuscoin_challenge_config where id = 'english');
$$;

-- 5. 연속 학습일(streak) 계산 — 오늘 또는 어제부터 거꾸로 연속 claimed 일수
create or replace function focuscoin_get_streak(p_device_id text, p_challenge_id text default 'english')
returns int
language plpgsql stable
as $$
declare
    v_streak int := 0;
    v_check date := (now() at time zone 'Asia/Seoul')::date;
    v_found boolean;
begin
    -- 오늘 아직 안 했으면 어제부터 계산
    select exists(
        select 1 from focuscoin_challenge_progress
        where device_id = p_device_id and challenge_id = p_challenge_id
          and date = v_check and claimed
    ) into v_found;
    if not v_found then
        v_check := v_check - 1;
    end if;

    loop
        select exists(
            select 1 from focuscoin_challenge_progress
            where device_id = p_device_id and challenge_id = p_challenge_id
              and date = v_check and claimed
        ) into v_found;
        exit when not v_found;
        v_streak := v_streak + 1;
        v_check := v_check - 1;
    end loop;

    return v_streak;
end;
$$;

-- 6. RLS — 프로토타입 단계: anon 읽기 허용 + progress는 본인 행 upsert 허용
alter table focuscoin_challenge_config enable row level security;
alter table focuscoin_words enable row level security;
alter table focuscoin_challenge_progress enable row level security;

create policy "config_public_read" on focuscoin_challenge_config
    for select to anon, authenticated using (true);

create policy "words_public_read" on focuscoin_words
    for select to anon, authenticated using (true);

create policy "progress_anon_select" on focuscoin_challenge_progress
    for select to anon, authenticated using (true);

create policy "progress_anon_insert" on focuscoin_challenge_progress
    for insert to anon, authenticated with check (true);

create policy "progress_anon_update" on focuscoin_challenge_progress
    for update to anon, authenticated using (true) with check (true);

-- ═══════════════════════════════════════════════════════════
-- 시드 데이터
-- ═══════════════════════════════════════════════════════════

insert into focuscoin_challenge_config (id, title, reward_cash, pass_threshold, questions_per_test, words_per_day, ad_seconds)
values ('english', '영단어 챌린지', 10, 3, 5, 5, 5)
on conflict (id) do nothing;

insert into focuscoin_words (en, kr, pos, ex, ex_kr) values
    ('diligent', '부지런한', 'adj.', 'She is a diligent student.', '그녀는 성실한 학생이다.'),
    ('accomplish', '성취하다', 'v.', 'You can accomplish anything.', '당신은 무엇이든 이룰 수 있다.'),
    ('gratitude', '감사', 'n.', 'Express your gratitude.', '감사를 표현하세요.'),
    ('perseverance', '인내', 'n.', 'Success requires perseverance.', '성공은 인내를 필요로 한다.'),
    ('consistent', '꾸준한', 'adj.', 'Be consistent in your habits.', '습관을 꾸준히 유지하세요.'),
    ('determine', '결심하다', 'v.', 'I determined to start early.', '나는 일찍 시작하기로 결심했다.'),
    ('improve', '개선하다', 'v.', 'Practice will improve your skills.', '연습이 실력을 향상시킨다.'),
    ('focus', '집중하다', 'v.', 'Focus on one thing at a time.', '한 번에 한 가지에 집중하세요.'),
    ('habit', '습관', 'n.', 'Good habits change your life.', '좋은 습관이 인생을 바꾼다.'),
    ('reward', '보상', 'n.', 'Effort brings its own reward.', '노력은 그 자체로 보상을 가져온다.'),
    ('patience', '인내심', 'n.', 'Patience is a virtue.', '인내는 미덕이다.'),
    ('achieve', '달성하다', 'v.', 'She achieved her goal.', '그녀는 목표를 달성했다.'),
    ('confident', '자신감 있는', 'adj.', 'Be confident in yourself.', '자신을 믿으세요.'),
    ('overcome', '극복하다', 'v.', 'You can overcome any challenge.', '어떤 어려움도 극복할 수 있다.'),
    ('progress', '발전', 'n.', 'Small steps make big progress.', '작은 걸음이 큰 발전을 만든다.')
on conflict (en) do nothing;
