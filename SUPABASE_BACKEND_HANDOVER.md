# 플래닛 일일 챌린지 — Supabase 백엔드 명세 (v2)

물 마시기 챌린지의 백엔드 명세. 구 `planit-study.com` API는 **완전 제거**되었고, Supabase로 대체되었습니다.

> **적용일**: 2026-06-07 / **상태**: 운영 중 (라이브 E2E 검증 완료)
> 구 버전 서버 명세(`SERVER_DEV_HANDOVER`, planit API 기준)는 폐기. 이 문서가 단일 기준.

---

## 1. 인프라

| 항목 | 값 |
|---|---|
| Supabase 프로젝트 | `gjnwriqewsrwpbtxqbea` (org: aidzet / focuscoin) |
| Base URL | `https://gjnwriqewsrwpbtxqbea.supabase.co` |
| 인증 | anon key (클라이언트 노출용, 컴포넌트 코드에 포함) |
| 스키마 SQL | `focuscoin_water_schema.sql` (이 저장소) |

---

## 2. 테이블

### `water_logs` — 유저 x 날짜당 1행

| 컬럼 | 타입 | 설명 |
|---|---|---|
| `id` | uuid PK | |
| `user_id` | text | WebView URL의 `userId` |
| `log_date` | date | KST 기준 날짜. `unique(user_id, log_date)` |
| `cups` | int | 0~8 |
| `claimed` | bool | 오늘 보상 수령 여부 |
| `last_added_at` | timestamptz | 쿨다운 계산 기준 |

### `mission_completions` — 보상 적립 원장 (3개 미션 공용 설계)

| 컬럼 | 타입 | 설명 |
|---|---|---|
| `id` | uuid PK | |
| `user_id` | text | |
| `mission_type` | text | `DAILY_WATER_CUPS` (영단어/영양제 확장 시 재사용) |
| `completed_date` | date | KST. `unique(user_id, mission_type, completed_date)` → **1일 1회 강제** |
| `reward_coins` | int | 현재 10 |

> 부모 앱 캐시 잔액 반영은 이 테이블을 기준으로 연동 (부모 서버가 조회하거나, 추후 webhook/Edge Function 추가 가능).

---

## 3. 보안 모델

- 두 테이블 모두 **RLS 활성화 + anon 권한 전면 revoke** → anon 키로 테이블 직접 읽기/쓰기 **불가능** (실측: `permission denied`)
- 모든 조작은 아래 **RPC 함수 3개**(`security definer`)로만 가능
- 쿨다운, 잔 수 상한, 1일 1회 보상이 전부 **서버에서 강제** — 클라이언트 조작/새로고침으로 우회 불가

---

## 4. RPC 함수 (REST: `POST /rest/v1/rpc/{함수명}`)

공통 헤더: `apikey: {anon}`, `Authorization: Bearer {anon}`, `Content-Type: application/json`

### 4-1. `get_water_status` — 오늘 상태 조회

```json
요청: {"p_user_id": "abc123"}
응답: {"ok": true, "cups": 4, "claimed": false, "cooldown_remain_sec": 32}
```

### 4-2. `add_water_cups` — 물 등록 (쿨다운 1분 서버 강제)

```json
요청: {"p_user_id": "abc123", "p_cups": 2}   // p_cups: 1~4만 허용
성공: {"ok": true, "cups": 6, "claimed": false, "cooldown_remain_sec": 60}
쿨다운: {"ok": false, "error": "cooldown", "cooldown_remain_sec": 41, "cups": 6, "claimed": false}
기타 에러: "invalid_user" | "invalid_cups" | "already_full" | "already_claimed"
```

### 4-3. `claim_water_reward` — 보상 수령 (8잔 완료 + 1일 1회)

```json
요청: {"p_user_id": "abc123"}
성공: {"ok": true, "reward_coins": 10}     // mission_completions에 INSERT + claimed=true
중복: {"ok": false, "error": "duplicate"}
미달: {"ok": false, "error": "not_full"}
```

---

## 5. 정책 요약

| 정책 | 값 | 강제 위치 |
|---|---|---|
| 하루 목표 | 8잔 | 서버 (cups 상한) |
| 1회 등록 최대 | 4잔 | 서버 (`invalid_cups`) |
| 등록 간 쿨다운 | **1분** | 서버 (`last_added_at` 비교) |
| 보상 | +10 코인, 1일 1회 | 서버 (unique 제약) |
| 날짜 리셋 | KST 00:00 | 서버 (`Asia/Seoul` 기준 date) |

> 쿨다운 변경 이력: 5분 → 3분 → 1분 (2026-06-07). 변경 시 ① 함수 2개 SQL 교체(`get_water_status`, `add_water_cups`의 interval/60초 값) + ② `Challenge_water.tsx` 토스트 문구/`COOLDOWN_MS` 동시 수정 필요.

---

## 6. 검증 완료 시나리오 (2026-06-07 실측)

| # | 시나리오 | 결과 |
|---|---|---|
| 1 | 신규 유저 상태 조회 | ✅ 0잔 |
| 2 | 물 등록 → DB 저장 → 새로고침 유지 | ✅ |
| 3 | 1분 내 재등록 | ✅ 차단 + 토스트 |
| 4 | 1회 5잔 이상 등록 시도 | ✅ `invalid_cups` |
| 5 | 8잔 미만 보상 시도 | ✅ `not_full` |
| 6 | 8잔 → 광고 → 보상 적립 | ✅ 🎉 팝업 + `mission_completions` 기록 |
| 7 | 같은 날 보상 재시도 | ✅ `duplicate` |
| 8 | 보상 후 새로고침 | ✅ "오늘 챌린지 완료 ✓" 유지 |
| 9 | 유저 A/B 동시 사용 | ✅ 잔 수, 쿨다운, 보상 전부 독립 |
| 10 | anon 키로 테이블 직접 접근 | ✅ 차단 (`permission denied`) |

---

## 7. 추후 확장

- **영단어/영양제 연동**: `mission_completions`에 `DAILY_ENGLISH_QUIZ` / `DAILY_VITAMIN_CHECK` 타입으로 동일 패턴 적용 (영양제는 등록 목록용 테이블 추가 필요)
- **부모 앱 캐시 잔액 연동**: `mission_completions` 조회 API 또는 Supabase Edge Function → 부모 서버 webhook
- **streak (연속 일수)**: `mission_completions`에서 날짜 연속성 집계로 구현 가능
- **AdMob SSV**: 광고 시청 토큰 검증을 `claim_water_reward`에 파라미터로 추가
- **Rate Limit**: 비정상 다발 호출 대비 (현재는 쿨다운이 사실상 rate limit 역할)
