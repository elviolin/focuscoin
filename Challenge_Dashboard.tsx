// Challenge_Dashboard.tsx — 외부 일일미션 대시보드 (Framer 컴포넌트)
// 영단어/영양제/물 3개 일일미션의 참여/학습/통과/적립 현황
// 데이터: focuscoin Supabase (challenge_progress, challenge_config)

import React, { useState, useEffect } from "react"
import { addPropertyControls, ControlType } from "framer"

/* ═══════════════════════════════════════════
   Supabase (focuscoin 프로젝트)
   ═══════════════════════════════════════════ */
const SUPA_URL_DEFAULT = "https://gjnwriqewsrwpbtxqbea.supabase.co"
const SUPA_KEY_DEFAULT =
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdqbndyaXFld3Nyd3BidHhxYmVhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzOTIzNjQsImV4cCI6MjA5NTk2ODM2NH0.XAcRHkdHh8WmwhJgYht__CPmopQadvWVR3h7c8uFswU"

const MISSIONS = [
    {
        id: "english",
        label: "영단어 챌린지",
        emoji: "📖",
        linked: true,
    },
    {
        id: "pill",
        label: "영양제 챙기기",
        emoji: "💊",
        linked: false,
    },
    {
        id: "water",
        label: "물 마시기",
        emoji: "💧",
        linked: false,
    },
]

function kstToday(): string {
    return new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10)
}
function kstDaysAgo(n: number): string {
    return new Date(Date.now() + 9 * 3600 * 1000 - n * 86400000)
        .toISOString()
        .slice(0, 10)
}

/* ═══════════════════════════════════════════
   Design Tokens (FocusCoin 공통)
   ═══════════════════════════════════════════ */
const T = {
    rBtn: 12,
    rCard: 18,
    rPill: 100,
    tXs: 11,
    tSm: 13,
    tMd: 14,
    tLg: 16,
    tXl: 20,
    t2xl: 26,
    t3xl: 38,
    wBody: 500,
    wLabel: 600,
    wBold: 700,
    cBg: "#FFFFFF",
    cPage: "#F6F7F9",
    cCard: "#F8F9FA",
    cDivider: "#F2F4F6",
    cBorder: "#EEF0F3",
    cText: "#191F28",
    cText2: "#6B7684",
    cText3: "#9AA0A8",
    cText4: "#C9CDD2",
    cGreen: "#06C167",
    cGreenBg: "#E6F9F0",
    cGreenDk: "#048A4A",
    cInfo: "#3DADFF",
    cInfoBg: "#EAF5FF",
    cWarn: "#F59E0B",
    cWarnBg: "#FEF3E2",
    shadow: "0 1px 3px rgba(25,31,40,0.05), 0 1px 2px rgba(25,31,40,0.03)",
}

const FONT =
    "'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
const MONO = "'JetBrains Mono', 'Consolas', monospace"

const CSS = `
  .fcd *, .fcd *::before, .fcd *::after { box-sizing: border-box; }
  .fcd-body::-webkit-scrollbar { width: 8px; }
  .fcd-body::-webkit-scrollbar-thumb { background: ${T.cBorder}; border-radius: 100px; }
  .fcd-btn { transition: all 0.15s ease; }
  .fcd-btn:active { transform: scale(0.96); }
  .fcd-btn:hover { opacity: 0.85; }
  .fcd-loading-dot { animation: fcd-dot-pulse 1.2s ease-in-out infinite; }
  .fcd-loading-dot:nth-child(2) { animation-delay: 0.15s; }
  .fcd-loading-dot:nth-child(3) { animation-delay: 0.3s; }
  @keyframes fcd-dot-pulse {
    0%, 100% { opacity: 0.25; transform: scale(0.9); }
    50% { opacity: 1; transform: scale(1.1); }
  }
`

/* ═══════════════════════════════════════════
   Types
   ═══════════════════════════════════════════ */
type ProgressRow = {
    id: string
    device_id: string
    challenge_id: string
    date: string
    studied_word_ids: string[]
    test_score: number | null
    test_passed: boolean
    claimed: boolean
    claimed_at: string | null
    created_at: string
    updated_at: string
}
type Period = "today" | "7d" | "30d" | "all"

/* ═══════════════════════════════════════════
   Main Component
   ═══════════════════════════════════════════ */
export default function MissionDashboard({
    supabaseUrl = SUPA_URL_DEFAULT,
    supabaseKey = SUPA_KEY_DEFAULT,
}: {
    supabaseUrl?: string
    supabaseKey?: string
}) {
    const [rows, setRows] = useState<ProgressRow[]>([])
    const [configs, setConfigs] = useState<any[]>([])
    const [loading, setLoading] = useState(true)
    const [loadError, setLoadError] = useState(false)
    const [period, setPeriod] = useState<Period>("7d")
    const [refreshKey, setRefreshKey] = useState(0)

    /* ─── 데이터 로드 ─── */
    useEffect(() => {
        let cancelled = false
        setLoading(true)
        setLoadError(false)

        const h = {
            apikey: supabaseKey,
            Authorization: `Bearer ${supabaseKey}`,
            "Content-Type": "application/json",
        }

        let dateFilter = ""
        if (period === "today") dateFilter = `&date=gte.${kstToday()}`
        else if (period === "7d") dateFilter = `&date=gte.${kstDaysAgo(6)}`
        else if (period === "30d") dateFilter = `&date=gte.${kstDaysAgo(29)}`

        Promise.all([
            fetch(
                `${supabaseUrl}/rest/v1/focuscoin_challenge_progress?select=*${dateFilter}&order=updated_at.desc&limit=500`,
                { headers: h }
            ).then((r) => r.json()),
            fetch(
                `${supabaseUrl}/rest/v1/focuscoin_challenge_config?select=*`,
                { headers: h }
            ).then((r) => r.json()),
        ])
            .then(([progRows, cfgRows]) => {
                if (cancelled) return
                if (Array.isArray(progRows)) setRows(progRows)
                if (Array.isArray(cfgRows)) setConfigs(cfgRows)
                setLoading(false)
            })
            .catch(() => {
                if (cancelled) return
                setLoadError(true)
                setLoading(false)
            })

        return () => {
            cancelled = true
        }
    }, [supabaseUrl, supabaseKey, period, refreshKey])

    /* ─── 집계 ─── */
    const agg = (list: ProgressRow[]) => {
        const participated = list.length
        const uniqueUsers = new Set(list.map((r) => r.device_id)).size
        const studiedDone = list.filter(
            (r) => (r.studied_word_ids || []).length >= 5
        ).length
        const passed = list.filter((r) => r.test_passed).length
        const claimed = list.filter((r) => r.claimed).length
        return { participated, uniqueUsers, studiedDone, passed, claimed }
    }
    const total = agg(rows)
    const passRate =
        total.participated > 0
            ? Math.round((total.passed / total.participated) * 100)
            : 0
    const claimRate =
        total.participated > 0
            ? Math.round((total.claimed / total.participated) * 100)
            : 0

    const isUserId = (id: string) =>
        !id.startsWith("fc_") && id !== "fc_anonymous"

    const fmtTime = (iso: string) => {
        if (!iso) return "-"
        const d = new Date(new Date(iso).getTime() + 9 * 3600 * 1000)
        return d.toISOString().slice(5, 16).replace("T", " ")
    }

    const periodLabel: Record<Period, string> = {
        today: "오늘",
        "7d": "7일",
        "30d": "30일",
        all: "전체",
    }

    /* ─── Loading (모든 훅 뒤, 메인 return 직전) ─── */
    if (loading) {
        return (
            <div className="fcd" style={rootSty}>
                <style>{CSS}</style>
                <div style={loadingWrapSty}>
                    <span className="fcd-loading-dot" style={loadingDotSty} />
                    <span className="fcd-loading-dot" style={loadingDotSty} />
                    <span className="fcd-loading-dot" style={loadingDotSty} />
                </div>
            </div>
        )
    }

    return (
        <div className="fcd" style={rootSty}>
            <style>{CSS}</style>

            <div className="fcd-body" style={bodySty}>
                {/* ─── Header ─── */}
                <div style={hdrSty}>
                    <h1 style={hdrTitleSty}>외부 일일미션 대시보드</h1>
                    <div style={hdrRightSty}>
                        <div style={periodGroupSty}>
                            {(["today", "7d", "30d", "all"] as Period[]).map(
                                (p) => (
                                    <button
                                        key={p}
                                        className="fcd-btn"
                                        onClick={() => setPeriod(p)}
                                        style={{
                                            ...periodBtnSty,
                                            background:
                                                period === p
                                                    ? T.cText
                                                    : "transparent",
                                            color:
                                                period === p
                                                    ? "#FFF"
                                                    : T.cText2,
                                        }}
                                    >
                                        {periodLabel[p]}
                                    </button>
                                )
                            )}
                        </div>
                        <button
                            className="fcd-btn"
                            style={refreshBtnSty}
                            onClick={() => setRefreshKey((k) => k + 1)}
                        >
                            새로고침
                        </button>
                    </div>
                </div>

                {loadError && (
                    <div style={errorBannerSty}>
                        데이터를 불러오지 못했습니다. 새로고침을 눌러 주세요.
                    </div>
                )}

                {/* ─── KPI Cards ─── */}
                <div style={kpiGridSty}>
                    <div style={kpiCardSty}>
                        <div style={kpiIconSty}>🙋</div>
                        <div style={kpiLabelSty}>참여</div>
                        <div style={kpiNumSty}>{total.participated}</div>
                        <div style={kpiMetaSty}>
                            고유 사용자 {total.uniqueUsers}명
                        </div>
                    </div>
                    <div style={kpiCardSty}>
                        <div style={kpiIconSty}>✏️</div>
                        <div style={kpiLabelSty}>학습 완료</div>
                        <div style={kpiNumSty}>{total.studiedDone}</div>
                        <div style={kpiMetaSty}>5개 단어 모두 학습</div>
                    </div>
                    <div style={kpiCardSty}>
                        <div style={kpiIconSty}>🎯</div>
                        <div style={kpiLabelSty}>테스트 통과</div>
                        <div style={{ ...kpiNumSty, color: T.cGreen }}>
                            {total.passed}
                        </div>
                        <div style={kpiMetaSty}>통과율 {passRate}%</div>
                    </div>
                    <div style={kpiCardSty}>
                        <div style={kpiIconSty}>💰</div>
                        <div style={kpiLabelSty}>보상 적립</div>
                        <div style={{ ...kpiNumSty, color: T.cGreenDk }}>
                            {total.claimed}
                        </div>
                        <div style={kpiMetaSty}>적립율 {claimRate}%</div>
                    </div>
                </div>

                {/* ─── 퍼널 ─── */}
                <div style={cardSty}>
                    <div style={cardTitleSty}>완료 퍼널</div>
                    <div style={funnelWrapSty}>
                        {[
                            { label: "참여", n: total.participated, c: "#B0B8C1" },
                            { label: "학습 완료", n: total.studiedDone, c: T.cInfo },
                            { label: "테스트 통과", n: total.passed, c: T.cGreen },
                            { label: "보상 적립", n: total.claimed, c: T.cGreenDk },
                        ].map((s, i) => {
                            const max = Math.max(total.participated, 1)
                            const w = Math.max(
                                Math.round((s.n / max) * 100),
                                s.n > 0 ? 8 : 2
                            )
                            return (
                                <div key={i} style={funnelRowSty}>
                                    <span style={funnelLabelSty}>{s.label}</span>
                                    <div style={funnelBarBgSty}>
                                        <div
                                            style={{
                                                ...funnelBarSty,
                                                width: `${w}%`,
                                                background: s.c,
                                            }}
                                        />
                                    </div>
                                    <span style={funnelNumSty}>{s.n}</span>
                                </div>
                            )
                        })}
                    </div>
                </div>

                {/* ─── 미션별 현황 ─── */}
                <div style={missionGridSty}>
                    {MISSIONS.map((m) => {
                        const list = rows.filter(
                            (r) => r.challenge_id === m.id
                        )
                        const a = agg(list)
                        const cfg = configs.find((c) => c.id === m.id)
                        return (
                            <div key={m.id} style={cardSty}>
                                <div style={missionHeadSty}>
                                    <span style={missionTitleSty}>
                                        {m.emoji} {m.label}
                                    </span>
                                    <span
                                        style={
                                            m.linked
                                                ? linkedBadgeSty
                                                : unlinkedBadgeSty
                                        }
                                    >
                                        {m.linked ? "운영 중" : "준비 중"}
                                    </span>
                                </div>
                                <div style={missionStatsSty}>
                                    <div style={missionStatSty}>
                                        <div style={missionStatNumSty}>
                                            {a.participated}
                                        </div>
                                        <div style={missionStatLabelSty}>
                                            참여
                                        </div>
                                    </div>
                                    <div style={missionStatSty}>
                                        <div style={missionStatNumSty}>
                                            {a.passed}
                                        </div>
                                        <div style={missionStatLabelSty}>
                                            통과
                                        </div>
                                    </div>
                                    <div style={missionStatSty}>
                                        <div
                                            style={{
                                                ...missionStatNumSty,
                                                color: T.cGreenDk,
                                            }}
                                        >
                                            {a.claimed}
                                        </div>
                                        <div style={missionStatLabelSty}>
                                            적립
                                        </div>
                                    </div>
                                    <div style={missionStatSty}>
                                        <div style={missionStatNumSty}>
                                            +{cfg ? cfg.reward_cash : 10}
                                        </div>
                                        <div style={missionStatLabelSty}>
                                            캐시
                                        </div>
                                    </div>
                                </div>
                            </div>
                        )
                    })}
                </div>

                {/* ─── 최근 활동 ─── */}
                <div style={cardSty}>
                    <div style={cardTitleRowSty}>
                        <span style={cardTitleSty}>최근 활동</span>
                        <span style={cardMetaSty}>
                            최근 {Math.min(rows.length, 15)}건 (총 {rows.length}
                            건)
                        </span>
                    </div>
                    {rows.length === 0 ? (
                        <div style={emptySty}>
                            선택한 기간에 활동 기록이 없습니다
                        </div>
                    ) : (
                        <table style={tableSty}>
                            <thead>
                                <tr>
                                    <th style={thSty}>일시 (KST)</th>
                                    <th style={thSty}>사용자</th>
                                    <th style={thSty}>미션</th>
                                    <th style={thSty}>학습</th>
                                    <th style={thSty}>점수</th>
                                    <th style={thSty}>통과</th>
                                    <th style={thSty}>적립</th>
                                </tr>
                            </thead>
                            <tbody>
                                {rows.slice(0, 15).map((r) => {
                                    const mission = MISSIONS.find(
                                        (m) => m.id === r.challenge_id
                                    )
                                    return (
                                        <tr key={r.id}>
                                            <td style={tdSty}>
                                                {fmtTime(
                                                    r.updated_at ||
                                                        r.created_at
                                                )}
                                            </td>
                                            <td style={tdSty}>
                                                <span
                                                    style={
                                                        isUserId(r.device_id)
                                                            ? userBadgeSty
                                                            : deviceBadgeSty
                                                    }
                                                >
                                                    {isUserId(r.device_id)
                                                        ? "회원"
                                                        : "기기"}
                                                </span>{" "}
                                                <span style={tdMonoSty}>
                                                    {r.device_id.length > 18
                                                        ? r.device_id.slice(
                                                              0,
                                                              18
                                                          ) + "…"
                                                        : r.device_id}
                                                </span>
                                            </td>
                                            <td style={tdSty}>
                                                {mission
                                                    ? mission.emoji +
                                                      " " +
                                                      mission.label
                                                    : r.challenge_id}
                                            </td>
                                            <td style={tdSty}>
                                                {
                                                    (
                                                        r.studied_word_ids ||
                                                        []
                                                    ).length
                                                }
                                                개
                                            </td>
                                            <td style={tdSty}>
                                                {r.test_score === null ||
                                                r.test_score === undefined
                                                    ? "-"
                                                    : r.test_score + "점"}
                                            </td>
                                            <td style={tdSty}>
                                                <span
                                                    style={
                                                        r.test_passed
                                                            ? okBadgeSty
                                                            : grayBadgeSty
                                                    }
                                                >
                                                    {r.test_passed
                                                        ? "통과"
                                                        : "-"}
                                                </span>
                                            </td>
                                            <td style={tdSty}>
                                                <span
                                                    style={
                                                        r.claimed
                                                            ? okBadgeSty
                                                            : grayBadgeSty
                                                    }
                                                >
                                                    {r.claimed ? "적립" : "-"}
                                                </span>
                                            </td>
                                        </tr>
                                    )
                                })}
                            </tbody>
                        </table>
                    )}
                </div>
            </div>
        </div>
    )
}

/* ═══════════════════════════════════════════
   Styles
   ═══════════════════════════════════════════ */
const rootSty: React.CSSProperties = {
    width: "100%",
    height: "100%",
    background: T.cPage,
    color: T.cText,
    fontFamily: FONT,
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
    WebkitFontSmoothing: "antialiased",
}
const loadingWrapSty: React.CSSProperties = {
    flex: 1,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
}
const loadingDotSty: React.CSSProperties = {
    width: 8,
    height: 8,
    borderRadius: "50%",
    background: T.cGreen,
    display: "inline-block",
}
const bodySty: React.CSSProperties = {
    flex: 1,
    overflowY: "auto",
    padding: "32px 40px 56px",
    display: "flex",
    flexDirection: "column",
    gap: 14,
    maxWidth: 1160,
    margin: "0 auto",
    width: "100%",
}
const hdrSty: React.CSSProperties = {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 16,
    flexWrap: "wrap",
    marginBottom: 6,
}
const hdrTitleSty: React.CSSProperties = {
    margin: 0,
    fontSize: T.t2xl,
    fontWeight: T.wBold,
    letterSpacing: -0.6,
}
const hdrRightSty: React.CSSProperties = {
    display: "flex",
    gap: 8,
    alignItems: "center",
}
const periodGroupSty: React.CSSProperties = {
    display: "flex",
    background: T.cBg,
    borderRadius: T.rBtn,
    padding: 3,
    gap: 2,
    boxShadow: T.shadow,
}
const periodBtnSty: React.CSSProperties = {
    border: "none",
    borderRadius: 9,
    padding: "7px 14px",
    fontSize: T.tSm,
    fontWeight: T.wLabel,
    fontFamily: FONT,
    cursor: "pointer",
}
const refreshBtnSty: React.CSSProperties = {
    border: "none",
    background: T.cBg,
    borderRadius: T.rBtn,
    padding: "10px 16px",
    fontSize: T.tSm,
    fontWeight: T.wLabel,
    fontFamily: FONT,
    color: T.cText,
    cursor: "pointer",
    boxShadow: T.shadow,
}
const errorBannerSty: React.CSSProperties = {
    background: "#FEF2F2",
    color: "#B91C1C",
    borderRadius: T.rBtn,
    padding: "12px 16px",
    fontSize: T.tSm,
    fontWeight: T.wLabel,
}
const kpiGridSty: React.CSSProperties = {
    display: "grid",
    gridTemplateColumns: "repeat(4, 1fr)",
    gap: 14,
}
const kpiCardSty: React.CSSProperties = {
    background: T.cBg,
    borderRadius: T.rCard,
    padding: "20px 22px",
    boxShadow: T.shadow,
}
const kpiIconSty: React.CSSProperties = {
    fontSize: 22,
    marginBottom: 12,
    lineHeight: 1,
}
const kpiLabelSty: React.CSSProperties = {
    fontSize: T.tSm,
    color: T.cText2,
    fontWeight: T.wLabel,
    marginBottom: 8,
}
const kpiNumSty: React.CSSProperties = {
    fontSize: T.t3xl,
    fontWeight: T.wBold,
    letterSpacing: -1.2,
    lineHeight: 1,
    color: T.cText,
}
const kpiMetaSty: React.CSSProperties = {
    fontSize: T.tXs,
    color: T.cText3,
    marginTop: 9,
    fontWeight: T.wBody,
}
const cardSty: React.CSSProperties = {
    background: T.cBg,
    borderRadius: T.rCard,
    padding: "22px 24px",
    boxShadow: T.shadow,
}
const cardTitleSty: React.CSSProperties = {
    fontSize: T.tLg,
    fontWeight: T.wBold,
    letterSpacing: -0.3,
}
const cardTitleRowSty: React.CSSProperties = {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "baseline",
    marginBottom: 14,
}
const cardMetaSty: React.CSSProperties = {
    fontSize: T.tXs,
    color: T.cText3,
    fontWeight: T.wBody,
}
const funnelWrapSty: React.CSSProperties = {
    display: "flex",
    flexDirection: "column",
    gap: 12,
    marginTop: 16,
}
const funnelRowSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: 14,
}
const funnelLabelSty: React.CSSProperties = {
    width: 88,
    fontSize: T.tSm,
    color: T.cText2,
    fontWeight: T.wLabel,
    flexShrink: 0,
}
const funnelBarBgSty: React.CSSProperties = {
    flex: 1,
    height: 26,
    background: T.cDivider,
    borderRadius: 8,
    overflow: "hidden",
}
const funnelBarSty: React.CSSProperties = {
    height: "100%",
    borderRadius: 8,
    transition: "width 0.5s ease",
}
const funnelNumSty: React.CSSProperties = {
    width: 42,
    textAlign: "right",
    fontSize: T.tMd,
    fontWeight: T.wBold,
    flexShrink: 0,
}
const missionGridSty: React.CSSProperties = {
    display: "grid",
    gridTemplateColumns: "repeat(3, 1fr)",
    gap: 14,
}
const missionHeadSty: React.CSSProperties = {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: 16,
}
const missionTitleSty: React.CSSProperties = {
    fontSize: T.tLg,
    fontWeight: T.wBold,
    letterSpacing: -0.3,
}
const linkedBadgeSty: React.CSSProperties = {
    fontSize: T.tXs,
    fontWeight: T.wBold,
    padding: "4px 10px",
    borderRadius: T.rPill,
    background: T.cGreenBg,
    color: T.cGreenDk,
}
const unlinkedBadgeSty: React.CSSProperties = {
    fontSize: T.tXs,
    fontWeight: T.wBold,
    padding: "4px 10px",
    borderRadius: T.rPill,
    background: T.cWarnBg,
    color: T.cWarn,
}
const missionStatsSty: React.CSSProperties = {
    display: "grid",
    gridTemplateColumns: "repeat(4, 1fr)",
    gap: 6,
    background: T.cCard,
    borderRadius: T.rBtn,
    padding: "14px 8px",
}
const missionStatSty: React.CSSProperties = { textAlign: "center" }
const missionStatNumSty: React.CSSProperties = {
    fontSize: T.tXl,
    fontWeight: T.wBold,
    letterSpacing: -0.5,
}
const missionStatLabelSty: React.CSSProperties = {
    fontSize: T.tXs,
    color: T.cText3,
    marginTop: 4,
    fontWeight: T.wBody,
}
const emptySty: React.CSSProperties = {
    padding: "36px 0",
    textAlign: "center",
    fontSize: T.tSm,
    color: T.cText3,
    fontWeight: T.wBody,
}
const tableSty: React.CSSProperties = {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: T.tSm,
}
const thSty: React.CSSProperties = {
    textAlign: "left",
    padding: "8px 10px",
    fontSize: T.tXs,
    color: T.cText3,
    fontWeight: T.wLabel,
    borderBottom: `1px solid ${T.cDivider}`,
    whiteSpace: "nowrap",
}
const tdSty: React.CSSProperties = {
    padding: "11px 10px",
    borderBottom: `1px solid ${T.cDivider}`,
    color: T.cText,
    fontWeight: T.wBody,
    whiteSpace: "nowrap",
}
const tdMonoSty: React.CSSProperties = {
    fontFamily: MONO,
    fontSize: T.tXs,
    color: T.cText2,
}
const userBadgeSty: React.CSSProperties = {
    fontSize: T.tXs,
    fontWeight: T.wBold,
    padding: "2px 8px",
    borderRadius: T.rPill,
    background: T.cInfoBg,
    color: T.cInfo,
}
const deviceBadgeSty: React.CSSProperties = {
    fontSize: T.tXs,
    fontWeight: T.wBold,
    padding: "2px 8px",
    borderRadius: T.rPill,
    background: T.cDivider,
    color: T.cText3,
}
const okBadgeSty: React.CSSProperties = {
    fontSize: T.tXs,
    fontWeight: T.wBold,
    padding: "3px 10px",
    borderRadius: T.rPill,
    background: T.cGreenBg,
    color: T.cGreenDk,
}
const grayBadgeSty: React.CSSProperties = {
    fontSize: T.tXs,
    fontWeight: T.wBold,
    padding: "3px 10px",
    borderRadius: T.rPill,
    background: T.cDivider,
    color: T.cText4,
}

/* ═══════════════════════════════════════════
   Property Controls
   ═══════════════════════════════════════════ */
addPropertyControls(MissionDashboard, {
    supabaseUrl: {
        type: ControlType.String,
        title: "Supabase URL",
        defaultValue: SUPA_URL_DEFAULT,
    },
    supabaseKey: {
        type: ControlType.String,
        title: "Supabase Key",
        defaultValue: SUPA_KEY_DEFAULT,
    },
})
