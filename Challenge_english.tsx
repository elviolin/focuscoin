// EnglishWordChallenge.tsx — Framer 컴포넌트
// 매일 영단어 5개 학습 → 2지선다 테스트 → 3문제 이상 정답 → 짧은 광고 → +N캐시
// 외곽 폰 목업 chrome 없음, 부모 frame 사이즈 = 실제 앱 화면
// v2: Supabase 실데이터 연동 — 캐시 금액/통과 기준/광고 시간/단어/연속일 전부 DB 관리

import React, { useState, useRef, useEffect } from "react"
import { addPropertyControls, ControlType } from "framer"

/* ═══════════════════════════════════════════
   Supabase (focuscoin 프로젝트)
   ═══════════════════════════════════════════ */
const SUPA_URL_DEFAULT = "https://gjnwriqewsrwpbtxqbea.supabase.co"
const SUPA_KEY_DEFAULT =
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdqbndyaXFld3Nyd3BidHhxYmVhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzOTIzNjQsImV4cCI6MjA5NTk2ODM2NH0.XAcRHkdHh8WmwhJgYht__CPmopQadvWVR3h7c8uFswU"

function getDeviceId(): string {
    if (typeof window === "undefined") return "ssr"
    try {
        let id = window.localStorage.getItem("fc_device_id")
        if (!id) {
            id =
                "fc_" +
                Math.random().toString(36).slice(2, 10) +
                Date.now().toString(36)
            window.localStorage.setItem("fc_device_id", id)
        }
        return id
    } catch (e) {
        return "fc_anonymous"
    }
}

function kstToday(): string {
    return new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10)
}

/* ═══════════════════════════════════════════
   Design Tokens
   ═══════════════════════════════════════════ */
const T = {
    rBtn: 12,
    rCard: 16,
    rSheet: 28,
    rPill: 100,

    tXs: 11,
    tSm: 13,
    tMd: 14,
    tLg: 16,
    tXl: 20,
    t2xl: 28,
    t3xl: 40,

    wBody: 500,
    wLabel: 600,
    wBold: 700,

    cBg: "#FFFFFF",
    cCard: "#F8F9FA",
    cDivider: "#F2F4F6",
    cBorder: "#ECEEF0",
    cText: "#191F28",
    cText2: "#6B7684",
    cText3: "#9AA0A8",
    cText4: "#C9CDD2",
    cGreen: "#06C167",
    cGreenBg: "#E6F9F0",
    cGreenDk: "#048A4A",
    cRed: "#EF4444",
    cRedBg: "#FEF2F2",
    cRedDk: "#B91C1C",
}

const FONT =
    "'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"

/* ═══════════════════════════════════════════
   Embedded CSS (animations + pseudo states)
   ═══════════════════════════════════════════ */
const CSS = `
  .ewc *, .ewc *::before, .ewc *::after { box-sizing: border-box; }
  .ewc-body::-webkit-scrollbar { width: 0; }
  .ewc-body { scrollbar-width: none; -webkit-overflow-scrolling: touch; }

  .ewc-cta-chip { animation: ewc-cta-chip-pulse 1.8s ease-in-out infinite; }
  @keyframes ewc-cta-chip-pulse {
    0%, 100% { transform: scale(1); }
    50% { transform: scale(1.08); }
  }

  .ewc-fb { animation: ewc-fb-pop 0.4s cubic-bezier(0.34, 1.56, 0.64, 1); }
  @keyframes ewc-fb-pop {
    0% { transform: scale(0.3); opacity: 0; }
    60% { transform: scale(1.1); opacity: 1; }
    100% { transform: scale(1); }
  }

  .ewc-opt-correct { animation: ewc-opt-celebrate 0.5s ease-out; }
  @keyframes ewc-opt-celebrate {
    0% { transform: scale(1); }
    40% { transform: scale(1.03); }
    100% { transform: scale(1); }
  }

  .ewc-overlay { animation: ewc-fade-in 0.18s ease; }
  @keyframes ewc-fade-in { from { opacity: 0; } to { opacity: 1; } }

  .ewc-sheet-bottom { animation: ewc-slide-up 0.28s cubic-bezier(0.32, 0.72, 0, 1); }
  @keyframes ewc-slide-up {
    from { transform: translateY(100%); }
    to { transform: translateY(0); }
  }

  .ewc-sheet-center { animation: ewc-pop-in 0.22s cubic-bezier(0.32, 0.72, 0, 1); }
  @keyframes ewc-pop-in {
    from { opacity: 0; transform: scale(0.94); }
    to { opacity: 1; transform: scale(1); }
  }

  .ewc-cta:active:not(:disabled),
  .ewc-m-btn:active,
  .ewc-test-opt:active:not(:disabled) {
    transform: scale(0.985);
  }

  .ewc-word-row:active { background: ${T.cCard}; }
  .ewc-cta-green:not(:disabled):hover { background: ${T.cGreenDk}; }

  .ewc-loading-dot { animation: ewc-dot-pulse 1.2s ease-in-out infinite; }
  .ewc-loading-dot:nth-child(2) { animation-delay: 0.15s; }
  .ewc-loading-dot:nth-child(3) { animation-delay: 0.3s; }
  @keyframes ewc-dot-pulse {
    0%, 100% { opacity: 0.25; transform: scale(0.9); }
    50% { opacity: 1; transform: scale(1.1); }
  }
`

/* ═══════════════════════════════════════════
   Types
   ═══════════════════════════════════════════ */
type WordItem = {
    id: string
    en: string
    kr: string
    pos: string
    ex: string
    exKr: string
}
type Popup = null | "study" | "test" | "test_result" | "ad" | "reward"
type TestQuestion = {
    id: string
    en: string
    options: string[]
    correct: string
}
type ChallengeConfig = {
    rewardCash: number
    passThreshold: number
    adSeconds: number
}

/* ═══════════════════════════════════════════
   Fallback Word Pool (미리보기 모드 / DB 연결 실패 시)
   ═══════════════════════════════════════════ */
const FALLBACK_WORDS: WordItem[] = [
    {
        id: "w1",
        en: "diligent",
        kr: "부지런한",
        pos: "adj.",
        ex: "She is a diligent student.",
        exKr: "그녀는 성실한 학생이다.",
    },
    {
        id: "w2",
        en: "accomplish",
        kr: "성취하다",
        pos: "v.",
        ex: "You can accomplish anything.",
        exKr: "당신은 무엇이든 이룰 수 있다.",
    },
    {
        id: "w3",
        en: "gratitude",
        kr: "감사",
        pos: "n.",
        ex: "Express your gratitude.",
        exKr: "감사를 표현하세요.",
    },
    {
        id: "w4",
        en: "perseverance",
        kr: "인내",
        pos: "n.",
        ex: "Success requires perseverance.",
        exKr: "성공은 인내를 필요로 한다.",
    },
    {
        id: "w5",
        en: "consistent",
        kr: "꾸준한",
        pos: "adj.",
        ex: "Be consistent in your habits.",
        exKr: "습관을 꾸준히 유지하세요.",
    },
]
const DEFAULT_CONFIG: ChallengeConfig = {
    rewardCash: 10,
    passThreshold: 3,
    adSeconds: 5,
}
const DISTRACTORS = [
    "용기",
    "실패",
    "노력",
    "시간",
    "목표",
    "도전",
    "걱정",
    "행복",
    "지혜",
    "후회",
]

/* ═══════════════════════════════════════════
   Icons
   ═══════════════════════════════════════════ */
const SpeakerIcon = ({ size = 14 }: { size?: number }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
        <path
            d="M11 5L6 9H2V15H6L11 19V5Z"
            stroke="currentColor"
            strokeWidth="1.8"
            strokeLinejoin="round"
        />
        <path
            d="M15.5 8.5C16.4 9.4 17 10.7 17 12C17 13.3 16.4 14.6 15.5 15.5"
            stroke="currentColor"
            strokeWidth="1.8"
            strokeLinecap="round"
        />
    </svg>
)
const CheckIcon = ({ size = 14 }: { size?: number }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
        <path
            d="M5 13L9 17L19 7"
            stroke="currentColor"
            strokeWidth="3"
            strokeLinecap="round"
            strokeLinejoin="round"
        />
    </svg>
)
const ArrowIcon = ({ size = 14 }: { size?: number }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
        <path
            d="M9 6L15 12L9 18"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
            strokeLinejoin="round"
        />
    </svg>
)

/* ═══════════════════════════════════════════
   Helpers
   ═══════════════════════════════════════════ */
function shuffle<U>(a: U[]): U[] {
    // Fisher-Yates 셔플 — sort+random 방식은 정답이 첫 번째에 몰리는 편향 발생
    const arr = [...a]
    for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1))
        const tmp = arr[i]
        arr[i] = arr[j]
        arr[j] = tmp
    }
    return arr
}
function buildTestQuestions(words: WordItem[]): TestQuestion[] {
    return shuffle(words).map((w) => {
        const pool = [
            ...words.filter((x) => x.id !== w.id).map((x) => x.kr),
            ...DISTRACTORS,
        ].filter((x) => x !== w.kr)
        const unique = Array.from(new Set(pool))
        const wrong = shuffle(unique).slice(0, 1)
        return {
            id: w.id,
            en: w.en,
            options: shuffle([w.kr, ...wrong]),
            correct: w.kr,
        }
    })
}

/* ═══════════════════════════════════════════
   Main Component
   ═══════════════════════════════════════════ */
export default function EnglishWordChallenge({
    mode = "live",
    previewState = "empty",
    supabaseUrl = SUPA_URL_DEFAULT,
    supabaseKey = SUPA_KEY_DEFAULT,
    challengeId = "english",
}: {
    mode?: string
    previewState?: string
    supabaseUrl?: string
    supabaseKey?: string
    challengeId?: string
}) {
    const [words, setWords] = useState<WordItem[]>(FALLBACK_WORDS)
    const [config, setConfig] = useState<ChallengeConfig>(DEFAULT_CONFIG)
    const [loading, setLoading] = useState(mode === "live")
    const [studied, setStudied] = useState<Record<string, boolean>>({})
    const [streak, setStreak] = useState(0)
    const [tested, setTested] = useState(false)
    const [testPassed, setTestPassed] = useState(false)
    const [claimed, setClaimed] = useState(false)
    const [popup, setPopup] = useState<Popup>(null)
    const [currentWord, setCurrentWord] = useState<WordItem | null>(null)
    const [testQs, setTestQs] = useState<TestQuestion[]>([])
    const [testIdx, setTestIdx] = useState(0)
    const [testAnswers, setTestAnswers] = useState<Record<number, string>>({})
    const [testScore, setTestScore] = useState(0)
    const [selected, setSelected] = useState<string | null>(null)
    const [adProgress, setAdProgress] = useState(0)
    const [adTimer, setAdTimer] = useState(DEFAULT_CONFIG.adSeconds)
    const adIvRef = useRef<ReturnType<typeof setInterval> | null>(null)
    const deviceIdRef = useRef<string>("")
    const studiedRef = useRef<Record<string, boolean>>({})
    studiedRef.current = studied

    const rewardCash = config.rewardCash
    const passThreshold = config.passThreshold

    /* ─── Supabase REST helpers ─── */
    const restHeaders = () => ({
        apikey: supabaseKey,
        Authorization: `Bearer ${supabaseKey}`,
        "Content-Type": "application/json",
    })

    const persistProgress = (patch: Record<string, any>) => {
        if (mode !== "live") return
        fetch(
            `${supabaseUrl}/rest/v1/focuscoin_challenge_progress?on_conflict=device_id,challenge_id,date`,
            {
                method: "POST",
                keepalive: true,
                headers: {
                    ...restHeaders(),
                    Prefer: "resolution=merge-duplicates,return=minimal",
                },
                body: JSON.stringify({
                    device_id: deviceIdRef.current,
                    challenge_id: challengeId,
                    date: kstToday(),
                    ...patch,
                }),
            }
        ).catch(() => {})
    }

    /* ─── Live mode: DB 로드 (설정 + 오늘의 단어 + 진행 기록 + 연속일) ─── */
    useEffect(() => {
        if (mode !== "live") return
        let cancelled = false
        deviceIdRef.current = getDeviceId()
        setLoading(true)

        const h = {
            apikey: supabaseKey,
            Authorization: `Bearer ${supabaseKey}`,
            "Content-Type": "application/json",
        }

        Promise.all([
            fetch(
                `${supabaseUrl}/rest/v1/focuscoin_challenge_config?id=eq.${challengeId}&select=*`,
                { headers: h }
            ).then((r) => r.json()),
            fetch(`${supabaseUrl}/rest/v1/rpc/focuscoin_get_daily_words`, {
                method: "POST",
                headers: h,
                body: "{}",
            }).then((r) => r.json()),
            fetch(
                `${supabaseUrl}/rest/v1/focuscoin_challenge_progress?device_id=eq.${encodeURIComponent(deviceIdRef.current)}&challenge_id=eq.${challengeId}&date=eq.${kstToday()}&select=*`,
                { headers: h }
            ).then((r) => r.json()),
            fetch(`${supabaseUrl}/rest/v1/rpc/focuscoin_get_streak`, {
                method: "POST",
                headers: h,
                body: JSON.stringify({
                    p_device_id: getDeviceId(),
                    p_challenge_id: challengeId,
                }),
            }).then((r) => r.json()),
        ])
            .then(([cfgRows, dailyWords, progRows, streakVal]) => {
                if (cancelled) return
                if (Array.isArray(cfgRows) && cfgRows[0]) {
                    setConfig({
                        rewardCash: cfgRows[0].reward_cash ?? 10,
                        passThreshold: cfgRows[0].pass_threshold ?? 3,
                        adSeconds: cfgRows[0].ad_seconds ?? 5,
                    })
                }
                if (Array.isArray(dailyWords) && dailyWords.length > 0) {
                    setWords(
                        dailyWords.map((w: any) => ({
                            id: w.id,
                            en: w.en,
                            kr: w.kr,
                            pos: w.pos || "",
                            ex: w.ex || "",
                            exKr: w.ex_kr || "",
                        }))
                    )
                }
                if (Array.isArray(progRows) && progRows[0]) {
                    const p = progRows[0]
                    const m: Record<string, boolean> = {}
                    ;(p.studied_word_ids || []).forEach((id: string) => {
                        m[id] = true
                    })
                    setStudied(m)
                    setTested(p.test_score !== null && p.test_score !== undefined)
                    setTestPassed(!!p.test_passed)
                    if (typeof p.test_score === "number")
                        setTestScore(p.test_score)
                    setClaimed(!!p.claimed)
                }
                if (typeof streakVal === "number") setStreak(streakVal)
                setLoading(false)
            })
            .catch(() => {
                if (cancelled) return
                // DB 연결 실패 → 데모 단어로 동작 (저장은 안 됨)
                setLoading(false)
            })

        return () => {
            cancelled = true
        }
    }, [mode, supabaseUrl, supabaseKey, challengeId])

    /* ─── Preview state sync (미리보기 모드 전용) ─── */
    useEffect(() => {
        if (mode !== "preview") return
        if (adIvRef.current) {
            clearInterval(adIvRef.current)
            adIvRef.current = null
        }
        setWords(FALLBACK_WORDS)
        setConfig(DEFAULT_CONFIG)
        setLoading(false)
        setPopup(null)
        setSelected(null)
        setCurrentWord(null)
        setAdProgress(0)
        setAdTimer(DEFAULT_CONFIG.adSeconds)

        const allStudiedMap: Record<string, boolean> = {
            w1: true,
            w2: true,
            w3: true,
            w4: true,
            w5: true,
        }

        switch (previewState) {
            case "empty":
                setStudied({})
                setStreak(3)
                setTested(false)
                setTestPassed(false)
                setClaimed(false)
                break
            case "studying":
                setStudied({ w1: true, w2: true, w3: true })
                setStreak(3)
                setTested(false)
                setTestPassed(false)
                setClaimed(false)
                break
            case "ready":
                setStudied(allStudiedMap)
                setStreak(3)
                setTested(false)
                setTestPassed(false)
                setClaimed(false)
                break
            case "in_test":
                setStudied(allStudiedMap)
                setStreak(3)
                setTested(false)
                setTestPassed(false)
                setClaimed(false)
                setTestQs(buildTestQuestions(FALLBACK_WORDS))
                setTestIdx(0)
                setTestAnswers({})
                setPopup("test")
                break
            case "test_pass":
                setStudied(allStudiedMap)
                setStreak(3)
                setTested(true)
                setTestPassed(true)
                setTestScore(5)
                setClaimed(false)
                setTestQs(buildTestQuestions(FALLBACK_WORDS))
                setPopup("test_result")
                break
            case "test_fail":
                setStudied(allStudiedMap)
                setStreak(3)
                setTested(true)
                setTestPassed(false)
                setTestScore(2)
                setClaimed(false)
                setTestQs(buildTestQuestions(FALLBACK_WORDS))
                setPopup("test_result")
                break
            case "ad_playing":
                setStudied(allStudiedMap)
                setStreak(3)
                setTested(true)
                setTestPassed(true)
                setTestScore(5)
                setClaimed(false)
                setAdProgress(60)
                setAdTimer(2)
                setPopup("ad")
                break
            case "reward":
                setStudied(allStudiedMap)
                setStreak(3)
                setTested(true)
                setTestPassed(true)
                setTestScore(5)
                setClaimed(false)
                setPopup("reward")
                break
            case "claimed":
                setStudied(allStudiedMap)
                setStreak(4)
                setTested(true)
                setTestPassed(true)
                setTestScore(5)
                setClaimed(true)
                break
            default:
                setStudied({})
                setStreak(3)
                setTested(false)
                setTestPassed(false)
                setClaimed(false)
        }
    }, [mode, previewState])

    useEffect(() => {
        return () => {
            if (adIvRef.current) clearInterval(adIvRef.current)
        }
    }, [])

    const studiedCount = Object.values(studied).filter(Boolean).length
    const allStudied = words.length > 0 && studiedCount >= words.length
    const pct =
        words.length > 0
            ? Math.round((studiedCount / words.length) * 100)
            : 0

    /* ─── Handlers ─── */
    const handleStudyWord = (w: WordItem) => {
        setCurrentWord(w)
        setPopup("study")
    }
    const handleMarkStudied = () => {
        if (currentWord) {
            const next = { ...studiedRef.current, [currentWord.id]: true }
            setStudied(next)
            setPopup(null)
            setCurrentWord(null)
            persistProgress({
                studied_word_ids: Object.keys(next).filter((k) => next[k]),
            })
        }
    }
    const handleSpeak = (text: string) => {
        if (typeof window !== "undefined" && "speechSynthesis" in window) {
            const u = new SpeechSynthesisUtterance(text)
            u.lang = "en-US"
            u.rate = 0.9
            window.speechSynthesis.cancel()
            window.speechSynthesis.speak(u)
        }
    }
    const handleStartTest = () => {
        setTestQs(buildTestQuestions(words))
        setTestIdx(0)
        setTestAnswers({})
        setSelected(null)
        setPopup("test")
    }
    const handleAnswer = (opt: string) => {
        if (selected) return
        setSelected(opt)
        const newAnswers = { ...testAnswers, [testIdx]: opt }
        setTestAnswers(newAnswers)
        setTimeout(() => {
            setSelected(null)
            if (testIdx < testQs.length - 1) {
                setTestIdx(testIdx + 1)
            } else {
                let score = 0
                testQs.forEach((q, i) => {
                    if (newAnswers[i] === q.correct) score++
                })
                const passed = score >= passThreshold
                setTestScore(score)
                setTestPassed(passed)
                setTested(true)
                setPopup("test_result")
                persistProgress({
                    test_score: score,
                    test_passed: passed,
                })
            }
        }, 1100)
    }
    const handleRetry = () => {
        setStudied({})
        setTested(false)
        setTestPassed(false)
        setPopup(null)
        persistProgress({
            studied_word_ids: [],
            test_score: null,
            test_passed: false,
        })
    }
    const handleWatchAd = () => {
        const total = config.adSeconds
        setPopup("ad")
        setAdProgress(0)
        setAdTimer(total)
        let e = 0
        adIvRef.current = setInterval(() => {
            e++
            setAdProgress(Math.min((e / total) * 100, 100))
            setAdTimer(Math.max(total - e, 0))
            if (e >= total) {
                if (adIvRef.current) {
                    clearInterval(adIvRef.current)
                    adIvRef.current = null
                }
                setTimeout(() => setPopup("reward"), 250)
            }
        }, 1000)
    }
    const handleClaim = () => {
        setClaimed(true)
        setStreak((p) => p + 1)
        setPopup(null)
        persistProgress({
            claimed: true,
            claimed_at: new Date().toISOString(),
        })
    }
    const handleClosePopup = () => {
        if (adIvRef.current) {
            clearInterval(adIvRef.current)
            adIvRef.current = null
        }
        setPopup(null)
    }

    /* ─── CTA decision ─── */
    let ctaHint = ""
    let ctaContent: React.ReactNode = null
    if (claimed) {
        ctaContent = (
            <button disabled style={ctaSoftSty} className="ewc-cta">
                오늘 챌린지 완료 ✓
            </button>
        )
    } else if (tested && testPassed) {
        ctaContent = (
            <button
                onClick={handleWatchAd}
                style={ctaGreenSty}
                className="ewc-cta ewc-cta-green"
            >
                <span>+{rewardCash}캐시 받기</span>
                <span style={chipSty} className="ewc-cta-chip">
                    광고
                </span>
            </button>
        )
    } else if (tested && !testPassed) {
        ctaContent = (
            <button
                onClick={handleRetry}
                style={ctaDarkSty}
                className="ewc-cta"
            >
                다시 학습하기
            </button>
        )
    } else if (allStudied) {
        ctaHint = `${words.length}문제 중 ${passThreshold}문제만 맞히면 통과`
        ctaContent = (
            <button
                onClick={handleStartTest}
                style={ctaGreenSty}
                className="ewc-cta ewc-cta-green"
            >
                <span>테스트 시작</span>
                <span style={chipSty} className="ewc-cta-chip">
                    +{rewardCash}캐시
                </span>
            </button>
        )
    } else {
        ctaContent = (
            <button disabled style={ctaDisabledSty} className="ewc-cta">
                단어 {words.length - studiedCount}개 남았어요
            </button>
        )
    }

    /* ─── Loading placeholder (모든 훅 선언 뒤, 메인 return 직전 — Rules of Hooks) ─── */
    if (loading) {
        return (
            <div className="ewc" style={rootSty}>
                <style>{CSS}</style>
                <div style={loadingWrapSty}>
                    <span className="ewc-loading-dot" style={loadingDotSty} />
                    <span className="ewc-loading-dot" style={loadingDotSty} />
                    <span className="ewc-loading-dot" style={loadingDotSty} />
                </div>
            </div>
        )
    }

    return (
        <div className="ewc" style={rootSty}>
            <style>{CSS}</style>

            {/* Header */}
            <div style={hdrSty}>
                <span style={hdrTitleSty}>영단어 챌린지</span>
            </div>

            {/* Body */}
            <div className="ewc-body" style={bodySty}>
                {/* Hero card */}
                <div style={heroSty}>
                    <div style={heroTopSty}>
                        <span style={heroLabelSty}>오늘의 학습</span>
                    </div>
                    <div style={heroProgressRowSty}>
                        <span style={heroNumSty}>{studiedCount}</span>
                        <span style={heroTotalSty}>/ {words.length}</span>
                        <span style={heroPctSty}>{pct}%</span>
                    </div>
                    <div style={heroBarSty}>
                        <div style={{ ...heroBarFillSty, width: `${pct}%` }} />
                    </div>
                    <div style={heroFootSty}>
                        <span style={heroFootItemSty}>
                            {allStudied
                                ? testPassed
                                    ? "테스트 통과"
                                    : tested
                                      ? "재도전"
                                      : "테스트 가능"
                                : `${words.length - studiedCount}개 남음`}
                        </span>
                    </div>
                </div>

                {/* Section head */}
                <div style={sectionHeadSty}>
                    <span style={sectionTitleSty}>오늘의 단어</span>
                    <span style={sectionMetaSty}>
                        {studiedCount} / {words.length}
                    </span>
                </div>

                {/* Word list */}
                <div style={wordListSty}>
                    {words.map((w, i) => {
                        const done = !!studied[w.id]
                        return (
                            <button
                                key={w.id}
                                className="ewc-word-row"
                                onClick={() => handleStudyWord(w)}
                                style={{
                                    ...wordRowSty,
                                    borderTop:
                                        i === 0
                                            ? "none"
                                            : `1px solid ${T.cDivider}`,
                                }}
                            >
                                <span style={done ? checkOnSty : checkOffSty}>
                                    {done && (
                                        <span
                                            style={{
                                                color: "#FFF",
                                                display: "flex",
                                            }}
                                        >
                                            <CheckIcon size={13} />
                                        </span>
                                    )}
                                </span>
                                <div style={wordInfoSty}>
                                    <div
                                        style={{
                                            ...wordEnSty,
                                            color: done ? T.cText4 : T.cText,
                                        }}
                                    >
                                        {w.en}
                                        <span style={wordPosSty}>{w.pos}</span>
                                    </div>
                                    <div
                                        style={{
                                            ...wordKrSty,
                                            color: done ? T.cText2 : T.cText4,
                                        }}
                                    >
                                        {done ? w.kr : "탭하여 학습"}
                                    </div>
                                </div>
                                {!done && (
                                    <span style={wordArrowSty}>
                                        <ArrowIcon size={14} />
                                    </span>
                                )}
                            </button>
                        )
                    })}
                </div>
            </div>

            {/* CTA */}
            <div style={ctaWrapSty}>
                {ctaHint && <div style={ctaHintSty}>{ctaHint}</div>}
                {ctaContent}
            </div>

            {/* ─── Study Popup ─── */}
            {popup === "study" && currentWord && (
                <div
                    className="ewc-overlay"
                    style={overlayBottomSty}
                    onClick={handleClosePopup}
                >
                    <div
                        className="ewc-sheet-bottom"
                        style={sheetBottomSty}
                        onClick={(e) => e.stopPropagation()}
                    >
                        <div style={sheetHandleSty} />
                        <div style={studyHeadSty}>
                            <span style={studyEnSty}>{currentWord.en}</span>
                            <button
                                style={spkSty}
                                onClick={() => handleSpeak(currentWord.en)}
                            >
                                <SpeakerIcon size={14} />
                            </button>
                        </div>
                        <div style={studyPosSty}>{currentWord.pos}</div>
                        <div style={studySectionSty}>
                            <div style={{ ...studyTagSty, color: T.cGreen }}>
                                뜻
                            </div>
                            <div style={studyMeanSty}>{currentWord.kr}</div>
                        </div>
                        <div style={studySectionSty}>
                            <div style={studyTagSty}>예문</div>
                            <div style={studyExEnSty}>{currentWord.ex}</div>
                            <div style={studyExKrSty}>{currentWord.exKr}</div>
                        </div>
                        {studied[currentWord.id] ? (
                            <button
                                className="ewc-m-btn"
                                style={mBtnSoftSty}
                                onClick={handleClosePopup}
                            >
                                닫기
                            </button>
                        ) : (
                            <button
                                className="ewc-m-btn"
                                style={mBtnGreenSty}
                                onClick={handleMarkStudied}
                            >
                                학습 완료
                            </button>
                        )}
                    </div>
                </div>
            )}

            {/* ─── Test Popup ─── */}
            {popup === "test" &&
                testQs.length > 0 &&
                (() => {
                    const q = testQs[testIdx]
                    if (!q) return null
                    return (
                        <div className="ewc-overlay" style={overlayCenterSty}>
                            <div
                                className="ewc-sheet-center"
                                style={sheetCenterSty}
                            >
                                <div style={testCounterSty}>
                                    {testIdx + 1} / {testQs.length}
                                </div>
                                <div style={testPbSty}>
                                    <div
                                        style={{
                                            ...testPbFillSty,
                                            width: `${((testIdx + 1) / testQs.length) * 100}%`,
                                        }}
                                    />
                                </div>
                                <div style={feedbackAreaSty}>
                                    {!selected && (
                                        <div style={testPromptSty}>
                                            다음 단어의 뜻은?
                                        </div>
                                    )}
                                    {selected === q.correct && (
                                        <div
                                            className="ewc-fb"
                                            style={fbCorrectSty}
                                        >
                                            🎉 정답!
                                        </div>
                                    )}
                                    {selected && selected !== q.correct && (
                                        <div
                                            className="ewc-fb"
                                            style={fbWrongSty}
                                        >
                                            😢 아쉬워요
                                        </div>
                                    )}
                                </div>
                                <div style={testWordSty}>{q.en}</div>
                                <button
                                    style={testSpkSty}
                                    onClick={() => handleSpeak(q.en)}
                                >
                                    <SpeakerIcon size={13} />
                                    <span>발음</span>
                                </button>
                                <div style={testOptsSty}>
                                    {q.options.map((opt, i) => {
                                        const isCorrect =
                                            selected && opt === q.correct
                                        const isWrong =
                                            selected &&
                                            opt === selected &&
                                            opt !== q.correct
                                        let optStyle: React.CSSProperties = {
                                            ...testOptBaseSty,
                                        }
                                        let className = "ewc-test-opt"
                                        if (isCorrect) {
                                            optStyle = {
                                                ...optStyle,
                                                ...testOptCorrectSty,
                                            }
                                            className += " ewc-opt-correct"
                                        } else if (isWrong) {
                                            optStyle = {
                                                ...optStyle,
                                                ...testOptWrongSty,
                                            }
                                        }
                                        return (
                                            <button
                                                key={i}
                                                className={className}
                                                style={optStyle}
                                                onClick={() =>
                                                    handleAnswer(opt)
                                                }
                                                disabled={!!selected}
                                            >
                                                <span style={optLeftSty}>
                                                    <span
                                                        style={{
                                                            ...optNumSty,
                                                            background:
                                                                isCorrect
                                                                    ? T.cGreen
                                                                    : isWrong
                                                                      ? T.cRed
                                                                      : T.cBorder,
                                                            color:
                                                                isCorrect ||
                                                                isWrong
                                                                    ? "#FFF"
                                                                    : T.cText2,
                                                        }}
                                                    >
                                                        {i + 1}
                                                    </span>
                                                    {opt}
                                                </span>
                                            </button>
                                        )
                                    })}
                                </div>
                            </div>
                        </div>
                    )
                })()}

            {/* ─── Test Result Popup ─── */}
            {popup === "test_result" && (
                <div className="ewc-overlay" style={overlayCenterSty}>
                    <div className="ewc-sheet-center" style={sheetCenterSty}>
                        <div
                            style={{
                                ...scoreCardSty,
                                background: testPassed ? T.cGreenBg : T.cRedBg,
                            }}
                        >
                            <div style={scoreEmojiSty}>
                                {testPassed ? "🎉" : "😢"}
                            </div>
                            <div
                                style={{
                                    ...scoreLabelSty,
                                    color: testPassed ? T.cGreenDk : T.cRedDk,
                                }}
                            >
                                {testPassed ? "테스트 통과" : "아쉬워요"}
                            </div>
                            <div
                                style={{
                                    ...scoreNumSty,
                                    color: testPassed ? T.cGreenDk : T.cRedDk,
                                }}
                            >
                                {testScore}
                                <span style={scoreTotalSty}>
                                    {" "}
                                    / {testQs.length || words.length}
                                </span>
                            </div>
                        </div>
                        {testPassed ? (
                            <>
                                <button
                                    className="ewc-m-btn"
                                    style={mBtnGreenSty}
                                    onClick={() => {
                                        handleClosePopup()
                                        setTimeout(handleWatchAd, 150)
                                    }}
                                >
                                    +{rewardCash}캐시 받기
                                </button>
                                <button
                                    className="ewc-m-btn"
                                    style={mBtnGhostSty}
                                    onClick={handleClosePopup}
                                >
                                    닫기
                                </button>
                            </>
                        ) : (
                            <>
                                <button
                                    className="ewc-m-btn"
                                    style={mBtnDarkSty}
                                    onClick={handleRetry}
                                >
                                    다시 학습하기
                                </button>
                                <button
                                    className="ewc-m-btn"
                                    style={mBtnGhostSty}
                                    onClick={handleClosePopup}
                                >
                                    닫기
                                </button>
                            </>
                        )}
                    </div>
                </div>
            )}

            {/* ─── Ad Popup ─── */}
            {popup === "ad" && (
                <div className="ewc-overlay" style={overlayCenterSty}>
                    <div
                        className="ewc-sheet-center"
                        style={{ ...sheetCenterSty, maxWidth: 340 }}
                    >
                        <div
                            style={{
                                fontSize: T.tSm,
                                fontWeight: T.wLabel,
                                color: T.cText2,
                                marginBottom: 14,
                                textAlign: "center",
                            }}
                        >
                            광고 시청 중
                        </div>
                        <div style={adCanvasSty}>
                            <span
                                style={{
                                    fontSize: T.tSm,
                                    color: T.cText3,
                                    fontWeight: T.wBody,
                                }}
                            >
                                광고 영역 (AdMob)
                            </span>
                        </div>
                        <div style={adBarSty}>
                            <div
                                style={{
                                    ...adBarFillSty,
                                    width: `${adProgress}%`,
                                }}
                            />
                        </div>
                        <div style={adTimerSty}>
                            {adTimer > 0 ? `${adTimer}초 남음` : "완료"}
                        </div>
                    </div>
                </div>
            )}

            {/* ─── Reward Popup ─── */}
            {popup === "reward" && (
                <div className="ewc-overlay" style={overlayCenterSty}>
                    <div className="ewc-sheet-center" style={sheetCenterSty}>
                        <div style={scoreEmojiSty}>🎉</div>
                        <div style={mTitleSty}>
                            {rewardCash}캐시가 적립되었어요
                        </div>
                        <div style={mSubSty}>내일도 학습하러 와주세요</div>
                        <button
                            className="ewc-m-btn"
                            style={mBtnGreenSty}
                            onClick={handleClaim}
                        >
                            확인
                        </button>
                    </div>
                </div>
            )}
        </div>
    )
}

/* ═══════════════════════════════════════════
   Styles
   ═══════════════════════════════════════════ */
const rootSty: React.CSSProperties = {
    width: "100%",
    height: "100%",
    position: "relative",
    background: T.cBg,
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

const hdrSty: React.CSSProperties = { padding: "20px 20px 8px", flexShrink: 0 }
const hdrTitleSty: React.CSSProperties = {
    fontSize: T.tLg,
    fontWeight: T.wBold,
    letterSpacing: -0.3,
    color: T.cText,
}

const bodySty: React.CSSProperties = {
    flex: 1,
    overflowY: "auto",
    padding: "8px 16px 0",
}

const heroSty: React.CSSProperties = {
    background: T.cBg,
    border: `1px solid ${T.cBorder}`,
    borderRadius: T.rCard,
    padding: 20,
    marginBottom: 12,
}
const heroTopSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: 16,
}
const heroLabelSty: React.CSSProperties = {
    fontSize: T.tSm,
    color: T.cText2,
    fontWeight: T.wLabel,
}
const rewardChipSty: React.CSSProperties = {
    display: "inline-flex",
    alignItems: "center",
    padding: "4px 10px",
    borderRadius: T.rPill,
    background: T.cGreenBg,
    color: T.cGreenDk,
    fontSize: T.tXs,
    fontWeight: T.wBold,
    letterSpacing: -0.1,
}
const rewardChipDoneSty: React.CSSProperties = {
    ...rewardChipSty,
    background: T.cDivider,
    color: T.cText2,
}
const heroProgressRowSty: React.CSSProperties = {
    display: "flex",
    alignItems: "baseline",
    gap: 4,
    marginBottom: 14,
}
const heroNumSty: React.CSSProperties = {
    fontSize: T.t3xl,
    fontWeight: T.wBold,
    letterSpacing: -1.5,
    color: T.cText,
    lineHeight: 1,
}
const heroTotalSty: React.CSSProperties = {
    fontSize: T.tXl,
    fontWeight: T.wLabel,
    color: T.cText4,
    letterSpacing: -0.8,
}
const heroPctSty: React.CSSProperties = {
    marginLeft: "auto",
    fontSize: T.tSm,
    fontWeight: T.wBold,
    color: T.cGreen,
}
const heroBarSty: React.CSSProperties = {
    height: 8,
    background: T.cDivider,
    borderRadius: T.rPill,
    overflow: "hidden",
}
const heroBarFillSty: React.CSSProperties = {
    height: "100%",
    background: T.cGreen,
    borderRadius: T.rPill,
    transition: "width 0.6s cubic-bezier(0.4, 0, 0.2, 1)",
}
const heroFootSty: React.CSSProperties = {
    marginTop: 16,
    paddingTop: 14,
    borderTop: `1px solid ${T.cDivider}`,
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
}
const heroFootItemSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: 5,
    fontSize: T.tSm,
    color: T.cText2,
    fontWeight: T.wBody,
}

const sectionHeadSty: React.CSSProperties = {
    padding: "16px 4px 10px",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "baseline",
}
const sectionTitleSty: React.CSSProperties = {
    fontSize: T.tMd,
    fontWeight: T.wLabel,
    color: T.cText,
    letterSpacing: -0.2,
}
const sectionMetaSty: React.CSSProperties = {
    fontSize: T.tSm,
    color: T.cText3,
    fontWeight: T.wBody,
}

const wordListSty: React.CSSProperties = {
    background: T.cBg,
    border: `1px solid ${T.cBorder}`,
    borderRadius: T.rCard,
    overflow: "hidden",
    marginBottom: 12,
}
const wordRowSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: 13,
    padding: "14px 16px",
    cursor: "pointer",
    background: "transparent",
    border: "none",
    fontFamily: FONT,
    textAlign: "left",
    width: "100%",
    transition: "background 0.12s",
}
const checkOffSty: React.CSSProperties = {
    width: 24,
    height: 24,
    borderRadius: "50%",
    border: `1.5px solid #DCDFE3`,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
    color: "#FFF",
    transition: "all 0.2s",
}
const checkOnSty: React.CSSProperties = {
    ...checkOffSty,
    border: "none",
    background: T.cGreen,
}
const wordInfoSty: React.CSSProperties = { flex: 1, minWidth: 0 }
const wordEnSty: React.CSSProperties = {
    fontSize: T.tLg,
    fontWeight: T.wLabel,
    letterSpacing: -0.3,
    lineHeight: 1.3,
}
const wordPosSty: React.CSSProperties = {
    fontSize: T.tXs,
    color: T.cText4,
    fontWeight: T.wBody,
    marginLeft: 5,
}
const wordKrSty: React.CSSProperties = {
    fontSize: T.tSm,
    marginTop: 3,
    fontWeight: T.wBody,
}
const wordArrowSty: React.CSSProperties = {
    color: T.cText4,
    flexShrink: 0,
    display: "flex",
}

const ctaWrapSty: React.CSSProperties = {
    padding: "14px 20px 24px",
    flexShrink: 0,
}
const ctaHintSty: React.CSSProperties = {
    textAlign: "center",
    fontSize: T.tSm,
    color: T.cText2,
    marginBottom: 10,
    fontWeight: T.wLabel,
}
const ctaBaseSty: React.CSSProperties = {
    width: "100%",
    padding: "17px 20px",
    border: "none",
    borderRadius: T.rBtn,
    fontSize: T.tLg,
    fontWeight: T.wBold,
    fontFamily: FONT,
    cursor: "pointer",
    letterSpacing: -0.3,
    transition: "transform 0.1s, opacity 0.15s",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 10,
}
const ctaGreenSty: React.CSSProperties = {
    ...ctaBaseSty,
    background: T.cGreen,
    color: "#FFF",
}
const ctaDarkSty: React.CSSProperties = {
    ...ctaBaseSty,
    background: T.cText,
    color: "#FFF",
}
const ctaSoftSty: React.CSSProperties = {
    ...ctaBaseSty,
    background: T.cDivider,
    color: T.cText,
    cursor: "default",
}
const ctaDisabledSty: React.CSSProperties = {
    ...ctaBaseSty,
    background: T.cDivider,
    color: T.cText4,
    cursor: "not-allowed",
}
const chipSty: React.CSSProperties = {
    display: "inline-flex",
    alignItems: "center",
    padding: "4px 10px",
    borderRadius: T.rPill,
    background: "rgba(255,255,255,0.22)",
    fontSize: T.tXs,
    fontWeight: T.wBold,
    letterSpacing: -0.1,
}

const overlayBottomSty: React.CSSProperties = {
    position: "absolute",
    inset: 0,
    background: "rgba(20,23,28,0.5)",
    display: "flex",
    alignItems: "flex-end",
    justifyContent: "center",
    zIndex: 100,
}
const overlayCenterSty: React.CSSProperties = {
    ...overlayBottomSty,
    alignItems: "center",
    padding: 20,
}
const sheetBottomSty: React.CSSProperties = {
    width: "100%",
    background: T.cBg,
    borderRadius: `${T.rSheet}px ${T.rSheet}px 0 0`,
    padding: "24px 22px 28px",
}
const sheetCenterSty: React.CSSProperties = {
    width: "100%",
    maxWidth: 320,
    background: T.cBg,
    borderRadius: T.rCard,
    padding: "28px 22px 22px",
    textAlign: "center",
}
const sheetHandleSty: React.CSSProperties = {
    width: 36,
    height: 4,
    background: "#E5E8EB",
    borderRadius: T.rPill,
    margin: "0 auto 18px",
}

const studyHeadSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: 10,
    marginBottom: 4,
}
const studyEnSty: React.CSSProperties = {
    fontSize: T.t2xl,
    fontWeight: T.wBold,
    letterSpacing: -1,
    color: T.cText,
}
const spkSty: React.CSSProperties = {
    background: T.cDivider,
    border: "none",
    borderRadius: T.rBtn,
    padding: "8px 10px",
    cursor: "pointer",
    color: T.cText,
    display: "inline-flex",
    alignItems: "center",
    fontFamily: FONT,
}
const studyPosSty: React.CSSProperties = {
    fontSize: T.tSm,
    color: T.cText3,
    marginBottom: 22,
    fontWeight: T.wBody,
    textAlign: "left",
}
const studySectionSty: React.CSSProperties = {
    marginBottom: 18,
    textAlign: "left",
}
const studyTagSty: React.CSSProperties = {
    fontSize: T.tXs,
    color: T.cText3,
    fontWeight: T.wBold,
    marginBottom: 6,
    letterSpacing: 0.3,
}
const studyMeanSty: React.CSSProperties = {
    fontSize: T.tXl,
    fontWeight: T.wLabel,
    letterSpacing: -0.3,
    color: T.cText,
}
const studyExEnSty: React.CSSProperties = {
    fontSize: T.tMd,
    lineHeight: 1.55,
    color: T.cText,
    fontWeight: T.wBody,
}
const studyExKrSty: React.CSSProperties = {
    fontSize: T.tSm,
    color: T.cText2,
    lineHeight: 1.55,
    marginTop: 4,
    fontWeight: T.wBody,
}

const mBtnBaseSty: React.CSSProperties = {
    width: "100%",
    padding: "16px 0",
    border: "none",
    borderRadius: T.rBtn,
    fontSize: T.tLg,
    fontWeight: T.wBold,
    fontFamily: FONT,
    cursor: "pointer",
    marginTop: 8,
    letterSpacing: -0.3,
    transition: "transform 0.1s",
}
const mBtnGreenSty: React.CSSProperties = {
    ...mBtnBaseSty,
    background: T.cGreen,
    color: "#FFF",
}
const mBtnDarkSty: React.CSSProperties = {
    ...mBtnBaseSty,
    background: T.cText,
    color: "#FFF",
}
const mBtnSoftSty: React.CSSProperties = {
    ...mBtnBaseSty,
    background: T.cDivider,
    color: T.cText,
}
const mBtnGhostSty: React.CSSProperties = {
    ...mBtnBaseSty,
    background: "transparent",
    color: T.cText3,
    padding: "12px 0",
    fontWeight: T.wLabel,
    fontSize: T.tMd,
}
const mTitleSty: React.CSSProperties = {
    fontSize: T.tXl,
    fontWeight: T.wBold,
    marginBottom: 6,
    letterSpacing: -0.4,
    color: T.cText,
}
const mSubSty: React.CSSProperties = {
    fontSize: T.tMd,
    color: T.cText2,
    marginBottom: 22,
    lineHeight: 1.55,
    fontWeight: T.wBody,
}

const testCounterSty: React.CSSProperties = {
    fontSize: T.tSm,
    fontWeight: T.wBold,
    color: T.cText,
    marginBottom: 8,
    textAlign: "left",
}
const testPbSty: React.CSSProperties = {
    height: 4,
    background: T.cDivider,
    borderRadius: T.rPill,
    overflow: "hidden",
    marginBottom: 22,
}
const testPbFillSty: React.CSSProperties = {
    height: "100%",
    background: T.cGreen,
    borderRadius: T.rPill,
    transition: "width 0.4s ease",
}
const feedbackAreaSty: React.CSSProperties = {
    minHeight: 32,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 10,
}
const testPromptSty: React.CSSProperties = {
    fontSize: T.tSm,
    color: T.cText3,
    textAlign: "center",
    fontWeight: T.wBody,
}
const fbCorrectSty: React.CSSProperties = {
    fontSize: T.tMd,
    fontWeight: T.wBold,
    padding: "6px 14px",
    borderRadius: T.rPill,
    color: T.cGreenDk,
    background: T.cGreenBg,
}
const fbWrongSty: React.CSSProperties = {
    ...fbCorrectSty,
    color: T.cRedDk,
    background: T.cRedBg,
}
const testWordSty: React.CSSProperties = {
    fontSize: T.t3xl,
    fontWeight: T.wBold,
    letterSpacing: -1.2,
    textAlign: "center",
    marginBottom: 12,
    color: T.cText,
}
const testSpkSty: React.CSSProperties = {
    background: T.cDivider,
    border: "none",
    borderRadius: T.rPill,
    padding: "6px 12px",
    cursor: "pointer",
    color: T.cText,
    display: "flex",
    width: "fit-content",
    alignItems: "center",
    gap: 5,
    fontSize: T.tXs,
    fontWeight: T.wLabel,
    fontFamily: FONT,
    margin: "0 auto 22px",
}
const testOptsSty: React.CSSProperties = {
    display: "flex",
    flexDirection: "column",
    gap: 10,
    marginTop: "auto",
}
const testOptBaseSty: React.CSSProperties = {
    padding: "22px 20px",
    background: T.cCard,
    border: "2px solid transparent",
    borderRadius: T.rBtn,
    fontSize: T.tLg,
    fontWeight: T.wBold,
    color: T.cText,
    fontFamily: FONT,
    cursor: "pointer",
    textAlign: "left",
    letterSpacing: -0.2,
    transition: "all 0.18s, transform 0.1s",
    width: "100%",
}
const testOptCorrectSty: React.CSSProperties = {
    background: T.cGreenBg,
    borderColor: T.cGreen,
    color: T.cGreenDk,
}
const testOptWrongSty: React.CSSProperties = {
    background: T.cRedBg,
    borderColor: T.cRed,
    color: T.cRedDk,
}
const optLeftSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: 14,
}
const optNumSty: React.CSSProperties = {
    width: 28,
    height: 28,
    borderRadius: "50%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: T.tSm,
    fontWeight: T.wBold,
    flexShrink: 0,
}

const scoreCardSty: React.CSSProperties = {
    padding: "28px 16px",
    borderRadius: T.rCard,
    marginBottom: 16,
    textAlign: "center",
}
const scoreEmojiSty: React.CSSProperties = {
    fontSize: T.t3xl,
    marginBottom: 8,
    lineHeight: 1,
}
const scoreLabelSty: React.CSSProperties = {
    fontSize: T.tSm,
    marginBottom: 8,
    fontWeight: T.wBold,
    letterSpacing: 0.3,
}
const scoreNumSty: React.CSSProperties = {
    fontSize: T.t3xl,
    fontWeight: T.wBold,
    letterSpacing: -2,
    lineHeight: 1,
}
const scoreTotalSty: React.CSSProperties = {
    fontSize: T.tXl,
    color: T.cText4,
    fontWeight: T.wLabel,
}

const adCanvasSty: React.CSSProperties = {
    height: 200,
    background: T.cDivider,
    borderRadius: T.rBtn,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 16,
}
const adBarSty: React.CSSProperties = {
    height: 4,
    background: T.cBorder,
    borderRadius: T.rPill,
    overflow: "hidden",
    marginBottom: 8,
}
const adBarFillSty: React.CSSProperties = {
    height: "100%",
    background: T.cGreen,
    transition: "width 0.8s linear",
    borderRadius: T.rPill,
}
const adTimerSty: React.CSSProperties = {
    fontSize: T.tSm,
    color: T.cText2,
    textAlign: "center",
    fontWeight: T.wLabel,
}

/* ═══════════════════════════════════════════
   Property Controls (Framer 패널)
   ═══════════════════════════════════════════ */
addPropertyControls(EnglishWordChallenge, {
    mode: {
        type: ControlType.Enum,
        title: "데이터",
        options: ["live", "preview"],
        optionTitles: ["실제 DB 연동", "미리보기(데모)"],
        defaultValue: "live",
    },
    previewState: {
        type: ControlType.Enum,
        title: "미리보기",
        options: [
            "empty",
            "studying",
            "ready",
            "in_test",
            "test_pass",
            "test_fail",
            "ad_playing",
            "reward",
            "claimed",
        ],
        optionTitles: [
            "시작 전",
            "학습 중 (3/5)",
            "테스트 준비 (5/5)",
            "테스트 진행 중",
            "테스트 통과 결과",
            "테스트 실패 결과",
            "광고 재생 중",
            "캐시 적립 팝업",
            "오늘 챌린지 완료",
        ],
        defaultValue: "empty",
        hidden: (props: any) => props.mode !== "preview",
    },
    supabaseUrl: {
        type: ControlType.String,
        title: "Supabase URL",
        defaultValue: SUPA_URL_DEFAULT,
        hidden: (props: any) => props.mode !== "live",
    },
    supabaseKey: {
        type: ControlType.String,
        title: "Supabase Key",
        defaultValue: SUPA_KEY_DEFAULT,
        hidden: (props: any) => props.mode !== "live",
    },
    challengeId: {
        type: ControlType.String,
        title: "챌린지 ID",
        defaultValue: "english",
        hidden: (props: any) => props.mode !== "live",
    },
})
