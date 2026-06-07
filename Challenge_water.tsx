import React, { useState, useRef, useEffect } from "react"
import { addPropertyControls, ControlType } from "framer"

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
}
const FONT =
    "'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
const GOAL = 8
const COOLDOWN_MS = 1 * 60 * 1000
const SUPABASE_URL = "https://gjnwriqewsrwpbtxqbea.supabase.co"
const SUPABASE_ANON_KEY =
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdqbndyaXFld3Nyd3BidHhxYmVhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzOTIzNjQsImV4cCI6MjA5NTk2ODM2NH0.XAcRHkdHh8WmwhJgYht__CPmopQadvWVR3h7c8uFswU"

function getQueryParam(name: string): string {
    if (typeof window === "undefined") return ""
    try {
        return new URLSearchParams(window.location.search).get(name) || ""
    } catch {
        return ""
    }
}

async function rpc(
    fn: string,
    body: Record<string, unknown>
): Promise<any> {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/rpc/${fn}`, {
        method: "POST",
        keepalive: true,
        headers: {
            apikey: SUPABASE_ANON_KEY,
            Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
            "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
    })
    if (!res.ok) throw new Error(`rpc_${fn}_${res.status}`)
    return res.json()
}

type ClaimResult = "success" | "duplicate" | "error"

async function claimReward(
    userId: string,
    mockResult?: ClaimResult
): Promise<ClaimResult> {
    if (mockResult || !userId) {
        await new Promise((r) => setTimeout(r, 600))
        return mockResult || "success"
    }
    try {
        const r = await rpc("claim_water_reward", { p_user_id: userId })
        if (r && r.ok) return "success"
        if (r && (r.error === "duplicate" || r.error === "already_claimed"))
            return "duplicate"
        return "error"
    } catch {
        return "error"
    }
}

const CSS = `
  .wc *, .wc *::before, .wc *::after { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  .wc button { -webkit-appearance: none; appearance: none; font-family: inherit; color: inherit; -webkit-touch-callout: none; user-select: none; -webkit-user-select: none; }
  .wc button:focus { outline: none; }
  .wc button::-moz-focus-inner { border: 0; padding: 0; }
  .wc-body::-webkit-scrollbar { width: 0; }
  .wc-body { scrollbar-width: none; -webkit-overflow-scrolling: touch; }
  .wc-cta-chip { animation: wc-cta-chip-pulse 1.8s ease-in-out infinite; }
  @keyframes wc-cta-chip-pulse { 0%,100%{transform:scale(1);} 50%{transform:scale(1.08);} }
  .wc-cup-fill { animation: wc-cup-pop 0.36s cubic-bezier(0.34,1.56,0.64,1); }
  @keyframes wc-cup-pop { 0%{transform:scale(0.5);} 60%{transform:scale(1.18);} 100%{transform:scale(1);} }
  .wc-overlay { animation: wc-fade-in 0.18s ease; }
  @keyframes wc-fade-in { from{opacity:0;} to{opacity:1;} }
  .wc-sheet-center { animation: wc-pop-in 0.22s cubic-bezier(0.32,0.72,0,1); }
  @keyframes wc-pop-in { from{opacity:0; transform:scale(0.94);} to{opacity:1; transform:scale(1);} }
  .wc-toast { animation: wc-toast-in 0.25s cubic-bezier(0.32,0.72,0,1); }
  @keyframes wc-toast-in { from{opacity:0; transform:translate(-50%,8px);} to{opacity:1; transform:translate(-50%,0);} }
  .wc-cta:active:not(:disabled), .wc-m-btn:active, .wc-quick:active:not(:disabled) { transform: scale(0.985); }
  @media (hover: hover) {
    .wc-cta-green:not(:disabled):hover { background: ${T.cGreenDk}; }
    .wc-quick:not(:disabled):hover { background: ${T.cGreen}; color: #FFFFFF; }
  }
`

type Popup = null | "ad" | "reward" | "duplicate" | "error"

const CupIcon = ({ filled, size = 52 }: { filled: boolean; size?: number }) => {
    const w = size,
        h = size * 1.17
    return (
        <svg width={w} height={h} viewBox="0 0 24 28" fill="none">
            <path
                d="M5 5 H19 L17.5 24 Q17.3 26 15.3 26 H8.7 Q6.7 26 6.5 24 Z"
                stroke={filled ? T.cGreen : "#DCDFE3"}
                fill="transparent"
                strokeWidth="1.8"
                strokeLinejoin="round"
            />
            {filled && (
                <path
                    d="M6 11 H18 L17.1 23.5 Q17 25 15.4 25 H8.6 Q7 25 6.9 23.5 Z"
                    fill={T.cGreen}
                />
            )}
        </svg>
    )
}

export default function WaterChallenge({
    previewState = "empty",
    mockApiResult = "success",
}: {
    previewState?: string
    mockApiResult?: ClaimResult
}) {
    const [count, setCount] = useState(0)
    const [claimed, setClaimed] = useState(false)
    const [popup, setPopup] = useState<Popup>(null)
    const [adProgress, setAdProgress] = useState(0)
    const [adTimer, setAdTimer] = useState(5)
    const [toast, setToast] = useState("")
    const [ready, setReady] = useState(false)
    const [cooldownLeft, setCooldownLeft] = useState(0)
    const cdIvRef = useRef<ReturnType<typeof setInterval> | null>(null)
    const adIvRef = useRef<ReturnType<typeof setInterval> | null>(null)
    const toastTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
    const lastAddRef = useRef<number>(0)
    const busyRef = useRef(false)
    const [userIdStr] = useState<string>(() => getQueryParam("userId"))
    const claimPromiseRef = useRef<Promise<ClaimResult> | null>(null)

    useEffect(() => {
        if (!userIdStr) {
            setReady(true)
            return
        }
        let cancelled = false
        const load = (attempt: number) => {
            rpc("get_water_status", { p_user_id: userIdStr })
                .then((r) => {
                    if (cancelled) return
                    if (r && r.ok) {
                        setCount(typeof r.cups === "number" ? r.cups : 0)
                        setClaimed(!!r.claimed)
                        if (
                            !r.claimed &&
                            typeof r.cups === "number" &&
                            r.cups < GOAL &&
                            typeof r.cooldown_remain_sec === "number" &&
                            r.cooldown_remain_sec > 0
                        ) {
                            startCooldown(r.cooldown_remain_sec)
                        }
                    }
                    setReady(true)
                })
                .catch(() => {
                    if (cancelled) return
                    if (attempt < 2) {
                        setTimeout(() => {
                            if (!cancelled) load(attempt + 1)
                        }, 1200)
                    } else {
                        setReady(true)
                    }
                })
        }
        load(0)
        return () => {
            cancelled = true
        }
    }, [userIdStr])

    useEffect(() => {
        if (userIdStr) return
        if (adIvRef.current) {
            clearInterval(adIvRef.current)
            adIvRef.current = null
        }
        if (toastTimerRef.current) {
            clearTimeout(toastTimerRef.current)
            toastTimerRef.current = null
        }
        setPopup(null)
        setToast("")
        lastAddRef.current = 0
        if (cdIvRef.current) {
            clearInterval(cdIvRef.current)
            cdIvRef.current = null
        }
        setCooldownLeft(0)
        setAdProgress(0)
        setAdTimer(5)
        switch (previewState) {
            case "empty":
                setCount(0)
                setClaimed(false)
                break
            case "partial":
                setCount(3)
                setClaimed(false)
                break
            case "almost_full":
                setCount(7)
                setClaimed(false)
                break
            case "all_full":
                setCount(8)
                setClaimed(false)
                break
            case "ad_playing":
                setCount(8)
                setClaimed(false)
                setAdProgress(60)
                setAdTimer(2)
                setPopup("ad")
                break
            case "reward":
                setCount(8)
                setClaimed(false)
                setPopup("reward")
                break
            case "claimed":
                setCount(8)
                setClaimed(true)
                break
            default:
                setCount(0)
                setClaimed(false)
        }
    }, [previewState])

    useEffect(() => {
        return () => {
            if (adIvRef.current) clearInterval(adIvRef.current)
            if (toastTimerRef.current) clearTimeout(toastTimerRef.current)
            if (cdIvRef.current) clearInterval(cdIvRef.current)
        }
    }, [])

    const allFull = count >= GOAL
    const pct = Math.min(Math.round((count / GOAL) * 100), 100)

    const showToast = (msg: string) => {
        setToast(msg)
        if (toastTimerRef.current) clearTimeout(toastTimerRef.current)
        toastTimerRef.current = setTimeout(() => setToast(""), 2200)
    }
    const syncStatus = () => {
        if (!userIdStr) return
        rpc("get_water_status", { p_user_id: userIdStr })
            .then((r) => {
                if (r && r.ok) {
                    setCount(typeof r.cups === "number" ? r.cups : 0)
                    setClaimed(!!r.claimed)
                }
            })
            .catch(() => {})
    }
    const cooldownToast = (remainSec: number) => {
        const remain = Math.max(1, Math.round(remainSec))
        showToast(
            remain >= 60
                ? "다음 물 등록은 1분 뒤부터 가능해요!"
                : `다음 물 등록은 ${remain}초 뒤부터 가능해요!`
        )
    }
    const startCooldown = (sec: number) => {
        const s = Math.max(0, Math.round(sec))
        if (cdIvRef.current) {
            clearInterval(cdIvRef.current)
            cdIvRef.current = null
        }
        setCooldownLeft(s)
        if (s <= 0) return
        cdIvRef.current = setInterval(() => {
            setCooldownLeft((prev) => {
                if (prev <= 1) {
                    if (cdIvRef.current) {
                        clearInterval(cdIvRef.current)
                        cdIvRef.current = null
                    }
                    return 0
                }
                return prev - 1
            })
        }, 1000)
    }
    const notifyParent = () => {
        try {
            const payload = JSON.stringify({
                type: "MISSION_COMPLETED",
                mission: "DAILY_WATER_CUPS",
                rewardCoins: 10,
                userId: userIdStr,
            })
            const w = window as any
            if (w.ReactNativeWebView && w.ReactNativeWebView.postMessage) {
                w.ReactNativeWebView.postMessage(payload)
            } else if (window.parent && window.parent !== window) {
                window.parent.postMessage(payload, "*")
            }
        } catch {}
    }
    const handleAddCup = (n: number = 1) => {
        if (claimed || allFull || !ready || cooldownLeft > 0) return
        if (!userIdStr) {
            const now = Date.now()
            if (lastAddRef.current && now - lastAddRef.current < COOLDOWN_MS) {
                cooldownToast((COOLDOWN_MS - (now - lastAddRef.current)) / 1000)
                return
            }
            lastAddRef.current = now
            setCount((c) => {
                const next = Math.min(c + n, GOAL)
                if (next < GOAL) startCooldown(COOLDOWN_MS / 1000)
                return next
            })
            return
        }
        if (busyRef.current) return
        busyRef.current = true
        rpc("add_water_cups", { p_user_id: userIdStr, p_cups: n })
            .then((r) => {
                if (r && r.ok) {
                    const cups = typeof r.cups === "number" ? r.cups : 0
                    setCount(cups)
                    if (cups < GOAL) {
                        startCooldown(
                            typeof r.cooldown_remain_sec === "number"
                                ? r.cooldown_remain_sec
                                : 60
                        )
                    }
                } else if (r && r.error === "cooldown") {
                    const remain =
                        typeof r.cooldown_remain_sec === "number"
                            ? r.cooldown_remain_sec
                            : 60
                    cooldownToast(remain)
                    startCooldown(remain)
                    if (typeof r.cups === "number") setCount(r.cups)
                } else if (
                    r &&
                    (r.error === "already_full" ||
                        r.error === "already_claimed")
                ) {
                    if (typeof r.cups === "number") setCount(r.cups)
                    if (r.claimed) setClaimed(true)
                } else {
                    showToast("일시적인 오류가 발생했어요")
                    syncStatus()
                }
            })
            .catch(() => {
                showToast("일시적인 오류가 발생했어요")
                syncStatus()
            })
            .finally(() => {
                busyRef.current = false
            })
    }
    const handleStartReward = () => {
        setPopup("ad")
        setAdProgress(0)
        setAdTimer(5)
        const mock = userIdStr ? undefined : mockApiResult
        claimPromiseRef.current = claimReward(userIdStr, mock)
        let e = 0
        adIvRef.current = setInterval(() => {
            e++
            setAdProgress(Math.min((e / 5) * 100, 100))
            setAdTimer(Math.max(5 - e, 0))
            if (e >= 5) {
                if (adIvRef.current) {
                    clearInterval(adIvRef.current)
                    adIvRef.current = null
                }
                claimPromiseRef.current?.then((result) => {
                    setTimeout(() => {
                        if (result === "success") setPopup("reward")
                        else if (result === "duplicate") setPopup("duplicate")
                        else setPopup("error")
                    }, 200)
                })
            }
        }, 1000)
    }
    const handleAckDuplicate = () => {
        setClaimed(true)
        setPopup(null)
    }
    const handleRetry = () => {
        setPopup("ad")
        setAdTimer(0)
        setAdProgress(100)
        const mock = userIdStr ? undefined : mockApiResult
        claimPromiseRef.current = claimReward(userIdStr, mock)
        claimPromiseRef.current.then((result) => {
            setTimeout(() => {
                if (result === "success") setPopup("reward")
                else if (result === "duplicate") setPopup("duplicate")
                else setPopup("error")
            }, 400)
        })
    }
    const handleClaim = () => {
        setClaimed(true)
        setPopup(null)
        if (userIdStr) notifyParent()
    }
    const handleClosePopup = () => {
        if (adIvRef.current) {
            clearInterval(adIvRef.current)
            adIvRef.current = null
        }
        setPopup(null)
    }

    let mainCta: React.ReactNode = null
    let showQuickChips = false
    if (!ready) {
        mainCta = (
            <button disabled style={ctaSoftSty} className="wc-cta">
                불러오는 중...
            </button>
        )
    } else if (claimed) {
        mainCta = (
            <button disabled style={ctaSoftSty} className="wc-cta">
                오늘 챌린지 완료 ✓
            </button>
        )
    } else if (allFull) {
        mainCta = (
            <button
                onClick={handleStartReward}
                style={ctaGreenSty}
                className="wc-cta wc-cta-green"
            >
                <span>모두 마셨어요</span>
                <span style={chipSty} className="wc-cta-chip">
                    광고
                </span>
            </button>
        )
    } else if (cooldownLeft > 0) {
        showQuickChips = true
        mainCta = (
            <button disabled style={ctaCooldownSty} className="wc-cta">
                다음 물 등록까지 {cooldownLeft}초
            </button>
        )
    } else {
        showQuickChips = true
        mainCta = (
            <button
                onClick={() => handleAddCup(1)}
                style={ctaGreenSty}
                className="wc-cta wc-cta-green"
            >
                + 한 잔 마셨어요
            </button>
        )
    }

    return (
        <div className="wc" style={rootSty}>
            <style>{CSS}</style>
            <div style={hdrSty}>
                <span style={hdrTitleSty}>물 마시기</span>
            </div>
            <div className="wc-body" style={bodySty}>
                <div style={heroSty}>
                    <div style={heroTopSty}>
                        <span style={heroLabelSty}>오늘의 물 섭취</span>
                    </div>
                    <div style={heroProgressRowSty}>
                        <span style={heroNumSty}>{count}</span>
                        <span style={heroTotalSty}>/ {GOAL} 컵</span>
                        <span style={heroPctSty}>{pct}%</span>
                    </div>
                    <div style={heroBarSty}>
                        <div style={{ ...heroBarFillSty, width: `${pct}%` }} />
                    </div>
                </div>
                <div style={sectionHeadSty}>
                    <span style={sectionTitleSty}>오늘의 8잔</span>
                    <span style={sectionMetaSty}>
                        {count} / {GOAL}
                    </span>
                </div>
                <div style={cupsGridCardSty}>
                    <div style={cupsGridSty}>
                        {Array.from({ length: GOAL }).map((_, i) => {
                            const filled = i < count
                            return (
                                <div key={i} style={cupCellSty}>
                                    <div
                                        key={filled ? `${i}-on` : `${i}-off`}
                                        className={filled ? "wc-cup-fill" : ""}
                                        style={{ display: "flex" }}
                                    >
                                        <CupIcon filled={filled} size={50} />
                                    </div>
                                </div>
                            )
                        })}
                    </div>
                </div>
            </div>
            <div style={ctaWrapSty}>
                {showQuickChips && (
                    <div style={quickRowSty}>
                        <button
                            className="wc-quick"
                            disabled={cooldownLeft > 0}
                            onClick={() => handleAddCup(2)}
                            style={{
                                ...quickChipSty,
                                ...(cooldownLeft > 0 ? quickChipOffSty : {}),
                            }}
                        >
                            +2잔
                        </button>
                        <button
                            className="wc-quick"
                            disabled={cooldownLeft > 0}
                            onClick={() => handleAddCup(3)}
                            style={{
                                ...quickChipSty,
                                ...(cooldownLeft > 0 ? quickChipOffSty : {}),
                            }}
                        >
                            +3잔
                        </button>
                        <button
                            className="wc-quick"
                            disabled={cooldownLeft > 0}
                            onClick={() => handleAddCup(4)}
                            style={{
                                ...quickChipSty,
                                ...(cooldownLeft > 0 ? quickChipOffSty : {}),
                            }}
                        >
                            +4잔
                        </button>
                    </div>
                )}
                {mainCta}
            </div>

            {toast && (
                <div className="wc-toast" style={toastSty}>
                    {toast}
                </div>
            )}

            {popup === "ad" && (
                <div className="wc-overlay" style={overlayCenterSty}>
                    <div
                        className="wc-sheet-center"
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
                            {adTimer > 0 ? "광고 시청 중" : "보상 처리 중..."}
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
                            {adTimer > 0
                                ? `${adTimer}초 남음`
                                : "잠시만 기다려주세요"}
                        </div>
                    </div>
                </div>
            )}
            {popup === "reward" && (
                <div className="wc-overlay" style={overlayCenterSty}>
                    <div className="wc-sheet-center" style={sheetCenterSty}>
                        <div style={scoreEmojiSty}>🎉</div>
                        <div style={mTitleSty}>보상이 적립되었어요</div>
                        <div style={mSubSty}>내일도 8잔 도전해요</div>
                        <button
                            onClick={handleClaim}
                            className="wc-m-btn"
                            style={mBtnGreenSty}
                        >
                            확인
                        </button>
                    </div>
                </div>
            )}
            {popup === "duplicate" && (
                <div className="wc-overlay" style={overlayCenterSty}>
                    <div className="wc-sheet-center" style={sheetCenterSty}>
                        <div style={scoreEmojiSty}>✅</div>
                        <div style={mTitleSty}>이미 받은 보상이에요</div>
                        <div style={mSubSty}>하루에 한 번만 받을 수 있어요</div>
                        <button
                            onClick={handleAckDuplicate}
                            className="wc-m-btn"
                            style={mBtnGreenSty}
                        >
                            확인
                        </button>
                    </div>
                </div>
            )}
            {popup === "error" && (
                <div className="wc-overlay" style={overlayCenterSty}>
                    <div className="wc-sheet-center" style={sheetCenterSty}>
                        <div style={scoreEmojiSty}>⚠️</div>
                        <div style={mTitleSty}>일시적인 오류가 발생했어요</div>
                        <div style={mSubSty}>잠시 후 다시 시도해주세요</div>
                        <button
                            onClick={handleRetry}
                            className="wc-m-btn"
                            style={mBtnGreenSty}
                        >
                            다시 시도
                        </button>
                        <button
                            onClick={handleClosePopup}
                            className="wc-m-btn"
                            style={mBtnGhostSty}
                        >
                            닫기
                        </button>
                    </div>
                </div>
            )}
        </div>
    )
}

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
    touchAction: "manipulation",
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
const cupsGridCardSty: React.CSSProperties = {
    background: T.cBg,
    border: `1px solid ${T.cBorder}`,
    borderRadius: T.rCard,
    padding: "22px 18px",
    marginBottom: 12,
}
const cupsGridSty: React.CSSProperties = {
    display: "grid",
    gridTemplateColumns: "repeat(4, 1fr)",
    gap: 14,
    justifyItems: "center",
    alignItems: "center",
}
const cupCellSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
}
const ctaWrapSty: React.CSSProperties = {
    padding: "14px 20px 24px",
    flexShrink: 0,
}
const quickRowSty: React.CSSProperties = {
    display: "flex",
    gap: 8,
    marginBottom: 10,
}
const quickChipSty: React.CSSProperties = {
    flex: 1,
    padding: "11px 0",
    background: T.cGreenBg,
    border: "none",
    borderRadius: T.rBtn,
    fontSize: T.tSm,
    fontWeight: T.wBold,
    color: T.cGreenDk,
    cursor: "pointer",
    fontFamily: FONT,
    letterSpacing: -0.1,
    transition: "background 0.15s, color 0.15s, transform 0.1s",
}
const quickChipOffSty: React.CSSProperties = {
    background: T.cDivider,
    color: T.cText4,
    cursor: "default",
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
    transition: "transform 0.1s, opacity 0.15s, background 0.15s",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 10,
}
const ctaGreenSty: React.CSSProperties = {
    ...ctaBaseSty,
    background: T.cGreen,
    color: "#FFFFFF",
}
const ctaSoftSty: React.CSSProperties = {
    ...ctaBaseSty,
    background: T.cDivider,
    color: T.cText,
    cursor: "default",
}
const ctaCooldownSty: React.CSSProperties = {
    ...ctaBaseSty,
    background: T.cDivider,
    color: T.cText3,
    fontWeight: T.wLabel,
    cursor: "default",
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
const toastSty: React.CSSProperties = {
    position: "absolute",
    left: "50%",
    bottom: 118,
    transform: "translateX(-50%)",
    background: "rgba(25,31,40,0.92)",
    color: "#FFFFFF",
    fontSize: T.tSm,
    fontWeight: T.wLabel,
    padding: "11px 18px",
    borderRadius: T.rPill,
    whiteSpace: "nowrap",
    zIndex: 90,
    letterSpacing: -0.2,
    pointerEvents: "none",
}
const overlayCenterSty: React.CSSProperties = {
    position: "absolute",
    inset: 0,
    background: "rgba(20,23,28,0.5)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 20,
    zIndex: 100,
}
const sheetCenterSty: React.CSSProperties = {
    width: "100%",
    maxWidth: 320,
    background: T.cBg,
    borderRadius: T.rCard,
    padding: "28px 22px 22px",
    textAlign: "center",
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
    color: "#FFFFFF",
}
const mBtnGhostSty: React.CSSProperties = {
    ...mBtnBaseSty,
    background: "transparent",
    color: T.cText3,
    padding: "12px 0",
    fontWeight: T.wLabel,
    fontSize: T.tMd,
    marginTop: 4,
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
const scoreEmojiSty: React.CSSProperties = {
    fontSize: T.t3xl,
    marginBottom: 8,
    lineHeight: 1,
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

addPropertyControls(WaterChallenge, {
    previewState: {
        type: ControlType.Enum,
        title: "미리보기",
        options: [
            "empty",
            "partial",
            "almost_full",
            "all_full",
            "ad_playing",
            "reward",
            "claimed",
        ],
        optionTitles: [
            "시작 전 (0/8)",
            "마시는 중 (3/8)",
            "거의 다 (7/8)",
            "모두 마심 (8/8)",
            "광고 재생 중",
            "캐시 적립 팝업",
            "오늘 챌린지 완료",
        ],
        defaultValue: "empty",
    },
    mockApiResult: {
        type: ControlType.Enum,
        title: "userId 없을 때 모의 응답",
        options: ["success", "duplicate", "error"],
        optionTitles: ["성공 (204)", "중복 (409)", "오류 (500)"],
        defaultValue: "success",
    },
})
