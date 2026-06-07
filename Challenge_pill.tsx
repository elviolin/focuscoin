// VitaminChallenge.tsx — Framer 컴포넌트
// 영양제 등록 → 매일 탭으로 체크 → 모두 챙기면 완료
// 외곽 폰 목업 chrome 없음, 부모 frame 사이즈 = 실제 앱 화면
//
// 동작 모드:
//   - userId 있음 (실제): URL ?userId=X → Supabase RPC로 영양제 목록·체크 영속
//   - userId 없음 (mock): previewState로 Framer 캔버스 미리보기 (메모리만)
// 외부 적립 API(POST /mission-completed)는 Phase 2에서 추가 예정

import React, { useState, useEffect } from "react"
import { addPropertyControls, ControlType } from "framer"

/* ═══════════════════════════════════════════
   Supabase (FocusCoin 공용 프로젝트)
   ═══════════════════════════════════════════ */
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
   Sheet motion (덜컹 방지: React state + transition)
   ═══════════════════════════════════════════ */
const EASE_OUT = "cubic-bezier(0.22, 1, 0.36, 1)"
const SHEET_DUR = 280 // ms (v3: 360 → 280, 반응 빠름)
const FADE_DUR = 200 // ms
const TRANS_BOTTOM = `transform ${SHEET_DUR}ms ${EASE_OUT}, opacity ${FADE_DUR}ms ease`
const TRANS_CENTER = `transform ${SHEET_DUR - 80}ms ${EASE_OUT}, opacity ${FADE_DUR}ms ease`
const TRANS_OVERLAY = `opacity ${FADE_DUR}ms ease`

/* ═══════════════════════════════════════════
   Embedded CSS
   ═══════════════════════════════════════════ */
const CSS = `
  .vc *, .vc *::before, .vc *::after {
    box-sizing: border-box;
    -webkit-tap-highlight-color: transparent;
  }
  .vc button, .vc input {
    -webkit-appearance: none;
    appearance: none;
    font-family: inherit;
    color: inherit;
  }
  .vc button {
    -webkit-touch-callout: none;
    user-select: none;
    -webkit-user-select: none;
  }
  .vc button:focus, .vc input:focus { outline: none; }
  .vc button::-moz-focus-inner { border: 0; padding: 0; }
  .vc-body::-webkit-scrollbar { width: 0; }
  .vc-body { scrollbar-width: none; -webkit-overflow-scrolling: touch; }

  .vc-check-pop { animation: vc-check-pop 0.32s cubic-bezier(0.34, 1.56, 0.64, 1); }
  @keyframes vc-check-pop {
    0% { transform: scale(0.6); }
    60% { transform: scale(1.15); }
    100% { transform: scale(1); }
  }

  .vc-cta:active:not(:disabled),
  .vc-m-btn:active,
  .vc-row-btn:active:not(:disabled),
  .vc-chip:active {
    transform: scale(0.985);
  }

  .vc-row-btn:active:not(:disabled) { background: ${T.cCard}; }

  @media (hover: hover) {
    .vc-cta-green:not(:disabled):hover { background: ${T.cGreenDk}; }
  }

  .vc-input:focus { outline: none; border-color: ${T.cGreen}; }
  .vc-input::placeholder { color: ${T.cText4}; font-weight: 500; }
`

/* ═══════════════════════════════════════════
   Types
   ═══════════════════════════════════════════ */
type Vitamin = {
    id: string
    name: string
    dose?: string
    time?: string
}
type Popup = null | "add" | "reward"

/* ═══════════════════════════════════════════
   Defaults / Data (mock 모드 미리보기용)
   ═══════════════════════════════════════════ */
const DEFAULT_VITAMINS: Vitamin[] = [
    { id: "v1", name: "비타민 C", dose: "1000mg", time: "아침" },
    { id: "v2", name: "종합비타민", dose: "1정", time: "점심" },
    { id: "v3", name: "오메가3", dose: "1캡슐", time: "저녁" },
    { id: "v4", name: "유산균", dose: "1포", time: "자기 전" },
    { id: "v5", name: "마그네슘", dose: "1정", time: "자기 전" },
]
const TIME_OPTIONS = ["아침", "점심", "저녁", "자기 전"]

/* ═══════════════════════════════════════════
   Icons
   ═══════════════════════════════════════════ */
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
const PlusIcon = ({ size = 16 }: { size?: number }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
        <path
            d="M12 5V19M5 12H19"
            stroke="currentColor"
            strokeWidth="2.4"
            strokeLinecap="round"
        />
    </svg>
)
const MinusIcon = ({ size = 12 }: { size?: number }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
        <path
            d="M5 12H19"
            stroke="currentColor"
            strokeWidth="3"
            strokeLinecap="round"
        />
    </svg>
)
const CloseIcon = ({ size = 14 }: { size?: number }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
        <path
            d="M6 6L18 18M6 18L18 6"
            stroke="currentColor"
            strokeWidth="2.2"
            strokeLinecap="round"
        />
    </svg>
)

/* ═══════════════════════════════════════════
   Main Component
   ═══════════════════════════════════════════ */
export default function VitaminChallenge({
    previewState = "empty",
}: {
    previewState?: string
}) {
    // userId 한 번만 추출 (mount 시점)
    const [userIdStr] = useState<string>(() => getQueryParam("userId"))
    // mock 모드면 즉시 ready, 실제 모드면 RPC 응답 후 ready
    const [ready, setReady] = useState<boolean>(() => !getQueryParam("userId"))

    const [vitamins, setVitamins] = useState<Vitamin[]>([])
    const [taken, setTaken] = useState<Record<string, boolean>>({})
    const [claimed, setClaimed] = useState(false)
    const [popup, setPopup] = useState<Popup>(null)
    const [editing, setEditing] = useState(false)
    const [newName, setNewName] = useState("")
    const [newDose, setNewDose] = useState("")
    const [newTime, setNewTime] = useState("")
    // 시트 등장/사라짐
    const [sheetReady, setSheetReady] = useState(false)

    /* ─── 실제 모드: Supabase에서 초기 상태 로드 ─── */
    useEffect(() => {
        if (!userIdStr) return
        let cancelled = false
        rpc("get_pill_status", { p_user_id: userIdStr })
            .then((r) => {
                if (cancelled) return
                if (r && r.ok) {
                    const items: Vitamin[] = (r.items || []).map((it: any) => ({
                        id: it.id,
                        name: it.name,
                        dose: it.dose || undefined,
                        time: it.time_slot || undefined,
                    }))
                    setVitamins(items)
                    const takenMap: Record<string, boolean> = {}
                    ;(r.taken_ids || []).forEach((id: string) => {
                        takenMap[id] = true
                    })
                    setTaken(takenMap)
                    setClaimed(!!r.claimed)
                }
                setReady(true)
            })
            .catch(() => {
                if (!cancelled) setReady(true)
            })
        return () => {
            cancelled = true
        }
    }, [userIdStr])

    /* ─── mock 모드: previewState 분기 (userId 없을 때만) ─── */
    useEffect(() => {
        if (userIdStr) return
        setPopup(null)
        setEditing(false)
        setNewName("")
        setNewDose("")
        setNewTime("")

        const allTakenMap: Record<string, boolean> = {
            v1: true,
            v2: true,
            v3: true,
            v4: true,
            v5: true,
        }

        switch (previewState) {
            case "empty":
                setVitamins([])
                setTaken({})
                setClaimed(false)
                break
            case "registered":
                setVitamins(DEFAULT_VITAMINS)
                setTaken({})
                setClaimed(false)
                break
            case "partial":
                setVitamins(DEFAULT_VITAMINS)
                setTaken({ v1: true, v2: true, v3: true })
                setClaimed(false)
                break
            case "all_taken":
                setVitamins(DEFAULT_VITAMINS)
                setTaken(allTakenMap)
                setClaimed(false)
                break
            case "adding":
                setVitamins(DEFAULT_VITAMINS)
                setTaken({ v1: true, v2: true })
                setClaimed(false)
                setPopup("add")
                break
            case "edit_mode":
                setVitamins(DEFAULT_VITAMINS)
                setTaken({ v1: true, v2: true })
                setClaimed(false)
                setEditing(true)
                break
            case "reward":
                setVitamins(DEFAULT_VITAMINS)
                setTaken(allTakenMap)
                setClaimed(false)
                setPopup("reward")
                break
            case "claimed":
                setVitamins(DEFAULT_VITAMINS)
                setTaken(allTakenMap)
                setClaimed(true)
                break
            default:
                setVitamins([])
                setTaken({})
                setClaimed(false)
        }
    }, [previewState, userIdStr])

    /* ─── Sheet enter/exit motion ─── */
    useEffect(() => {
        if (popup === null) {
            setSheetReady(false)
            return
        }
        setSheetReady(false)
        const id = requestAnimationFrame(() => {
            requestAnimationFrame(() => setSheetReady(true))
        })
        return () => cancelAnimationFrame(id)
    }, [popup])

    /* ─── Derived ─── */
    const total = vitamins.length
    const takenCount = vitamins.filter((v) => taken[v.id]).length
    const allTaken = total > 0 && takenCount === total
    const pct = total > 0 ? Math.round((takenCount / total) * 100) : 0

    /* ─── Handlers (낙관적 업데이트 + RPC) ─── */
    const handleToggle = async (id: string) => {
        if (claimed || editing) return
        const wasTaken = !!taken[id]
        // 낙관적
        setTaken((prev) => ({ ...prev, [id]: !wasTaken }))

        if (!userIdStr) return

        try {
            const r = await rpc("set_pill_taken", {
                p_user_id: userIdStr,
                p_pill_id: id,
                p_taken: !wasTaken,
            })
            if (!r || !r.ok) {
                // 롤백
                setTaken((prev) => ({ ...prev, [id]: wasTaken }))
            }
        } catch {
            setTaken((prev) => ({ ...prev, [id]: wasTaken }))
        }
    }

    const handleAddOpen = () => {
        setNewName("")
        setNewDose("")
        setNewTime("")
        setPopup("add")
    }

    const handleSaveVitamin = async () => {
        const name = newName.trim()
        if (!name) return
        const dose = newDose.trim() || undefined
        const time = newTime || undefined

        if (!userIdStr) {
            // mock 모드
            const newVit: Vitamin = {
                id: `v${Date.now()}`,
                name,
                dose,
                time,
            }
            setVitamins((prev) => [...prev, newVit])
            setPopup(null)
            return
        }

        // 실제 모드: 모달 먼저 닫고 RPC 호출
        setPopup(null)
        try {
            const r = await rpc("add_pill_item", {
                p_user_id: userIdStr,
                p_name: name,
                p_dose: dose || null,
                p_time_slot: time || null,
            })
            if (r && r.ok && r.id) {
                setVitamins((prev) => [
                    ...prev,
                    { id: r.id, name, dose, time },
                ])
            }
        } catch {
            // 실패 시 무시 (Phase 2에서 에러 토스트 고려)
        }
    }

    const handleDeleteVitamin = async (id: string) => {
        // 낙관적
        setVitamins((prev) => prev.filter((v) => v.id !== id))
        setTaken((prev) => {
            const next = { ...prev }
            delete next[id]
            return next
        })

        if (!userIdStr) return

        try {
            await rpc("delete_pill_item", {
                p_user_id: userIdStr,
                p_pill_id: id,
            })
        } catch {
            // 무시
        }
    }

    const handleComplete = () => {
        setPopup("reward")
    }

    const handleClaim = async () => {
        // 낙관적
        setClaimed(true)
        setPopup(null)

        if (!userIdStr) return

        try {
            await rpc("claim_pill_complete", { p_user_id: userIdStr })
            // Phase 2: 외부 적립 API도 여기서 호출
        } catch {
            // 무시 (Phase 2에서 응답 분기)
        }
    }

    const handleClosePopup = () => {
        setPopup(null)
    }

    /* ─── CTA decision ─── */
    let ctaContent: React.ReactNode = null
    if (total === 0) {
        ctaContent = (
            <button
                onClick={handleAddOpen}
                style={ctaGreenSty}
                className="vc-cta vc-cta-green"
            >
                영양제 등록하기
            </button>
        )
    } else if (claimed) {
        ctaContent = (
            <button disabled style={ctaSoftSty} className="vc-cta">
                오늘 챌린지 완료 ✓
            </button>
        )
    } else if (editing) {
        ctaContent = (
            <button
                onClick={() => setEditing(false)}
                style={ctaDarkSty}
                className="vc-cta"
            >
                편집 완료
            </button>
        )
    } else if (allTaken) {
        ctaContent = (
            <button
                onClick={handleComplete}
                style={ctaGreenSty}
                className="vc-cta vc-cta-green"
            >
                다 먹었어요!
            </button>
        )
    } else {
        ctaContent = (
            <button disabled style={ctaDisabledSty} className="vc-cta">
                {total - takenCount}개 남았어요
            </button>
        )
    }

    /* ─── Sheet inline motion styles ─── */
    const overlayMotionSty: React.CSSProperties = {
        opacity: sheetReady ? 1 : 0,
        transition: TRANS_OVERLAY,
    }
    // v3: slide 크기 크게 줄임 (100% → 40px) + scale 미세 (0.94 → 0.97)
    // mount 1 frame 깜빡 보여도 시각적으로 거의 안 보이게
    const sheetBottomMotionSty: React.CSSProperties = {
        transform: sheetReady ? "translateY(0)" : "translateY(40px)",
        opacity: sheetReady ? 1 : 0,
        transition: TRANS_BOTTOM,
        willChange: "transform, opacity",
    }
    const sheetCenterMotionSty: React.CSSProperties = {
        transform: sheetReady ? "scale(1)" : "scale(0.97)",
        opacity: sheetReady ? 1 : 0,
        transition: TRANS_CENTER,
        willChange: "transform, opacity",
    }

    // 초기 로드 전(실제 모드) 빈 placeholder — Rules of Hooks: 모든 훅 선언 후 위치
    if (!ready) {
        return <div className="vc" style={rootSty} />
    }

    return (
        <div className="vc" style={rootSty}>
            <style>{CSS}</style>

            {/* Header */}
            <div style={hdrSty}>
                <span style={hdrTitleSty}>영양제 챙기기</span>
                {total > 0 && !claimed && (
                    <button
                        onClick={() => setEditing(!editing)}
                        style={hdrBtnSty}
                    >
                        {editing ? "완료" : "편집"}
                    </button>
                )}
            </div>

            {/* Body or Empty State */}
            {total === 0 ? (
                <div style={emptySty}>
                    <div style={emptyEmojiSty}>💊</div>
                    <div style={emptyTitleSty}>첫 영양제를 등록해보세요</div>
                    <div style={emptySubSty}>매일 챙기는 습관 만들기</div>
                </div>
            ) : (
                <div className="vc-body" style={bodySty}>
                    {/* Hero card */}
                    <div style={heroSty}>
                        <div style={heroTopSty}>
                            <span style={heroLabelSty}>오늘의 챙김</span>
                            {claimed && (
                                <span style={rewardChipDoneSty}>완료</span>
                            )}
                        </div>
                        <div style={heroProgressRowSty}>
                            <span style={heroNumSty}>{takenCount}</span>
                            <span style={heroTotalSty}>/ {total}</span>
                            <span style={heroPctSty}>{pct}%</span>
                        </div>
                        <div style={heroBarSty}>
                            <div
                                style={{ ...heroBarFillSty, width: `${pct}%` }}
                            />
                        </div>
                    </div>

                    {/* Section head */}
                    <div style={sectionHeadSty}>
                        <span style={sectionTitleSty}>오늘의 영양제</span>
                        <span style={sectionMetaSty}>
                            {takenCount} / {total}
                        </span>
                    </div>

                    {/* Vitamin list */}
                    <div style={listSty}>
                        {vitamins.map((v, i) => {
                            const done = !!taken[v.id]
                            const meta = [v.dose, v.time]
                                .filter(Boolean)
                                .join(" · ")
                            return (
                                <div
                                    key={v.id}
                                    style={{
                                        borderTop:
                                            i === 0
                                                ? "none"
                                                : `1px solid ${T.cDivider}`,
                                    }}
                                >
                                    {editing ? (
                                        <div style={rowInnerSty}>
                                            <button
                                                onClick={() =>
                                                    handleDeleteVitamin(v.id)
                                                }
                                                style={deleteSty}
                                            >
                                                <MinusIcon size={12} />
                                            </button>
                                            <div style={infoSty}>
                                                <div
                                                    style={{
                                                        ...nameSty,
                                                        color: T.cText,
                                                    }}
                                                >
                                                    {v.name}
                                                </div>
                                                <div
                                                    style={{
                                                        ...metaSty,
                                                        color: T.cText4,
                                                    }}
                                                >
                                                    {meta || "권장 시간 없음"}
                                                </div>
                                            </div>
                                        </div>
                                    ) : (
                                        <button
                                            className="vc-row-btn"
                                            onClick={() => handleToggle(v.id)}
                                            disabled={claimed}
                                            style={rowBtnSty}
                                        >
                                            <span
                                                className={
                                                    done ? "vc-check-pop" : ""
                                                }
                                                style={
                                                    done
                                                        ? checkOnSty
                                                        : checkOffSty
                                                }
                                            >
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
                                            <div style={infoSty}>
                                                <div
                                                    style={{
                                                        ...nameSty,
                                                        color: done
                                                            ? T.cText4
                                                            : T.cText,
                                                    }}
                                                >
                                                    {v.name}
                                                </div>
                                                <div
                                                    style={{
                                                        ...metaSty,
                                                        color: done
                                                            ? T.cText3
                                                            : T.cText4,
                                                    }}
                                                >
                                                    {meta || "오늘 챙기기"}
                                                </div>
                                            </div>
                                        </button>
                                    )}
                                </div>
                            )
                        })}

                        {/* Add row (edit mode only) */}
                        {editing && (
                            <button onClick={handleAddOpen} style={addRowSty}>
                                <span style={addIconWrapSty}>
                                    <PlusIcon size={16} />
                                </span>
                                <span>영양제 추가</span>
                            </button>
                        )}
                    </div>
                </div>
            )}

            {/* CTA */}
            <div style={ctaWrapSty}>{ctaContent}</div>

            {/* ─── Add Modal (bottom sheet) ─── */}
            {popup === "add" && (
                <div
                    style={{ ...overlayBottomSty, ...overlayMotionSty }}
                    onClick={handleClosePopup}
                >
                    <div
                        style={{ ...sheetBottomSty, ...sheetBottomMotionSty }}
                        onClick={(e) => e.stopPropagation()}
                    >
                        <div style={sheetHandleSty} />
                        <div style={modalHeadSty}>
                            <span style={modalTitleSty}>영양제 추가</span>
                            <button
                                onClick={handleClosePopup}
                                style={modalCloseSty}
                            >
                                <CloseIcon size={14} />
                            </button>
                        </div>

                        <div style={fieldGroupSty}>
                            <label style={fieldLabelSty}>이름</label>
                            <input
                                type="text"
                                value={newName}
                                onChange={(e) => setNewName(e.target.value)}
                                placeholder="예: 비타민 D"
                                style={inputSty}
                                className="vc-input"
                                maxLength={20}
                                autoFocus
                            />
                        </div>

                        <div style={fieldGroupSty}>
                            <label style={fieldLabelSty}>
                                복용량 <span style={fieldOptSty}>(선택)</span>
                            </label>
                            <input
                                type="text"
                                value={newDose}
                                onChange={(e) => setNewDose(e.target.value)}
                                placeholder="예: 1000mg, 1정"
                                style={inputSty}
                                className="vc-input"
                                maxLength={20}
                            />
                        </div>

                        <div style={{ ...fieldGroupSty, marginBottom: 24 }}>
                            <label style={fieldLabelSty}>
                                권장 시간{" "}
                                <span style={fieldOptSty}>(선택)</span>
                            </label>
                            <div style={timeChipsSty}>
                                {TIME_OPTIONS.map((t) => {
                                    const active = newTime === t
                                    return (
                                        <button
                                            key={t}
                                            className="vc-chip"
                                            onClick={() =>
                                                setNewTime(active ? "" : t)
                                            }
                                            style={{
                                                ...timeChipSty,
                                                ...(active
                                                    ? timeChipActiveSty
                                                    : {}),
                                            }}
                                        >
                                            {t}
                                        </button>
                                    )
                                })}
                            </div>
                        </div>

                        <button
                            onClick={handleSaveVitamin}
                            disabled={!newName.trim()}
                            className="vc-m-btn"
                            style={{
                                ...mBtnGreenSty,
                                opacity: newName.trim() ? 1 : 0.4,
                                cursor: newName.trim()
                                    ? "pointer"
                                    : "not-allowed",
                            }}
                        >
                            저장
                        </button>
                    </div>
                </div>
            )}

            {/* ─── Reward Popup ─── */}
            {popup === "reward" && (
                <div style={{ ...overlayCenterSty, ...overlayMotionSty }}>
                    <div
                        style={{ ...sheetCenterSty, ...sheetCenterMotionSty }}
                    >
                        <div style={scoreEmojiSty}>🎉</div>
                        <div style={mTitleSty}>잘 챙기셨어요!</div>
                        <div style={mSubSty}>내일도 잊지 말고 챙겨요</div>
                        <button
                            onClick={handleClaim}
                            className="vc-m-btn"
                            style={mBtnGreenSty}
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

const hdrSty: React.CSSProperties = {
    padding: "20px 20px 8px",
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
}
const hdrTitleSty: React.CSSProperties = {
    fontSize: T.tLg,
    fontWeight: T.wBold,
    letterSpacing: -0.3,
    color: T.cText,
}
const hdrBtnSty: React.CSSProperties = {
    background: "transparent",
    border: "none",
    color: T.cText2,
    fontSize: T.tSm,
    fontWeight: T.wLabel,
    padding: "6px 10px",
    cursor: "pointer",
    fontFamily: FONT,
    borderRadius: T.rBtn,
}

const bodySty: React.CSSProperties = {
    flex: 1,
    overflowY: "auto",
    padding: "8px 16px 0",
}

/* Empty state */
const emptySty: React.CSSProperties = {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    padding: "40px 24px",
    textAlign: "center",
}
const emptyEmojiSty: React.CSSProperties = {
    fontSize: 64,
    marginBottom: 22,
    lineHeight: 1,
}
const emptyTitleSty: React.CSSProperties = {
    fontSize: T.tXl,
    fontWeight: T.wBold,
    color: T.cText,
    marginBottom: 8,
    letterSpacing: -0.4,
}
const emptySubSty: React.CSSProperties = {
    fontSize: T.tMd,
    color: T.cText2,
    fontWeight: T.wBody,
}

/* Hero */
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
const rewardChipDoneSty: React.CSSProperties = {
    display: "inline-flex",
    alignItems: "center",
    padding: "4px 10px",
    borderRadius: T.rPill,
    background: T.cDivider,
    color: T.cText2,
    fontSize: T.tXs,
    fontWeight: T.wBold,
    letterSpacing: -0.1,
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

/* Section */
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

/* List */
const listSty: React.CSSProperties = {
    background: T.cBg,
    border: `1px solid ${T.cBorder}`,
    borderRadius: T.rCard,
    overflow: "hidden",
    marginBottom: 12,
}
const rowInnerSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: 13,
    padding: "14px 16px",
}
const rowBtnSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: 13,
    padding: "14px 16px",
    width: "100%",
    background: "transparent",
    border: "none",
    textAlign: "left",
    fontFamily: FONT,
    cursor: "pointer",
    transition: "background 0.12s, transform 0.1s",
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
const infoSty: React.CSSProperties = { flex: 1, minWidth: 0 }
const nameSty: React.CSSProperties = {
    fontSize: T.tLg,
    fontWeight: T.wLabel,
    letterSpacing: -0.3,
    lineHeight: 1.3,
}
const metaSty: React.CSSProperties = {
    fontSize: T.tSm,
    marginTop: 3,
    fontWeight: T.wBody,
}
const deleteSty: React.CSSProperties = {
    width: 24,
    height: 24,
    borderRadius: "50%",
    background: T.cRed,
    border: "none",
    color: "#FFF",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
    cursor: "pointer",
    fontFamily: FONT,
    transition: "transform 0.1s",
}
const addRowSty: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: 13,
    padding: "14px 16px",
    width: "100%",
    background: "transparent",
    border: "none",
    borderTop: `1px solid ${T.cDivider}`,
    textAlign: "left",
    fontFamily: FONT,
    cursor: "pointer",
    color: T.cGreen,
    fontSize: T.tMd,
    fontWeight: T.wLabel,
    transition: "transform 0.1s",
}
const addIconWrapSty: React.CSSProperties = {
    width: 24,
    height: 24,
    borderRadius: "50%",
    background: T.cGreenBg,
    color: T.cGreen,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
}

/* CTA */
const ctaWrapSty: React.CSSProperties = {
    padding: "14px 20px 24px",
    flexShrink: 0,
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

/* Overlay/Sheet */
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

/* Modal head (add) */
const modalHeadSty: React.CSSProperties = {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: 22,
}
const modalTitleSty: React.CSSProperties = {
    fontSize: T.tXl,
    fontWeight: T.wBold,
    color: T.cText,
    letterSpacing: -0.4,
}
const modalCloseSty: React.CSSProperties = {
    width: 32,
    height: 32,
    borderRadius: "50%",
    background: T.cDivider,
    border: "none",
    color: T.cText2,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    fontFamily: FONT,
}

/* Form */
const fieldGroupSty: React.CSSProperties = { marginBottom: 16 }
const fieldLabelSty: React.CSSProperties = {
    display: "block",
    fontSize: T.tSm,
    color: T.cText2,
    fontWeight: T.wLabel,
    marginBottom: 8,
}
const fieldOptSty: React.CSSProperties = {
    color: T.cText3,
    fontWeight: T.wBody,
    fontSize: T.tXs,
    marginLeft: 4,
}
const inputSty: React.CSSProperties = {
    width: "100%",
    padding: "14px 16px",
    fontSize: T.tMd,
    fontWeight: T.wLabel,
    color: T.cText,
    background: T.cCard,
    border: `1.5px solid transparent`,
    borderRadius: T.rBtn,
    fontFamily: FONT,
    outline: "none",
    transition: "border-color 0.15s",
    WebkitAppearance: "none",
}
const timeChipsSty: React.CSSProperties = {
    display: "flex",
    flexWrap: "wrap",
    gap: 8,
}
const timeChipSty: React.CSSProperties = {
    padding: "8px 14px",
    background: T.cCard,
    border: `1.5px solid transparent`,
    borderRadius: T.rPill,
    fontSize: T.tSm,
    fontWeight: T.wLabel,
    color: T.cText2,
    fontFamily: FONT,
    cursor: "pointer",
    transition: "all 0.15s, transform 0.1s",
}
const timeChipActiveSty: React.CSSProperties = {
    background: T.cGreenBg,
    borderColor: T.cGreen,
    color: T.cGreenDk,
    fontWeight: T.wBold,
}

/* Modal buttons */
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
    transition: "transform 0.1s, opacity 0.15s",
}
const mBtnGreenSty: React.CSSProperties = {
    ...mBtnBaseSty,
    background: T.cGreen,
    color: "#FFF",
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

/* ═══════════════════════════════════════════
   Property Controls (Framer 미리보기)
   ═══════════════════════════════════════════ */
addPropertyControls(VitaminChallenge, {
    previewState: {
        type: ControlType.Enum,
        title: "미리보기",
        options: [
            "empty",
            "registered",
            "partial",
            "all_taken",
            "adding",
            "edit_mode",
            "reward",
            "claimed",
        ],
        optionTitles: [
            "시작 전 (영양제 0개)",
            "등록됨 (0/5)",
            "일부 챙김 (3/5)",
            "모두 챙김 (5/5)",
            "추가 모달 열림",
            "편집 모드",
            "완료 팝업",
            "오늘 챌린지 완료",
        ],
        defaultValue: "empty",
    },
})
