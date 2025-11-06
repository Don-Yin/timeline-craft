'use client';
import { useEffect, useRef, useState } from "react";

type ResizableColumnsProps = {
    left: React.ReactNode;
    right: React.ReactNode;
    defaultLeftFraction?: number; // 0..1
    minLeftFraction?: number; // 0..1
    maxLeftFraction?: number; // 0..1
    storageKey?: string;
    className?: string;
};

export default function ResizableColumns({
    left,
    right,
    defaultLeftFraction = 0.32,
    minLeftFraction = 0.2,
    maxLeftFraction = 0.8,
    storageKey = "resizable.columns.fraction",
    className = "",
}: ResizableColumnsProps) {
    const containerRef = useRef<HTMLDivElement | null>(null);
    const isDraggingRef = useRef(false);
    const [leftFraction, setLeftFraction] = useState<number>(defaultLeftFraction);

    // load saved fraction
    useEffect(() => {
        try {
            const v = localStorage.getItem(storageKey);
            if (v) {
                const parsed = Number(v);
                if (!Number.isNaN(parsed)) {
                    setLeftFraction(clamp(parsed, minLeftFraction, maxLeftFraction));
                }
            }
        } catch {
            // ignore
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    // save fraction
    useEffect(() => {
        try {
            localStorage.setItem(storageKey, String(leftFraction));
        } catch {
            // ignore
        }
    }, [leftFraction, storageKey]);

    // pointer events are attached on demand to avoid stale listeners

    function startDrag(e: React.PointerEvent<HTMLDivElement>) {
        isDraggingRef.current = true;
        (e.currentTarget as HTMLDivElement).setPointerCapture(e.pointerId);
        const onPointerMove = (ev: PointerEvent) => {
            if (!isDraggingRef.current || !containerRef.current) return;
            const rect = containerRef.current.getBoundingClientRect();
            const frac = (ev.clientX - rect.left) / rect.width;
            setLeftFraction(clamp(frac, minLeftFraction, maxLeftFraction));
            ev.preventDefault();
        };
        const onPointerUp = () => {
            isDraggingRef.current = false;
            document.body.style.cursor = "";
            document.removeEventListener("pointermove", onPointerMove);
            document.removeEventListener("pointerup", onPointerUp);
        };
        document.addEventListener("pointermove", onPointerMove);
        document.addEventListener("pointerup", onPointerUp);
        document.body.style.cursor = "col-resize";
        e.preventDefault();
    }

    const leftPct = Math.round(leftFraction * 1000) / 10; // 1 decimal
    const rightPct = Math.round((1 - leftFraction) * 1000) / 10;

    return (
        <div
            ref={containerRef}
            className={["grid items-stretch gap-5", className].join(" ")}
            style={{
                gridTemplateColumns: `${leftPct}% 8px ${rightPct}%`,
            }}
        >
            <div className="min-w-0">{left}</div>
            <div
                role="separator"
                aria-orientation="vertical"
                title="drag to resize"
                onPointerDown={startDrag}
                className="cursor-col-resize rounded bg-sidebar-border hover:bg-sidebar-accent"
            />
            <div className="min-w-0">{right}</div>
        </div>
    );
}

function clamp(v: number, min: number, max: number): number {
    if (v < min) return min;
    if (v > max) return max;
    return v;
}


