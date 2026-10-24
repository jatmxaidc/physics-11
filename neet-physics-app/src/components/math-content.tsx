"use client";

import katex from "katex";
import "katex/dist/katex.min.css";
import { Fragment, type ReactNode } from "react";

type FigParams = Record<string, string>;

function parseParams(raw: string): FigParams {
  const params: FigParams = {};
  for (const part of raw.split("|")) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) {
      params[trimmed] = "1";
      continue;
    }
    params[trimmed.slice(0, eq).trim()] = trimmed.slice(eq + 1).trim();
  }
  return params;
}

function renderKatex(tex: string, displayMode: boolean): string {
  try {
    return katex.renderToString(tex, {
      throwOnError: false,
      displayMode,
      strict: "ignore",
      output: "html",
    });
  } catch {
    return tex;
  }
}

function RightTriangleFig({ params }: { params: FigParams }) {
  const opp = params.opp ?? "a";
  const adj = params.adj ?? "b";
  const hyp = params.hyp ?? "c";
  const angle = params.theta ?? params.angle ?? "θ";
  return (
    <svg viewBox="0 0 280 200" role="img" aria-label="Right triangle diagram">
      <line x1="40" y1="160" x2="220" y2="160" stroke="currentColor" strokeWidth="2" />
      <line x1="40" y1="160" x2="40" y2="40" stroke="currentColor" strokeWidth="2" />
      <line x1="40" y1="40" x2="220" y2="160" stroke="currentColor" strokeWidth="2" />
      <rect x="40" y="140" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="1.5" />
      <path d="M70 160 A22 22 0 0 0 58 142" fill="none" stroke="#2470e8" strokeWidth="1.75" />
      <text x="78" y="152" fill="#2470e8" fontSize="14" fontFamily="serif">{angle}</text>
      <text x="118" y="178" fill="currentColor" fontSize="13" fontFamily="serif">{adj}</text>
      <text x="12" y="105" fill="currentColor" fontSize="13" fontFamily="serif">{opp}</text>
      <text x="140" y="85" fill="currentColor" fontSize="13" fontFamily="serif">{hyp}</text>
    </svg>
  );
}

function TowerFig({ params }: { params: FigParams }) {
  const h = params.h ?? "h";
  const angle = params.angle ?? "45°";
  return (
    <svg viewBox="0 0 300 210" role="img" aria-label="Angle of elevation diagram">
      <line x1="30" y1="170" x2="270" y2="170" stroke="currentColor" strokeWidth="2" />
      <rect x="200" y="50" width="28" height="120" fill="#dbe7f8" stroke="#2470e8" strokeWidth="1.75" />
      <line x1="60" y1="170" x2="214" y2="50" stroke="currentColor" strokeWidth="2" />
      <circle cx="60" cy="170" r="3.5" fill="#2470e8" />
      <path d="M85 170 A28 28 0 0 0 78 148" fill="none" stroke="#2470e8" strokeWidth="1.75" />
      <text x="92" y="158" fill="#2470e8" fontSize="13" fontFamily="serif">{angle}</text>
      <text x="236" y="115" fill="currentColor" fontSize="13" fontFamily="serif">{h}</text>
      <text x="120" y="192" fill="currentColor" fontSize="12" fontFamily="serif">
        distance = ?
      </text>
    </svg>
  );
}

function ParabolaFig() {
  return (
    <svg viewBox="0 0 260 200" role="img" aria-label="Upward opening parabola">
      <line x1="30" y1="170" x2="230" y2="170" stroke="#8aa0b8" strokeWidth="1.5" />
      <line x1="130" y1="20" x2="130" y2="180" stroke="#8aa0b8" strokeWidth="1.5" />
      <path
        d="M50 160 Q130 20 210 160"
        fill="none"
        stroke="#2470e8"
        strokeWidth="2.5"
      />
      <text x="200" y="40" fill="#2470e8" fontSize="14" fontFamily="serif">
        y = x²
      </text>
      <text x="220" y="185" fill="#8aa0b8" fontSize="12">
        x
      </text>
      <text x="138" y="30" fill="#8aa0b8" fontSize="12">
        y
      </text>
    </svg>
  );
}

function CircleFig({ params }: { params: FigParams }) {
  const cx = params.cx ?? "3";
  const cy = params.cy ?? "−2";
  return (
    <svg viewBox="0 0 260 200" role="img" aria-label="Circle with centre marked">
      <line x1="20" y1="100" x2="240" y2="100" stroke="#8aa0b8" strokeWidth="1.25" />
      <line x1="130" y1="20" x2="130" y2="180" stroke="#8aa0b8" strokeWidth="1.25" />
      <circle cx="155" cy="120" r="48" fill="rgba(36,112,232,.08)" stroke="#2470e8" strokeWidth="2" />
      <circle cx="155" cy="120" r="3" fill="#2470e8" />
      <text x="165" y="118" fill="#2470e8" fontSize="13" fontFamily="serif">
        ({cx}, {cy})
      </text>
    </svg>
  );
}

function AreaUnderFig({ params }: { params: FigParams }) {
  const label = params.y ?? "y = x";
  const a = params.a ?? "0";
  const b = params.b ?? "4";
  return (
    <svg viewBox="0 0 280 200" role="img" aria-label="Area under a line">
      <line x1="40" y1="160" x2="250" y2="160" stroke="#8aa0b8" strokeWidth="1.5" />
      <line x1="40" y1="20" x2="40" y2="160" stroke="#8aa0b8" strokeWidth="1.5" />
      <polygon points="40,160 200,40 200,160" fill="rgba(36,112,232,.18)" stroke="none" />
      <line x1="40" y1="160" x2="200" y2="40" stroke="#2470e8" strokeWidth="2.5" />
      <line x1="200" y1="40" x2="200" y2="160" stroke="#2470e8" strokeWidth="1" strokeDasharray="4 3" />
      <text x="210" y="35" fill="#2470e8" fontSize="13" fontFamily="serif">
        {label}
      </text>
      <text x="36" y="178" fill="#43566d" fontSize="12">
        {a}
      </text>
      <text x="194" y="178" fill="#43566d" fontSize="12">
        {b}
      </text>
    </svg>
  );
}

function AreaBetweenFig() {
  return (
    <svg viewBox="0 0 280 200" role="img" aria-label="Area between y equals x and y equals x squared">
      <line x1="40" y1="160" x2="250" y2="160" stroke="#8aa0b8" strokeWidth="1.5" />
      <line x1="40" y1="20" x2="40" y2="160" stroke="#8aa0b8" strokeWidth="1.5" />
      <path d="M40 160 L160 40 L160 160 Z" fill="rgba(36,112,232,.12)" />
      <path d="M40 160 Q100 160 160 40" fill="rgba(255,255,255,.55)" stroke="none" />
      <path d="M40 160 Q100 160 160 40" fill="none" stroke="#e07a3a" strokeWidth="2.25" />
      <line x1="40" y1="160" x2="160" y2="40" stroke="#2470e8" strokeWidth="2.25" />
      <text x="175" y="55" fill="#2470e8" fontSize="12" fontFamily="serif">
        y = x
      </text>
      <text x="175" y="95" fill="#e07a3a" fontSize="12" fontFamily="serif">
        y = x²
      </text>
      <text x="36" y="178" fill="#43566d" fontSize="12">
        0
      </text>
      <text x="154" y="178" fill="#43566d" fontSize="12">
        1
      </text>
    </svg>
  );
}

function RectangleFig({ params }: { params: FigParams }) {
  const side = params.side ?? "10 m";
  return (
    <svg viewBox="0 0 260 180" role="img" aria-label="Square of maximum area">
      <rect x="60" y="35" width="120" height="120" fill="rgba(36,112,232,.1)" stroke="#2470e8" strokeWidth="2" />
      <text x="95" y="100" fill="#2470e8" fontSize="14" fontFamily="serif">
        {side}
      </text>
      <text x="100" y="168" fill="#43566d" fontSize="12">
        perimeter fixed
      </text>
    </svg>
  );
}

function LineSlopeFig({ params }: { params: FigParams }) {
  const p1 = params.p1 ?? "(1, 2)";
  const p2 = params.p2 ?? "(3, 8)";
  return (
    <svg viewBox="0 0 280 200" role="img" aria-label="Line joining two points">
      <line x1="30" y1="170" x2="250" y2="170" stroke="#8aa0b8" strokeWidth="1.5" />
      <line x1="50" y1="20" x2="50" y2="180" stroke="#8aa0b8" strokeWidth="1.5" />
      <line x1="90" y1="140" x2="190" y2="50" stroke="#2470e8" strokeWidth="2.5" />
      <circle cx="90" cy="140" r="4" fill="#2470e8" />
      <circle cx="190" cy="50" r="4" fill="#2470e8" />
      <text x="98" y="155" fill="#43566d" fontSize="12" fontFamily="serif">
        {p1}
      </text>
      <text x="198" y="48" fill="#43566d" fontSize="12" fontFamily="serif">
        {p2}
      </text>
    </svg>
  );
}

function Figure({ kind, params }: { kind: string; params: FigParams }) {
  switch (kind) {
    case "rightTriangle":
      return <RightTriangleFig params={params} />;
    case "tower":
      return <TowerFig params={params} />;
    case "parabola":
      return <ParabolaFig />;
    case "circle":
      return <CircleFig params={params} />;
    case "areaUnder":
      return <AreaUnderFig params={params} />;
    case "areaBetween":
      return <AreaBetweenFig />;
    case "rectangle":
      return <RectangleFig params={params} />;
    case "lineSlope":
      return <LineSlopeFig params={params} />;
    default:
      return null;
  }
}

function tokenize(content: string): Array<
  | { type: "text"; value: string }
  | { type: "math"; value: string; display: boolean }
  | { type: "figure"; kind: string; params: FigParams }
> {
  const tokens: Array<
    | { type: "text"; value: string }
    | { type: "math"; value: string; display: boolean }
    | { type: "figure"; kind: string; params: FigParams }
  > = [];
  const pattern = /(\$\$[\s\S]*?\$\$|\$(?:\\\$|[^$])+?\$|\{\{fig:[^}]+\}\})/g;
  let last = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(content)) !== null) {
    if (match.index > last) {
      tokens.push({ type: "text", value: content.slice(last, match.index) });
    }
    const token = match[0];
    if (token.startsWith("{{fig:")) {
      const inner = token.slice(6, -2);
      const [kind, ...rest] = inner.split("|");
      tokens.push({ type: "figure", kind: kind.trim(), params: parseParams(rest.join("|")) });
    } else if (token.startsWith("$$") && token.endsWith("$$")) {
      tokens.push({ type: "math", value: token.slice(2, -2).trim(), display: true });
    } else {
      tokens.push({ type: "math", value: token.slice(1, -1).trim(), display: false });
    }
    last = match.index + token.length;
  }
  if (last < content.length) {
    tokens.push({ type: "text", value: content.slice(last) });
  }
  return tokens;
}

function renderText(value: string): ReactNode[] {
  return value.split("\n").map((line, index, arr) => (
    <Fragment key={`t-${index}`}>
      {line}
      {index < arr.length - 1 ? <br /> : null}
    </Fragment>
  ));
}

export function MathContent({
  content,
  className,
  as: Tag = "div",
}: {
  content: string;
  className?: string;
  as?: "div" | "span" | "p" | "h2" | "h3";
}) {
  if (!content) return null;
  const tokens = tokenize(content);
  const nodes = tokens.map((token, index) => {
    if (token.type === "text") {
      return <Fragment key={index}>{renderText(token.value)}</Fragment>;
    }
    if (token.type === "math") {
      return (
        <span
          key={index}
          className={token.display ? "math-block" : "math-inline"}
          dangerouslySetInnerHTML={{ __html: renderKatex(token.value, token.display) }}
        />
      );
    }
    return (
      <div className="question-figure" key={index}>
        <Figure kind={token.kind} params={token.params} />
      </div>
    );
  });

  return <Tag className={className ? `math-content ${className}` : "math-content"}>{nodes}</Tag>;
}
