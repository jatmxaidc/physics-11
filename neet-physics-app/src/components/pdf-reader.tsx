"use client";

import {
  ArrowLeft,
  Expand,
  Home,
  Maximize2,
  Minimize2,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import Link from "next/link";
import { useCallback, useEffect, useMemo, useState } from "react";

type PdfReaderProps = {
  materialId: string;
  title: string;
  kind: string;
  backHref: string;
  homeHref: string;
  initialUrl: string;
  expiresAt: number;
};

export default function PdfReader({
  materialId,
  title,
  kind,
  backHref,
  homeHref,
  initialUrl,
  expiresAt,
}: PdfReaderProps) {
  const [url, setUrl] = useState(initialUrl);
  const [zoom, setZoom] = useState(100);
  const [fitWidth, setFitWidth] = useState(true);
  const [fullscreen, setFullscreen] = useState(false);
  const [error, setError] = useState("");

  const refreshUrl = useCallback(async () => {
    try {
      const response = await fetch(`/api/materials/${materialId}?format=json`, {
        cache: "no-store",
      });
      if (!response.ok) throw new Error("Could not refresh the secure reading link.");
      const payload = (await response.json()) as { url: string };
      setUrl(payload.url);
      setError("");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unable to open the notes.");
    }
  }, [materialId]);

  useEffect(() => {
    if (url.startsWith("/api/materials/")) return;
    const remaining = expiresAt - Date.now();
    const refreshIn = Math.max(30_000, remaining - 60_000);
    const timer = window.setTimeout(() => {
      void refreshUrl();
    }, refreshIn);
    return () => window.clearTimeout(timer);
  }, [expiresAt, refreshUrl, url]);

  useEffect(() => {
    function onFullscreenChange() {
      setFullscreen(Boolean(document.fullscreenElement));
    }
    document.addEventListener("fullscreenchange", onFullscreenChange);
    return () => document.removeEventListener("fullscreenchange", onFullscreenChange);
  }, []);

  const frameSrc = useMemo(() => {
    const separator = url.includes("?") ? "&" : "?";
    if (url.startsWith("/api/materials/")) {
      return `${url}#toolbar=1&navpanes=0&scrollbar=1&view=FitH`;
    }
    return `${url}${separator}toolbar=1#view=FitH`;
  }, [url]);

  const scale = fitWidth ? 1 : zoom / 100;

  async function toggleFullscreen() {
    const shell = document.getElementById("pdf-reader-shell");
    if (!shell) return;
    if (!document.fullscreenElement) await shell.requestFullscreen();
    else await document.exitFullscreen();
  }

  return (
    <div id="pdf-reader-shell" className={`pdf-reader ${fullscreen ? "is-fullscreen" : ""}`}>
      <header className="pdf-reader-bar">
        <div className="pdf-reader-identity">
          <Link className="icon-button" href={backHref} aria-label="Back to classes">
            <ArrowLeft size={17} />
          </Link>
          <Link className="icon-button" href={homeHref} aria-label="Workspace home">
            <Home size={17} />
          </Link>
          <div>
            <span className="small-label">{kind.replace(/_/g, " ")} · focus reader</span>
            <h1>{title}</h1>
          </div>
        </div>
        <div className="pdf-reader-tools">
          <button
            type="button"
            className="icon-button"
            onClick={() => {
              setFitWidth(false);
              setZoom((value) => Math.max(75, value - 25));
            }}
            aria-label="Zoom out"
          >
            <ZoomOut size={17} />
          </button>
          <button
            type="button"
            className="pdf-zoom-chip"
            onClick={() => {
              setFitWidth(true);
              setZoom(100);
            }}
          >
            {fitWidth ? "Fit width" : `${zoom}%`}
          </button>
          <button
            type="button"
            className="icon-button"
            onClick={() => {
              setFitWidth(false);
              setZoom((value) => Math.min(200, value + 25));
            }}
            aria-label="Zoom in"
          >
            <ZoomIn size={17} />
          </button>
          <button
            type="button"
            className="icon-button"
            onClick={() => {
              setFitWidth(false);
              setZoom(150);
            }}
            aria-label="Comfortable zoom"
            title="150% zoom"
          >
            <Maximize2 size={17} />
          </button>
          <button type="button" className="icon-button" onClick={toggleFullscreen} aria-label="Fullscreen">
            {fullscreen ? <Minimize2 size={17} /> : <Expand size={17} />}
          </button>
          <Link className="button secondary pdf-back-classes" href={backHref}>
            My classes
          </Link>
        </div>
      </header>

      <div className="pdf-reader-stage">
        <div
          className={`pdf-reader-page ${fitWidth ? "is-fit" : "is-zoomed"}`}
          style={
            fitWidth
              ? undefined
              : {
                  width: `min(980px, 100%)`,
                  height: `calc((100vh - 120px) * ${scale})`,
                }
          }
        >
          {error ? (
            <div className="pdf-reader-error">
              <p>{error}</p>
              <button className="button primary" type="button" onClick={() => void refreshUrl()}>
                Retry
              </button>
            </div>
          ) : (
            <div
              className="pdf-reader-scale"
              style={
                fitWidth
                  ? undefined
                  : {
                      width: `${100 / scale}%`,
                      height: `${100 / scale}%`,
                      transform: `scale(${scale})`,
                      transformOrigin: "top left",
                    }
              }
            >
              <iframe title={title} src={frameSrc} className="pdf-reader-frame" />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
