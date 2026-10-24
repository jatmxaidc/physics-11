"use client";

import { useState } from "react";
import { RotateCcw } from "lucide-react";

export type FlashcardItem = {
  id: string;
  front: string;
  back: string;
  meta: string;
};

export default function FlashcardDeck({ cards }: { cards: FlashcardItem[] }) {
  const [index, setIndex] = useState(0);
  const [flipped, setFlipped] = useState(false);
  const card = cards[index];

  if (!card) return null;

  function go(next: number) {
    setFlipped(false);
    setIndex((current) => {
      const length = cards.length;
      return (current + next + length) % length;
    });
  }

  return (
    <section className="flashcard-deck">
      <div className="flashcard-meta">
        <span className="small-label">{card.meta}</span>
        <span>
          {index + 1} / {cards.length}
        </span>
      </div>
      <button
        type="button"
        className={`flashcard ${flipped ? "flipped" : ""}`}
        onClick={() => setFlipped((value) => !value)}
        aria-label={flipped ? "Show prompt" : "Show answer"}
      >
        <div className="flashcard-face front">
          <span className="small-label">PROMPT</span>
          <p>{card.front}</p>
          <small>Tap to reveal</small>
        </div>
        <div className="flashcard-face back">
          <span className="small-label">YOUR NOTE</span>
          <p>{card.back}</p>
          <small>Tap to flip back</small>
        </div>
      </button>
      <div className="flashcard-controls">
        <button className="button secondary" type="button" onClick={() => go(-1)}>
          Previous
        </button>
        <button className="button secondary" type="button" onClick={() => setFlipped((value) => !value)}>
          <RotateCcw size={16} /> Flip
        </button>
        <button className="button primary" type="button" onClick={() => go(1)}>
          Next
        </button>
      </div>
    </section>
  );
}
