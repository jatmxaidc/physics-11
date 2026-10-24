"use client";

import {
  BarChart3,
  BookOpen,
  CalendarDays,
  Home,
  Layers,
  LogOut,
  Map,
  Menu,
  MessageSquareText,
  NotebookPen,
  PencilLine,
  RotateCcw,
  Settings,
  Upload,
  UserRound,
  X,
} from "lucide-react";
import Image from "next/image";
import { usePathname } from "next/navigation";
import { useState } from "react";
import { logout } from "@/app/actions/auth";
import type { Profile } from "@/lib/auth";

const roleNavigation = {
  student: [
    ["Overview", "/student", Home],
    ["My classes", "/student/classes", BookOpen],
    ["Lecture planner", "/student/planner", Map],
    ["My notes", "/student/notes", NotebookPen],
    ["Flashcards", "/student/flashcards", Layers],
    ["Daily quizzes", "/student/assessments", PencilLine],
    ["Weekly tests", "/student/tests", CalendarDays],
    ["48-hour retry", "/student/retry", RotateCcw],
    ["My progress", "/student/progress", BarChart3],
    ["Account", "/student/account", Settings],
  ],
  teacher: [
    ["Dashboard", "/teacher", Home],
    ["Classes & materials", "/teacher/classes", Upload],
    ["Lecture planner", "/teacher/planner", CalendarDays],
    ["Assessments", "/teacher/assessments", PencilLine],
    ["Attendance & review", "/teacher/review", MessageSquareText],
    ["Students & results", "/teacher/students", BarChart3],
  ],
  parent: [
    ["Weekly summary", "/parent", Home],
    ["Tests & attendance", "/parent/tests", CalendarDays],
    ["Learning progress", "/parent/progress", BarChart3],
  ],
} as const;

export default function ProductionShell({
  profile,
  children,
}: {
  profile: Profile;
  children: React.ReactNode;
}) {
  const pathname = usePathname();
  const [mobileOpen, setMobileOpen] = useState(false);
  const initials = profile.full_name
    .split(/\s+/)
    .map((part) => part[0])
    .slice(0, 2)
    .join("")
    .toUpperCase();

  return (
    <div className="app-shell">
      {mobileOpen && (
        <button className="mobile-overlay" aria-label="Close navigation" onClick={() => setMobileOpen(false)} />
      )}
      <aside className={`sidebar ${mobileOpen ? "mobile-open" : ""}`}>
        <div className="sidebar-top">
          <div className="brand">
            <span className="brand-mark">
              <Image src="/smxphysics-logo.png" alt="" width={36} height={36} priority />
            </span>
            <div>
              <strong>SMXPhysics</strong>
              <span>Excellence Programme</span>
            </div>
          </div>
          <button className="icon-button mobile-close" onClick={() => setMobileOpen(false)}>
            <X size={18} />
          </button>
        </div>
        <div className="programme-pill">
          <span className="live-dot" />
          <div>
            <strong>Class 11 Physics</strong>
            <span>CBSE + NEET 2026–27</span>
          </div>
        </div>
        <nav className="nav-list">
          <span className="nav-kicker">{profile.role} workspace</span>
          {roleNavigation[profile.role].map(([label, href, Icon]) => {
            const active =
              pathname === href || (href !== `/${profile.role}` && pathname.startsWith(`${href}/`));
            return (
              <a
                key={href}
                href={href}
                className={`nav-item ${active ? "active" : ""}`}
                onClick={() => setMobileOpen(false)}
              >
                <Icon size={19} />
                <span>{label}</span>
              </a>
            );
          })}
        </nav>
        <div className="sidebar-bottom">
          <div className="sidebar-user">
            <span className="avatar">{initials}</span>
            <div>
              <strong>{profile.full_name}</strong>
              <span>{profile.role} account</span>
            </div>
          </div>
          <form action={logout}>
            <button className="nav-item" type="submit">
              <LogOut size={18} />
              <span>Sign out</span>
            </button>
          </form>
        </div>
      </aside>
      <div className="app-main">
        <header className="topbar">
          <button className="icon-button menu-button" onClick={() => setMobileOpen(true)}>
            <Menu size={19} />
          </button>
          <div className="topbar-title">
            <span className="sync-status">Secure workspace</span>
          </div>
          <div className="topbar-actions">
            <div className="topbar-user">
              <span className={`role-icon ${profile.role}`}>
                <UserRound size={15} />
              </span>
              <div>
                <strong>{profile.full_name}</strong>
                <small>{profile.role}</small>
              </div>
            </div>
          </div>
        </header>
        <main className="page-content">{children}</main>
      </div>
    </div>
  );
}
