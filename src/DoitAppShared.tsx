import { useEffect } from 'react';
import type { MouseEvent } from 'react';
import DoitAppThai from './DoitAppThai';
import type { Session } from './lib/dataset';

const AUTH_KEY = 'doit.auth.thai.v1';

function sharedSession(): Session {
  return {
    name: 'ใช้งานร่วมกัน',
    org: 'ทุกคนใช้ชุดข้อมูลเดียวกัน',
    token: 'shared-workspace',
    createdAt: 'shared',
  };
}

function ensureSharedSession(): void {
  if (typeof window === 'undefined') return;
  const next = sharedSession();
  try {
    const raw = window.localStorage.getItem(AUTH_KEY);
    const parsed = raw ? JSON.parse(raw) as Partial<Session> | null : null;
    if (!parsed || parsed.token !== next.token) {
      window.localStorage.setItem(AUTH_KEY, JSON.stringify(next));
    }
  } catch {
    window.localStorage.setItem(AUTH_KEY, JSON.stringify(next));
  }
}

function relabelLogoutButton(): void {
  if (typeof document === 'undefined') return;
  document.querySelectorAll('button').forEach((button) => {
    const text = button.textContent?.trim();
    if (text === 'ออกจากระบบ') button.textContent = 'เริ่มไฟล์ใหม่';
  });
}

ensureSharedSession();

export default function DoitAppShared() {
  useEffect(() => {
    ensureSharedSession();
    relabelLogoutButton();
    const observer = new MutationObserver(relabelLogoutButton);
    observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    return () => observer.disconnect();
  }, []);

  function handleClickCapture(event: MouseEvent<HTMLDivElement>) {
    const target = event.target as HTMLElement | null;
    const button = target?.closest('button');
    const text = button?.textContent?.trim();
    if (text === 'ออกจากระบบ' || text === 'เริ่มไฟล์ใหม่') {
      event.preventDefault();
      event.stopPropagation();
      ensureSharedSession();
      window.location.reload();
    }
  }

  return (
    <div onClickCapture={handleClickCapture}>
      <div style={{
        position: 'sticky',
        top: 0,
        zIndex: 20,
        padding: '12px 18px',
        background: '#0284c7',
        color: '#ffffff',
        fontFamily: 'Tahoma, Noto Sans Thai, system-ui, sans-serif',
        fontSize: 18,
        fontWeight: 900,
        lineHeight: 1.5,
        boxShadow: '0 8px 20px rgba(2, 132, 199, 0.24)',
      }}>
        โหมดใช้งานร่วมกัน: ไม่ต้องล็อกอิน ทุกคนใช้ข้อมูลชุดเดียวกันในเครื่องนี้ ถ้าเปลี่ยนเครื่องให้กด “ส่งออกฐานข้อมูล” ก่อน
      </div>
      <DoitAppThai />
    </div>
  );
}
