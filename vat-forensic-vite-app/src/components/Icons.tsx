import type { PropsWithChildren } from "react";

type IconProps = { className?: string };

function IconBase({ className, children }: PropsWithChildren<IconProps>) {
  return (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className} aria-hidden="true">
      {children}
    </svg>
  );
}

export function AlertTriangleIcon({ className }: IconProps) {
  return <IconBase className={className}><path d="M12 3 2.8 19a1.4 1.4 0 0 0 1.2 2h16a1.4 1.4 0 0 0 1.2-2L12 3Z" /><path d="M12 9v5" /><path d="M12 17h.01" /></IconBase>;
}
export function CheckCircleIcon({ className }: IconProps) {
  return <IconBase className={className}><circle cx="12" cy="12" r="9" /><path d="m8.5 12.5 2.2 2.2 4.8-5.2" /></IconBase>;
}
export function FileIcon({ className }: IconProps) {
  return <IconBase className={className}><path d="M14 2H7a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7Z" /><path d="M14 2v5h5" /><path d="M8 11h8" /><path d="M8 15h8" /><path d="M10 9v10" /><path d="M14 9v10" /></IconBase>;
}
export function FolderUpIcon({ className }: IconProps) {
  return <IconBase className={className}><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z" /><path d="M12 15V9" /><path d="m9.5 11.5 2.5-2.5 2.5 2.5" /></IconBase>;
}
export function LockIcon({ className }: IconProps) {
  return <IconBase className={className}><rect x="5" y="11" width="14" height="10" rx="2" /><path d="M8 11V8a4 4 0 1 1 8 0v3" /></IconBase>;
}
export function RefreshIcon({ className }: IconProps) {
  return <IconBase className={className}><path d="M21 12a9 9 0 0 1-15.5 6.4" /><path d="M3 12A9 9 0 0 1 18.5 5.6" /><path d="M3 16v2h2" /><path d="M21 8V6h-2" /></IconBase>;
}
export function ShieldIcon({ className }: IconProps) {
  return <IconBase className={className}><path d="M12 3 5 6v5c0 5 3.4 8.4 7 10 3.6-1.6 7-5 7-10V6l-7-3Z" /><path d="M12 8v5" /><path d="M12 16h.01" /></IconBase>;
}
export function UploadIcon({ className }: IconProps) {
  return <IconBase className={className}><path d="M12 16V6" /><path d="m8.5 9.5 3.5-3.5 3.5 3.5" /><path d="M4 18v1a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-1" /></IconBase>;
}