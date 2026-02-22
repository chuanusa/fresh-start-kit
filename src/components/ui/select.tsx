import * as React from "react"
import { cn } from "@/lib/utils"

// Simple select components
interface SelectProps { value: string; onValueChange: (value: string) => void; children: React.ReactNode }

const SelectContext = React.createContext<{ value: string; onValueChange: (v: string) => void; open: boolean; setOpen: (v: boolean) => void }>({
  value: '', onValueChange: () => {}, open: false, setOpen: () => {}
});

export function Select({ value, onValueChange, children }: SelectProps) {
  const [open, setOpen] = React.useState(false);
  return (
    <SelectContext.Provider value={{ value, onValueChange, open, setOpen }}>
      <div className="relative">{children}</div>
    </SelectContext.Provider>
  );
}

export function SelectTrigger({ className, children }: React.HTMLAttributes<HTMLButtonElement>) {
  const { open, setOpen } = React.useContext(SelectContext);
  return (
    <button
      onClick={() => setOpen(!open)}
      className={cn("flex h-10 items-center justify-between rounded-md border border-input bg-background px-3 py-2 text-sm ring-offset-background focus:outline-none focus:ring-2 focus:ring-ring", className)}
    >
      {children}
      <svg className="h-4 w-4 opacity-50" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" /></svg>
    </button>
  );
}

export function SelectValue() {
  const { value } = React.useContext(SelectContext);
  return <span>{value}</span>;
}

export function SelectContent({ children }: { children: React.ReactNode }) {
  const { open, setOpen } = React.useContext(SelectContext);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open, setOpen]);

  if (!open) return null;
  return (
    <div ref={ref} className="absolute top-full mt-1 z-50 w-full min-w-[8rem] rounded-md border bg-popover p-1 text-popover-foreground shadow-md max-h-60 overflow-y-auto">
      {children}
    </div>
  );
}

export function SelectItem({ value, children }: { value: string; children: React.ReactNode }) {
  const { value: selected, onValueChange, setOpen } = React.useContext(SelectContext);
  return (
    <div
      className={cn("relative flex cursor-pointer select-none items-center rounded-sm px-2 py-1.5 text-sm hover:bg-accent", value === selected && "bg-accent")}
      onClick={() => { onValueChange(value); setOpen(false); }}
    >
      {children}
    </div>
  );
}
