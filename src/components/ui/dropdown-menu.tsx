import * as React from "react"
import { cn } from "@/lib/utils"

// Simplified dropdown menu without radix dependency
interface DropdownMenuProps { children: React.ReactNode }
interface DropdownMenuTriggerProps { asChild?: boolean; children: React.ReactNode }
interface DropdownMenuContentProps extends React.HTMLAttributes<HTMLDivElement> { align?: string }

const DropdownContext = React.createContext<{ open: boolean; setOpen: (v: boolean) => void }>({ open: false, setOpen: () => {} });

export function DropdownMenu({ children }: DropdownMenuProps) {
  const [open, setOpen] = React.useState(false);
  return (
    <DropdownContext.Provider value={{ open, setOpen }}>
      <div className="relative">{children}</div>
    </DropdownContext.Provider>
  );
}

export function DropdownMenuTrigger({ asChild, children }: DropdownMenuTriggerProps) {
  const { open, setOpen } = React.useContext(DropdownContext);
  const handleClick = () => setOpen(!open);

  if (asChild && React.isValidElement(children)) {
    return React.cloneElement(children as React.ReactElement<{ onClick: () => void }>, { onClick: handleClick });
  }
  return <button onClick={handleClick}>{children}</button>;
}

export function DropdownMenuContent({ className, children, ...props }: DropdownMenuContentProps) {
  const { open, setOpen } = React.useContext(DropdownContext);
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
    <div ref={ref} className={cn("absolute right-0 bottom-full mb-1 z-50 min-w-[8rem] rounded-md border bg-popover p-1 text-popover-foreground shadow-md", className)} {...props}>
      {children}
    </div>
  );
}

export function DropdownMenuItem({ className, children, onClick, ...props }: React.HTMLAttributes<HTMLDivElement> & { onClick?: () => void }) {
  const { setOpen } = React.useContext(DropdownContext);
  return (
    <div
      className={cn("relative flex cursor-pointer select-none items-center gap-2 rounded-sm px-2 py-1.5 text-sm outline-none hover:bg-accent hover:text-accent-foreground", className)}
      onClick={() => { onClick?.(); setOpen(false); }}
      {...props}
    >
      {children}
    </div>
  );
}

export function DropdownMenuSeparator({ className, ...props }: React.HTMLAttributes<HTMLDivElement>) {
  return <div className={cn("-mx-1 my-1 h-px bg-border", className)} {...props} />;
}
