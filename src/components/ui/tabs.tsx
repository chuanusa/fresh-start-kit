import * as React from "react"
import { cn } from "@/lib/utils"

interface TabsProps { value: string; onValueChange: (value: string) => void; children: React.ReactNode; className?: string }
const TabsContext = React.createContext<{ value: string; onValueChange: (v: string) => void }>({ value: '', onValueChange: () => {} });

export function Tabs({ value, onValueChange, children, className }: TabsProps) {
  return (
    <TabsContext.Provider value={{ value, onValueChange }}>
      <div className={className}>{children}</div>
    </TabsContext.Provider>
  );
}

export function TabsList({ className, ...props }: React.HTMLAttributes<HTMLDivElement>) {
  return <div className={cn("inline-flex h-10 items-center justify-center rounded-md bg-muted p-1 text-muted-foreground", className)} {...props} />;
}

export function TabsTrigger({ value, className, children, ...props }: React.HTMLAttributes<HTMLButtonElement> & { value: string }) {
  const { value: selected, onValueChange } = React.useContext(TabsContext);
  return (
    <button
      className={cn(
        "inline-flex items-center justify-center whitespace-nowrap rounded-sm px-3 py-1.5 text-sm font-medium transition-all",
        value === selected ? "bg-background text-foreground shadow-sm" : "hover:bg-background/50"
      , className)}
      onClick={() => onValueChange(value)}
      {...props}
    >
      {children}
    </button>
  );
}

export function TabsContent({ value, className, ...props }: React.HTMLAttributes<HTMLDivElement> & { value: string }) {
  const { value: selected } = React.useContext(TabsContext);
  if (value !== selected) return null;
  return <div className={cn("mt-2", className)} {...props} />;
}
