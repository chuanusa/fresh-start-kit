import { useToast } from "@/hooks/use-toast";
import { cn } from "@/lib/utils";
import { X } from "lucide-react";

export function Toaster() {
  const { toasts } = useToast();

  return (
    <div className="fixed top-4 right-4 z-50 flex flex-col gap-2 w-80">
      {toasts.map((t) => (
        <div
          key={t.id}
          className={cn(
            "rounded-lg border p-4 shadow-lg transition-all animate-in slide-in-from-top-2",
            t.variant === "destructive"
              ? "border-destructive/50 bg-destructive text-destructive-foreground"
              : "border-border bg-card text-card-foreground"
          )}
        >
          {t.title && <p className="text-sm font-semibold">{t.title}</p>}
          {t.description && <p className="text-sm opacity-90 mt-1">{t.description}</p>}
        </div>
      ))}
    </div>
  );
}
