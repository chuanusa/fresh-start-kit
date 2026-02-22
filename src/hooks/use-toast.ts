import * as React from "react"

export interface ToastProps {
  title?: string;
  description?: string;
  variant?: 'default' | 'destructive';
}

interface ToastState {
  toasts: (ToastProps & { id: number })[];
}

const TOAST_LIMIT = 3;
let toastCount = 0;
const listeners: Array<(state: ToastState) => void> = [];
let memoryState: ToastState = { toasts: [] };

function dispatch(state: ToastState) {
  memoryState = state;
  listeners.forEach((listener) => listener(state));
}

function toast(props: ToastProps) {
  const id = toastCount++;
  const newToast = { ...props, id };
  dispatch({ toasts: [newToast, ...memoryState.toasts].slice(0, TOAST_LIMIT) });
  setTimeout(() => {
    dispatch({ toasts: memoryState.toasts.filter((t) => t.id !== id) });
  }, 4000);
}

function useToast() {
  const [state, setState] = React.useState<ToastState>(memoryState);

  React.useEffect(() => {
    listeners.push(setState);
    return () => {
      const index = listeners.indexOf(setState);
      if (index > -1) listeners.splice(index, 1);
    };
  }, []);

  return { ...state, toast };
}

export { useToast, toast };
