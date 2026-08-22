let toastTimer: ReturnType<typeof setTimeout> | undefined;

export type ToastType = 'info' | 'success' | 'error';

export function toast(message: string, typeOrDuration: ToastType | number = 'info', duration = 3000): void {
  const element = document.getElementById('toast');
  if (!element) return;

  const type: ToastType = typeof typeOrDuration === 'number' ? 'info' : typeOrDuration;
  const displayDuration = typeof typeOrDuration === 'number' ? typeOrDuration : duration;
  element.textContent = message;
  element.dataset.type = type;
  element.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    element.classList.remove('show');
  }, displayDuration);
}