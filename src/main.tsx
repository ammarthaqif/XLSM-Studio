import { createRoot } from 'react-dom/client';
import App from './App.tsx';
import './index.css';

// Unregister any obsolete service workers if present
if (typeof window !== 'undefined' && 'serviceWorker' in navigator) {
  navigator.serviceWorker.getRegistrations().then((registrations) => {
    for (const registration of registrations) {
      registration.unregister();
    }
  });
}

createRoot(document.getElementById('root')!).render(<App />);
