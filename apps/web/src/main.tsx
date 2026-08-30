import React from 'react';
import { createRoot } from 'react-dom/client';
import { Provider } from 'jotai';
import App from './App';
import './styles.css';

// Точка входа приложения.
// Provider из Jotai подключает глобальное хранилище атомов (см. src/atoms.ts),
// доступное всем компонентам через хуки useAtom/useAtomValue/useSetAtom.
createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <Provider>
      <App />
    </Provider>
  </React.StrictMode>,
);
