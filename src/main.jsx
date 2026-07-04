import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.jsx';

const style = document.createElement('style');
style.textContent = `
  html, body { margin: 0; padding: 0; height: 100%; background: #DCE6E3; }
  #root { min-height: 100vh; display: flex; }
  * { -webkit-font-smoothing: antialiased; }
`;
document.head.appendChild(style);

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
