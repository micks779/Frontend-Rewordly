/* global Office */
import React from 'react';
import ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import App from './App';
import './index.css';

// Initialize Fluent UI icons
initializeIcons();

// Initialize Office.js
Office.onReady(() => {
  ReactDOM.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
    document.getElementById('root')
  );
});

// If you want to start measuring performance in your app, pass a function
// to log results (for example: reportWebVitals(console.log))
// or send to an analytics endpoint. 