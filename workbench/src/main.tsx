import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { App } from './App';

initializeIcons();

const container = document.getElementById('root');
if (!container) throw new Error('Missing #root element');
createRoot(container).render(<App />);
