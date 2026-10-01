import { useState } from 'react';
import { SalesApp } from '../sales/SalesApp';
import { SplitApp } from './SplitApp';

export function App() {
  const [view, setView] = useState<'sales' | 'split'>('sales');
  return view === 'sales'
    ? <SalesApp onOpenSplit={() => setView('split')} />
    : <SplitApp onBack={() => setView('sales')} />;
}
