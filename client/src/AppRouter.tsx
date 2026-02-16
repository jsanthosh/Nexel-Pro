import React from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import App from './App';
import DocumentList from './components/DocumentList';

export default function AppRouter() {
  return (
    <Routes>
      <Route path="/" element={<DocumentList />} />
      <Route path="/doc/:id" element={<App />} />
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}
