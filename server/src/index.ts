import dotenv from 'dotenv';
import path from 'path';
dotenv.config({ path: path.resolve(__dirname, '../../.env') });
import express from 'express';
import cors from 'cors';
import chatRouter from './routes/chat';
import documentsRouter from './routes/documents';

const app = express();
const PORT = 3001;

app.use(cors({ origin: 'http://localhost:3000' }));
app.use(express.json({ limit: '50mb' }));

app.get('/health', (_req, res) => res.json({ status: 'ok' }));
app.use('/api', chatRouter);
app.use('/api', documentsRouter);

// Production: serve client build and handle client-side routing
if (process.env.NODE_ENV === 'production') {
  app.use(express.static(path.resolve(__dirname, '../../client/build')));
  app.get('*', (_req, res) => {
    res.sendFile(path.resolve(__dirname, '../../client/build/index.html'));
  });
}

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
