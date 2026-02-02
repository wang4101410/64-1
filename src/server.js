import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3001;
const DATA_FILE = path.join(__dirname, '..', 'server-data.json');

// 中間件
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// 確保數據文件存在
if (!fs.existsSync(DATA_FILE)) {
  fs.writeFileSync(DATA_FILE, JSON.stringify({}));
}

// 獲取數據的路由
app.get('/api/data/:userId', (req, res) => {
  try {
    const { userId } = req.params;
    const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    const userData = data[userId] || null;
    res.json({ success: true, data: userData });
  } catch (error) {
    console.error('Error reading data:', error);
    res.status(500).json({ success: false, error: 'Failed to read data' });
  }
});

// 保存數據的路由
app.post('/api/data/:userId', (req, res) => {
  try {
    const { userId } = req.params;
    const newData = req.body;

    // 讀取現有數據
    let data = {};
    if (fs.existsSync(DATA_FILE)) {
      data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    }

    // 更新用戶數據
    data[userId] = {
      ...newData,
      lastUpdated: new Date().toISOString()
    };

    // 寫入文件
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));

    res.json({ success: true, message: 'Data saved successfully' });
  } catch (error) {
    console.error('Error saving data:', error);
    res.status(500).json({ success: false, error: 'Failed to save data' });
  }
});

// 健康檢查路由
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});