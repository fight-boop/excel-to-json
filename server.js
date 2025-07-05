const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const app = express();

// 允许跨域请求
app.use(cors());

// 文件上传设置
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Excel转JSON处理函数
function excelToJson(buffer) {
    try {
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const result = {};
        
        // 处理所有工作表
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 'A' });
            
            // 提取标题行
            const headers = {};
            if (jsonData.length > 0) {
                const firstRow = jsonData[0];
                Object.keys(firstRow).forEach(col => {
                    headers[col] = firstRow[col] || `Column_${col}`;
                });
            }
            
            // 处理数据行
            const data = [];
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                const rowData = {};
                
                Object.keys(headers).forEach(col => {
                    rowData[headers[col]] = row[col] || null;
                });
                
                data.push(rowData);
            }
            
            result[sheetName] = data;
        });
        
        return result;
    } catch (error) {
        throw new Error('Excel处理失败: ' + error.message);
    }
}

// 文件转换路由
app.post('/convert', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ 
            success: false, 
            message: '未上传文件' 
        });
    }
    
    try {
        const jsonData = excelToJson(req.file.buffer);
        res.json({ 
            success: true, 
            result: jsonData 
        });
    } catch (error) {
        res.status(500).json({ 
            success: false, 
            message: error.message 
        });
    }
});

// 启动服务器
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`服务器运行在 http://localhost:${PORT}`);
});