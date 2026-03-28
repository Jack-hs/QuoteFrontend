import React, { useState } from 'react';

const ExcelUploader = () => {
    const [file, setFile] = useState(null);
    const [schoolName, setSchoolName] = useState('');
    const [fileType, setFileType] = useState('tuition');
    const [status, setStatus] = useState('');

    const handleFileChange = (e) => {
        setFile(e.target.files[0]);

        if (selectedFile) {
            setFile(selectedFile);

            // 嘗試從檔名推斷學校名稱
            const baseName = selectedFile.name.replace(/\.[^/.]+$/, "");
            const schoolName = baseName
            .replace(/_tuition$/i, "")
            .replace(/_localfee$/i, "");

            if (schoolName !== baseName) {
            setSchoolName(schoolName.toUpperCase());
            }
        }
    };

    const handleUpload = async (e) => {
        e.preventDefault();
        if (!file || !schoolName) {
            setStatus('請選擇檔案並輸入學校名稱！');
            return;
        }

        const formData = new FormData();
        formData.append('file', file);
        formData.append('schoolName', schoolName);
        formData.append('fileType', fileType);

        try {
            setStatus('上傳中並轉換格式...');
            
            // 替換成您的實際 API 網址，例如 http://localhost:5000/api/quote/upload-excel
            //const apiUrl = 'https://localhost:7080';
            const apiUrl = import.meta.env.VITE_API_URL || 'https://localhost:7080';
            const response = await fetch(apiUrl + '/api/quote/upload-excel', {
                method: 'POST',
                body: formData,
            });

            if (response.ok) {
                const data = await response.json();
                setStatus(`成功！產生檔案: ${data.fileName}`);
                // --- 關鍵：重置 input.value，讓下次可以再選同一個檔 ---
                e.target.reset();        // 重置整個 form
                setFile(null);
                setSchoolName('');       // 依需求要不要清除
            } else {
                const errText = await response.text();
                setStatus(`錯誤: ${errText}`);
            }
        } catch (error) {
            console.error(error);
            setStatus('連線或轉換發生異常');
        }
    };

    return (
        <div style={{ padding: '20px', maxWidth: '500px', margin: '0 auto' }}>
            <h2>更新學校 Excel 資料</h2>
            <form onSubmit={handleUpload} style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
                
                <div>
                    <label>學校名稱 (例如 PHILINTER)：</label><br />
                    <input 
                        type="text" 
                        value={schoolName} 
                        onChange={(e) => setSchoolName(e.target.value.toUpperCase())} 
                        placeholder="請輸入學校名稱"
                        required
                        style={{ width: '100%', padding: '8px' }}
                    />
                </div>

                <div>
                    <label>資料類型：</label><br />
                    <select 
                        value={fileType} 
                        onChange={(e) => setFileType(e.target.value)}
                        style={{ width: '100%', padding: '8px' }}
                    >
                        <option value="tuition">學費與住宿費 (Tuition)</option>
                        <option value="localfee">當地學雜費 (Local Fee)</option>
                    </select>
                </div>

                <div>
                    <label>上傳 Excel 檔案 (.xlsx)：</label><br />
                    <input 
                        type="file" 
                        accept=".xlsx, .xls" 
                        onChange={handleFileChange}
                        required
                    />
                    {/* 在這顯示選了幾個檔案 */}
                    {file ? (
                        <div style={{ color: '#007BFF', fontSize: '14px', marginTop: '5px' }}>
                        已選擇檔案：{file.name}
                        </div>
                    ) : (
                        <div style={{ color: '#999', fontSize: '14px', marginTop: '5px' }}>
                        目前尚未選擇任何檔案，請選擇 .xlsx 或 .xls 檔案
                        </div>
                    )}
                </div>

                <button type="submit" style={{ padding: '10px', backgroundColor: '#007BFF', color: '#FFF', border: 'none', cursor: 'pointer' }}>
                    上傳並轉為 JSON
                </button>
            </form>

            {status && (
                <div style={{ marginTop: '20px', padding: '10px', backgroundColor: '#f0f0f0', borderRadius: '5px' }}>
                    <strong>狀態：</strong> {status}
                </div>
            )}
        </div>
    );
};

export default ExcelUploader;
