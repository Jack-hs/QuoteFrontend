import React, { useState, useEffect } from 'react';
import axios from 'axios';

// 這裡接收一個 onRatesLoaded 屬性，用來把抓到的資料傳回給 App.jsx
export default function ExchangeRate({ onRatesLoaded }) {
  const [exchangeRates, setExchangeRates] = useState(null);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState('');

  useEffect(() => {
    const fetchRates = async () => {
      try {
        // 替換成你真正的 C# API 網址
        //const apiUrl = import.meta.env.VITE_API_URL || '';
        const apiUrl = 'https://localhost:7080';
        //const apiUrl = 'https://languverse-quotesystem-api-20260322221023-fxe9b4e0brcud5e9.eastasia-01.azurewebsites.net/api/ExchangeRate/latest';
        console.log('ExchangeRate.jsx' + apiUrl)
        const response = await axios.get(apiUrl +'/api/ExchangeRate/latest');
        
        // 1. 自己元件內部存一份用來顯示畫面
        setExchangeRates(response.data);
        
        // 2. 把資料往上傳給 App.jsx，讓 App.jsx 可以拿來算錢
        if (onRatesLoaded) {
          onRatesLoaded(response.data);
        }
      } catch (err) {
        console.error('抓取匯率失敗:', err);
        setError('無法取得最新匯率，請稍後再試。');
      } finally {
        setIsLoading(false);
      }
    };

    fetchRates();
  //}, [onRatesLoaded]); // 依賴陣列加入 onRatesLoaded
  // 🌟 ✅ 把這裡改成空陣列！代表「只在第一次顯示時抓一次」
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); 

  // 畫面渲染區塊
  return (
    <div style={{ backgroundColor: '#f0f9ff', padding: '16px', borderRadius: '8px', marginBottom: '24px' }}>
      <h3 style={{ fontSize: '18px', fontWeight: 'bold', margin: '0 0 8px 0' }}>💰 今日參考匯率</h3>
      
      {isLoading && <p>匯率載入中...</p>}
      {error && <p style={{ color: 'red' }}>{error}</p>}
      
      {exchangeRates && (
        <div style={{ display: 'flex', gap: '24px', fontSize: '14px', color: '#374151' }}>
          <div>
            <strong>🇺🇸 美金 (USD)</strong><br/>
            銀行: {exchangeRates.usd.bank}<br/>
            賣出價: <span style={{ color: '#2563eb', fontWeight: 'bold', fontSize: '16px' }}>{exchangeRates.usd.sellRate}</span>
          </div>
          <div>
            <strong>🇵🇭 披索 (PHP)</strong><br/>
            銀行: {exchangeRates.php.bank}<br/>
            賣出價: <span style={{ color: '#2563eb', fontWeight: 'bold', fontSize: '16px' }}>{exchangeRates.php.sellRate}</span>
          </div>
        </div>
      )}
    </div>
  );
}
