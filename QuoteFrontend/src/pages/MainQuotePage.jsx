import React, { useState, useEffect } from 'react';
import { DatePicker, Select, Button, Table, Card, Row, Col,
   InputNumber, Input, Switch, message, Modal,Form,Space} from 'antd';
import { CalculatorOutlined, FileExcelOutlined, DeleteOutlined, PlusOutlined,
  DollarOutlined,DownloadOutlined,UploadOutlined,UserOutlined,QuestionCircleOutlined} from '@ant-design/icons';
import axios from 'axios';
import dayjs from 'dayjs';
import 'antd/dist/reset.css';
// ✅ default import（不用大括號）
import FlightSearch from '../components/FlightSearch';
// 1. 引入剛剛寫好的匯率元件 (請確認路徑是否正確)
import ExchangeRate from '../components/ExchangeRate'; 



const handleSaveJson = () => {
  const cachedData = window.quoteCache || {};
  
  const quoteData = {
    // 全從快取取，無 request！
    basicInfo: {
      school: cachedData.request?.School || "PHILINTER",
      course: cachedData.request?.Course || "",
      roomType: cachedData.request?.RoomType || "",
      placeOfStay: cachedData.request?.Placeofstay || "",
      startDate: cachedData.request?.StartDate || "",
      endDate: cachedData.request?.EndDate || "",
      weeks: cachedData.request?.weeks || 0,
      usdRate: cachedData.request?.UsaExchangeRate || 32.5,
      phpRate: cachedData.request?.PhpExchangeRate || 0.6,
      airTicket: cachedData.request?.AirTicket || 0,
      visa: cachedData.request?.Visa || 0,
      insurance: cachedData.request?.Insurance || 0
    },
    courseFees: cachedData.quote?.courseFees || [],
    localFees: cachedData.quote?.localFees || [],
    otherFees: cachedData.quote?.otherFees || [],
    totals: cachedData.totals || {
      currentTotalUSD: 0,
      currentTotalNTD: 0,
      currentLocalTotalPeso:0,
      currentLocalTotalNTD:0,
      currentOtherTotalNTD:0,
      AllTotalNTD: 0,
    },
    timestamp: new Date().toISOString()
  };

  console.table(quoteData);
  message.success(`存檔 ${quoteData.courseFees.length}課程 + ${quoteData.localFees.length}當地費`);

  const jsonStr = JSON.stringify(quoteData, null, 2);
  const blob = new Blob([jsonStr], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `報價單_${quoteData.basicInfo.school}.json`;  // 直接用 school
  link.click();
  URL.revokeObjectURL(url);
};






const handleExportExcel = async () => {
  const cachedData = window.quoteCache || {};
  
  const quoteData = {
    // 全從快取取，無 request！
    basicInfo: {
      studentName: cachedData.request?.studentName || "",
      school: cachedData.request?.School || "PHILINTER",
      course: cachedData.request?.Course || "",
      roomType: cachedData.request?.RoomType || "",
      placeOfStay: cachedData.request?.Placeofstay || "",
      startDate: cachedData.request?.StartDate || "",
      endDate: cachedData.request?.EndDate || "",
      weeks: cachedData.request?.weeks || 0,
      usdRate: cachedData.request?.UsaExchangeRate || 32.5,
      phpRate: cachedData.request?.PhpExchangeRate || 0.6,
      airTicket: cachedData.request?.AirTicket || 0,
      visa: cachedData.request?.Visa || 0,
      insurance: cachedData.request?.Insurance || 0
    },
    courseFees: cachedData.quote?.courseFees || [],
    localFees: cachedData.quote?.localFees || [],
    otherFees: cachedData.quote?.otherFees || [],
    totals: cachedData.totals || {
      currentTotalUSD: 0,
      currentTotalNTD: 0,
      currentLocalTotalPeso:0,
      currentLocalTotalNTD:0,
      currentOtherTotalNTD:0,
      AllTotalNTD: 0,
    },
    timestamp: new Date().toISOString()
  };

  console.log('測試 quoteData:', quoteData);

  try {
    const response = await axios.post(apiUrl + '/api/ExportQuote/export', quoteData, {
      responseType: 'blob',
      headers: { 'Content-Type': 'application/json' }
    });

    // const blob = new Blob([response.data], { 
    //   type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    // });
    // const url = window.URL.createObjectURL(blob);
    // const link = document.createElement('a');
    // link.href = url;
    // link.download = '測試報價單.xlsx';
    // link.click();
    // window.URL.revokeObjectURL(url);
    // console.log('✅ 匯出成功！');
  } catch (error) {
    console.error('❌ 錯誤:', error.response?.status || error.message);
  }
};







const { RangePicker } = DatePicker;
const { Option } = Select;
const apiUrl = import.meta.env.VITE_API_URL || 'https://localhost:7080';
//const apiUrl = 'https://localhost:7080';

// axios.defaults.baseURL = 'https://localhost:5129';  // 你的後端端口

function App () {
  // 2. 在 App.jsx 準備一個 state 來存放匯率，為了後續算總價用
  const [currentRates, setCurrentRates] = useState(null);

  // 當匯率元件抓好資料時，會呼叫這個函式
  const handleRatesLoaded = (ratesData) => {
    setCurrentRates(ratesData);
    console.log("App.jsx 成功接收到匯率：", ratesData);
  };


  const [quote, setQuote] = useState(null);
  const [showFlightSearch, setShowFlightSearch] = useState(false);



  const [request, setRequest] = useState({
    StartDate: dayjs().format('YYYY-MM-DD'),
    EndDate: dayjs().add(4, 'week').format('YYYY-MM-DD'),
    School: 'PHILINTER',
    Course: 'IELTS INTENSIVE',
    RoomType: 'Single',
    Placeofstay: 'DORMITORY',
    UsaExchangeRate: 32.00,
    PhpExchangeRate: 0.6,
    AirTicket: 10000,
    Visa: 1200,
    Insurance: 1500,
    NeedGuardianFee: false, // 👈 新增這行：預設為不需要 (false)
    SchoolDiscount: 0 // 👈 新增這行：預設折扣為 0 (代表美金金額)
    
  });
  const [form] = Form.useForm();
  const [schools, setSchools] = useState([]);
  const [loading, setLoading] = useState(false);

  const [sheetOptions, setSheetOptions] = useState([]); 
  const [loadingSheets, setLoadingSheets] = useState(false);

  const [coursesList, setCoursesList] = useState([]); 
  const [roomsList, setRoomsList] = useState([]);     
  const [loadingDetails, setLoadingDetails] = useState(false); 

  const [dormitoryList, setDormitorysList] = useState([]);    
  
  // 新增狀態
const [selectedFile, setSelectedFile] = useState(null);


const handleSendQuoteData = async () => {
  const cachedData = window.quoteCache || {};
const quoteData = {
    // 全從快取取，無 request！
    basicInfo: {
      studentName: cachedData.request?.studentName || "",
      school: cachedData.request?.School || "PHILINTER",
      course: cachedData.request?.Course || "",
      roomType: cachedData.request?.RoomType || "",
      placeOfStay: cachedData.request?.Placeofstay || "",
      startDate: cachedData.request?.StartDate || "",
      endDate: cachedData.request?.EndDate || "",
      weeks: cachedData.request?.weeks || 0,
      usdRate: cachedData.request?.UsaExchangeRate || 32.5,
      phpRate: cachedData.request?.PhpExchangeRate || 0.6,
      airTicket: cachedData.request?.AirTicket || 0,
      visa: cachedData.request?.Visa || 0,
      insurance: cachedData.request?.Insurance || 0
    },
    // courseFees: cachedData.quote?.courseFees || [],
    // localFees: cachedData.quote?.localFees || [],
    // otherFees: cachedData.quote?.otherFees || [],
    // totals: cachedData.totals || {
    //   currentTotalUSD: 0,
    //   currentTotalNTD: 0,
    //   currentLocalTotalPeso:0,
    //   currentLocalTotalNTD:0,
    //   currentOtherTotalNTD:0,
    //   AllTotalNTD: 0,
    // },
    // ✅ 優先用 React state，手動編輯才會傳到後端！
    courseFees: quote?.courseFees || cachedData.quote?.courseFees || [],
    localFees: quote?.localFees || cachedData.quote?.localFees || [],
    otherFees: quote?.otherFees || cachedData.quote?.otherFees || [],
    totals: quote?.totals || cachedData.totals || {
      currentTotalUSD: 0,
      currentTotalNTD: 0,
      currentLocalTotalPeso: 0,
      currentLocalTotalNTD: 0,
      currentOtherTotalNTD: 0,
      AllTotalNTD: 0,
    },
    timestamp: new Date().toISOString()
  };
  console.table('✅ 傳送 quoteData:', quoteData);
  console.log('otherFees[0].content:', quoteData.otherFees[0]?.content);
  //console.log('測試 quoteData:', quoteData);
  
  // ✅ 轉 JSON 檔案 → FormData
  const jsonStr = JSON.stringify(quoteData, null, 2);
  const blob = new Blob([jsonStr], { type: 'application/json' });
  const file = new File([blob], `quote_${quoteData.basicInfo.school}.json`, { 
    type: 'application/json' 
  });
  
  const formData = new FormData();
  formData.append('quoteJson', file);  // 關鍵：匹配後端參數名
  
  try {
    const response = await axios.post('/api/ExportQuote/from-file', formData, {
      responseType: 'blob'  // Excel 檔案
    });
    
    // 自動下載
    const url = window.URL.createObjectURL(response.data);
    const link = document.createElement('a');
    link.href = url;
    link.download = `報價單_${quoteData.basicInfo.school}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
    
    message.success('✅ Excel 下載成功！');
  } catch (error) {
    console.error('❌ 錯誤:', error.response?.status, error.message);
    message.error(`匯出失敗: ${error.response?.status || error.message}`);
  }
};

// 1. 選擇檔案
const handleFileSelect = (e) => {
  setSelectedFile(e.target.files[0]);
  message.success(`選擇 ${e.target.files[0]?.name}`);
};

// 2. 上傳檔案到後端
const handleUploadJson = async () => {
  if (!selectedFile) return message.error('請選擇 JSON');

  const formData = new FormData();
  formData.append('quoteJson', selectedFile);
  console.log('quoteJson' + formData)

  try {
    const response = await axios.post(apiUrl + '/api/ExportQuote/from-file', formData, {
      responseType: 'blob',
      // headers: { 'Content-Type': 'multipart/form-data' }
    });

    
  } catch (error) {
    message.error('上傳失敗');
  }
};

  // ==========================================
  // 0. 獨立的抓取學校詳細資料函數 (課程、房型、頁籤)
  // ==========================================
  // 0. 獨立的抓取學校詳細資料函數 (課程、房型、頁籤)
  const fetchSchoolDetails = async (schoolName) => {
    if (!schoolName) return;

    // 1. 抓取頁籤
    setLoadingSheets(true);
    try {
      const response = await axios.get(apiUrl+`/api/quote/school-sheets/${schoolName}`);
      const dataIsArray = Array.isArray(response.data) ? response.data : [];
      setSheetOptions(dataIsArray); 
      
      if (dataIsArray.length > 0) {
        setRequest(prev => ({ ...prev, SheetName: dataIsArray[0] }));
      } else {
        setRequest(prev => ({ ...prev, SheetName: '' }));
      }
    } catch (error) {
      console.log("讀取頁籤失敗，忽略不計", error.message);
      setSheetOptions([]);
      setRequest(prev => ({ ...prev, SheetName: '' }));
    } finally {
      setLoadingSheets(false);
    }

    // 2. 抓取課程與房型
    setLoadingDetails(true);
    try {
      console.log(`準備發送 API 請求抓取 [${schoolName}] 的課程...`);
      const response = await axios.get(apiUrl+`/api/quote/school-details?schoolName=${schoolName}`);
      
      console.log("後端成功回傳：", response.data);

      const courses = response.data?.courses || [];
      const rooms = response.data?.rooms || [];
      const Placeofstays = response.data?.placeofstays  || [];
      
      if (courses.length === 0) {
          // 如果真的抓不到資料，雙重提示！
          message.warning(`學校【${schoolName}】的資料還沒建置喔！`);
          // 如果 message 壞掉，至少還有這個
          // alert(`警告：學校【${schoolName}】目前沒有課程資料！`); 
          
          setCoursesList([]);
          setRoomsList([]);
          setDormitorysList([]);
          setRequest(prev => ({ ...prev, Course: '', RoomType: '',  Placeofstay: ''}));
      } else {
          message.success(`成功載入【${schoolName}】的資料`);
          setCoursesList(courses);
          setRoomsList(rooms);
          setDormitorysList(Placeofstays);

          setRequest(prev => ({
            ...prev,
            Course: courses[0],
            RoomType: rooms.length > 0 ? rooms[0] : '',
            Placeofstay: Placeofstays.length > 0 ? Placeofstays[0] : ''
          }));
      }
      
    } catch (error) {
      console.error(`💥 抓取 [${schoolName}] 發生致命錯誤：`, error);
      
      // 雙重提示：保證你一定看得見！
      message.error(`無法取得【${schoolName}】的資料！`);
      alert(`發生錯誤：無法讀取 ${schoolName} 的資料\n錯誤代碼：${error.message}`);
      
      setCoursesList([]);
      setRoomsList([]);
      setDormitorysList([]);
      setRequest(prev => ({ ...prev, Course: '', RoomType: '',  Placeofstay: ''}));
      
    } finally {
      setLoadingDetails(false);
    }
  };


  // ==========================================
  // 1. 先定義操作函數 (新增、修改、刪除)
  // ==========================================
  const addCourseFeeRow = () => {
    setQuote(prev => {
      if (!prev) return prev;
      
      const newFees = [...(prev.courseFees || [])];
      
      newFees.push({ 
        key: `custom_${Date.now()}`,
        item: '手動新增項目', 
        content: '', 
        weeks: '', 
        people: 1, 
        unitPrice: 0, 
        amount: 0, 
        remark: '' 
      });

      return { ...prev, courseFees: newFees };
    });
  };

  const handleCourseFeeChange = (index, field, value) => {
    setQuote(prev => {
      if (!prev) return prev;
      
      const newFees = [...prev.courseFees];
      newFees[index] = { ...newFees[index], [field]: value };
      
      // 👇 不管改什麼欄位，都重新計算一次該列的 amount
      const p = Number(newFees[index].people) || 1; // 確保人數至少為1
      const u = Number(newFees[index].unitPrice) || 0;
      newFees[index].amount = p * u;

      return { 
        ...prev, 
        courseFees: newFees 
      };
    });
  };
  
  const removeCourseFeeRow = (indexToRemove) => {
      setQuote(prev => {
          if (!prev) return prev;
          const newFees = prev.courseFees.filter((_, idx) => idx !== indexToRemove);
          return {
              ...prev,
              courseFees: newFees
          };
      });
  };

  // ==========================================
  // 2. 再定義 Table Columns
  // ==========================================
  const courseFeeColumns = [
    { 
      title: '課程費用項目', 
      dataIndex: 'item',      
      key: 'item',
      align: 'center',  
      width: 150, 
      render: (text, record, index) => (
        <Input 
          value={text} 
          onChange={e => handleCourseFeeChange(index, 'item', e.target.value)} 
          // style={{ textAlign: 'center' }}
        />
      ) 
    },
    { 
      title: '費用內容', 
      dataIndex: 'content',   
      key: 'content',
      align: 'center',  
      width: 200,
      render: (text, record, index) => (
        <Input 
          value={text} 
          onChange={e => handleCourseFeeChange(index, 'content', e.target.value)} 
        />
      )
    },
    { title: '週數', dataIndex: 'weeks', key: 'weeks', align: 'center', width: 80 },
    { 
      title: '人數', 
      dataIndex: 'people', 
      key: 'people', 
      align: 'center', 
      width: 90,
      render: (val, record, index) => (
        <InputNumber 
          min={1} 
          value={val} 
          onChange={v => handleCourseFeeChange(index, 'people', v || 1)} 
          style={{ width: '100%' }}
        />
      )
    },
    { 
      title: '單價', 
      dataIndex: 'unitPrice', 
      key: 'unitPrice', 
      align: 'center',  
      width: 120,
      render: (val, record, index) => (
        <InputNumber 
          value={val} 
          prefix="US$"
          onChange={v => handleCourseFeeChange(index, 'unitPrice', v || 0)} 
          style={{ width: '100%' }}
        />
      )
    },
    { 
      title: '金額(美金)', 
      dataIndex: 'amount',    
      key: 'amount', 
      align: 'center',  
      width: 120, 
      render: (val, record, index) => {
        const calculatedAmount = Number(record.people || 1) * Number(record.unitPrice || 0);
        return (
          // <InputNumber 
          //   value={calculatedAmount} 
          //   readOnly 
          //   style={{ 
          //     width: '100%', 
          //     color: calculatedAmount < 0 ? 'green' : '#cf1322', 
          //     fontWeight: 'bold',
          //     backgroundColor: '#fafafa' 
          //   }}
          // />
          <span style={{ 
          fontWeight: 'bold', 
          color: '#cf1322',
          fontSize: 16
        }}>
          US${calculatedAmount.toLocaleString()}
        </span>
        );
      }
    },
    { 
      title: '備註', 
      dataIndex: 'remark', 
      key: 'remark',
      align: 'center', 
      render: (text, record, index) => (
        <Input 
          value={text} 
          onChange={e => handleCourseFeeChange(index, 'remark', e.target.value)} 
        />
      )
    }, 
    {
      title: '操作',
      key: 'action',
      align: 'center',
      width: 80,
      render: (_, record, index) => (
        <Button danger size="small" icon={<DeleteOutlined />}onClick={() => removeCourseFeeRow(index)}>
          刪除
        </Button>
      )
    }
  ];

  // ====== 當地雜費操作函數 ======
const addLocalFeeRow = () => {
  setQuote(prev => {
    if (!prev) return prev;
    const newFees = [...(prev.localFees || [])];
    newFees.push({
      key: `local_${Date.now()}`,
      item: '手動新增項目',
      content: '',
      times: '',
      people: 1,
      unitPrice: 0,
      amount: 0,
      remark: ''
    });
    return { ...prev, localFees: newFees };
  });
};

const handleLocalFeeChange = (index, field, value) => {
  setQuote(prev => {
    if (!prev) return prev;
    const newFees = [...prev.localFees];
    newFees[index] = { ...newFees[index], [field]: value };

    // 自動重算 amount = people * unitPrice
    const p = Number(newFees[index].people) || 1;
    const u = Number(newFees[index].unitPrice) || 0;
    newFees[index].amount = p * u;

    return { ...prev, localFees: newFees };
  });
};

const removeLocalFeeRow = (indexToRemove) => {
  setQuote(prev => {
    if (!prev) return prev;
    const newFees = prev.localFees.filter((_, idx) => idx !== indexToRemove);
    return { ...prev, localFees: newFees };
  });
};

// 👇 新增：手動新增其他費用行 用處：按鈕點擊新增一行（如接機費、行李費）。
const addOtherFeeRow = () => {
  setQuote(prev => {
    if (!prev) return prev;
    const newFees = [...(prev.otherFees || [])];
    newFees.push({
      key: `other_${Date.now()}`,
      item: '手動新增費用',
      content: '',
      people: 1,
      unitPrice: 0,
      amount: 0,
      remark: ''
    });
    return { ...prev, otherFees: newFees };
  });
};

// 👇 新增：修改其他費用儲存 + 自動重算 用處：修改人數/單價時，自動重算金額。
const handleOtherFeeChange = (index, field, value) => {
  setQuote(prev => {
    if (!prev) return prev;
    const newFees = [...(prev.otherFees || [])];
    newFees[index] = { ...newFees[index], [field]: value };
    
    // 👈 自動重算：金額 = 人數 × 單價
    const p = Number(newFees[index].people) || 1;
    const u = Number(newFees[index].unitPrice) || 0;
    newFees[index].amount = p * u;
    
    return { ...prev, otherFees: newFees };
  });
};


// 👇 新增其他費用刪除 用處：每行右邊的「刪除」按鈕。
const removeOtherFeeRow = (indexToRemove) => {
  setQuote(prev => {
    if (!prev) return prev;
    const newFees = prev.otherFees.filter((_, idx) => idx !== indexToRemove);
    return { ...prev, otherFees: newFees };
  });
};


// ====== 當地雜費欄位定義（披索） ======
const localFeeColumns = [
  { 
    title: '當地雜費項目', 
    dataIndex: 'item', 
    key: 'item',
    align: 'center',  
    width: 150,
    render: (text, record, index) => (
      <Input 
        value={text} 
        onChange={e => handleLocalFeeChange(index, 'item', e.target.value)} 
      />
    )
  },
  { 
    title: '費用內容', 
    dataIndex: 'content', 
    key: 'content',
    align: 'center',  
    width: 220,
    render: (text, record, index) => (
      <Input 
        value={text} 
        onChange={e => handleLocalFeeChange(index, 'content', e.target.value)} 
      />
    )
  },
  { 
    title: '次數/週數', dataIndex: 'weeks', key: 'weeks', align: 'center', width: 80 ,
    // title: '次數/週數', 
    // dataIndex: 'Weeks', 
    // key: 'Weeks', 
    // width: 100,
    // render: (text) => <span style={{ fontWeight: 'bold' }}>{text || ''}</span>  // 👈 唯讀
    // render: (text, record, index) => (
    //   <Input 
    //     value={text} 
    //     onChange={e => handleLocalFeeChange(index, 'Weeks', e.target.value)} 
    //   />
    // )
  },
  { 
    title: '人數', 
    dataIndex: 'people', 
    key: 'people', 
    align: 'center',
    width: 80,
    render: (val, record, index) => (
      <InputNumber 
        min={1}
        value={val} 
        onChange={v => handleLocalFeeChange(index, 'people', v || 1)} 
        style={{ width: '100%' }}
      />
    )
  },
  { 
    title: '單價 (披索)', 
    dataIndex: 'unitPrice', 
    key: 'unitPrice', 
    align: 'center',  
    width: 120,
    render: (val, record, index) => (
      <InputNumber 
        value={val} 
        prefix="₱"
        onChange={v => handleLocalFeeChange(index, 'unitPrice', v || 0)} 
        style={{ width: '100%' }}
      />
    )
  },
  // { 
  //   title: '金額 (披索)', 
  //   dataIndex: 'amount', 
  //   key: 'amount', 
  //   align: 'right',
  //   width: 120,
  //   render: (val, record) => {
  //     const p = Number(record.people || 1);
  //     const u = Number(record.unitPrice || 0);
  //     const calculated = p * u;
  //     return (
  //       <span style={{ fontWeight: 'bold' }}>
  //         {calculated.toLocaleString()}
  //       </span>
  //     );
  //   }
  // },
  { 
      title: '金額(披索)', 
      dataIndex: 'amount',    
      key: 'amount', 
      align: 'center',  
      width: 120, 
      render: (val, record, index) => {
        const calculatedAmount = Number(record.people || 1) * Number(record.unitPrice || 0);
        return (
          // <InputNumber 
          //   value={calculatedAmount} 
          //   readOnly 
          //   style={{ 
          //     width: '100%', 
          //     color: calculatedAmount < 0 ? 'green' : '#cf1322', 
          //     fontWeight: 'bold',
          //     backgroundColor: '#fafafa' 
          //   }}
          // />
          <span style={{ 
          fontWeight: 'bold', 
          color: '#cf1322',
          fontSize: 16
        }}>
          ₱{calculatedAmount.toLocaleString()}
        </span>
        );
      }
    },
  { 
    title: '備註', 
    dataIndex: 'remark', 
    key: 'remark',
    align: 'center', 
    render: (text, record, index) => (
      <Input 
        value={text} 
        onChange={e => handleLocalFeeChange(index, 'remark', e.target.value)} 
      />
    )
  },
  {
    title: '操作',
    key: 'action',
    align: 'center', 
    width: 80,
    render: (_, record, index) => (
      <Button danger size="small" icon={<DeleteOutlined />} onClick={() => removeLocalFeeRow(index)}>
        刪除
      </Button>
    )
  }
];


//
// ====== 機票簽證保險欄位定義（台幣） ======
const otherFeeColumns = [
  { 
    title: '費用項目', 
    dataIndex: 'item', 
    key: 'item',
    align: 'center',  
    width: 150,
    render: (text, record, index) => (
      <Input 
        value={text} 
        onChange={e => handleOtherFeeChange(index, 'item', e.target.value)} 
        placeholder="如：來回機票"
      />
    )
  },
  { 
      title: '費用內容', 
      dataIndex: 'content',   
      key: 'content',
      align: 'center',  
      width: 200,
      render: (text, record, index) => (
        <Input 
          value={text} 
          onChange={e => handleOtherFeeChange(index, 'content', e.target.value)} 
        />
      )
  },
  { title: '', dataIndex: 'weeks', key: 'weeks', align: 'center', width: 80 },
  { 
    title: '人數', 
    dataIndex: 'people', 
    key: 'people', 
    align: 'center', 
    width: 80,
    render: (val, record, index) => (
      <InputNumber 
        min={1}
        value={val} 
        onChange={v => handleOtherFeeChange(index, 'people', v || 1)} 
        style={{ width: 80 }}
      />
    )
  },
  { 
    title: '單價 (台幣)', 
    dataIndex: 'unitPrice', 
    key: 'unitPrice', 
    align: 'center', 
    width: 120,
    render: (val, record, index) => (
      <InputNumber 
        value={val} 
        onChange={v => handleOtherFeeChange(index, 'unitPrice', v || 0)} 
        prefix="NT$"
        style={{ width: '100%' }}
        precision={0}
      />
    )
  },
  { 
    title: '金額 (台幣)', 
    dataIndex: 'amount', 
    key: 'amount', 
    align: 'center', 
    width: 120,
    render: (val, record, index) => {
      const p = Number(record.people || 1);
      const u = Number(record.unitPrice || 0);
      const total = p * u;
      return (
        <span style={{ 
          fontWeight: 'bold', 
          color: '#cf1322',
          fontSize: 16
        }}>
          NT${total.toLocaleString()}
        </span>
      );
    }
  },
  { 
    title: '備註', 
    dataIndex: 'remark', 
    key: 'remark',
    width: 150,
    render: (text, record, index) => (
      <Input 
        value={text} 
        onChange={e => handleOtherFeeChange(index, 'remark', e.target.value)} 
        placeholder="航班/簽證類型"
      />
    )
  },
  {
    title: '操作',
    key: 'action',
    width: 80,
    align: 'center',
    render: (_, record, index) => (
      <Button 
        danger 
        size="small" 
        icon={<DeleteOutlined />}
        onClick={() => removeOtherFeeRow(index)}
      >
        刪除
      </Button>
    )
  }
];


  // ==========================================
  // 3. 呼叫後端的 Calculate API
  // ==========================================
  const calculate = async () => {
    


    setLoading(true);
    try {
      const weeks = Math.ceil(dayjs(request.EndDate).diff(dayjs(request.StartDate), 'day') / 7);
      const res = await axios.post(apiUrl+'/api/quote/calculate', { ...request, Weeks: weeks });
      
      console.log('API 成功！', res.data); 
      
      setQuote({
        ...res.data,
        courseFees: res.data.courseFees || [] ,
        localFees:  res.data.localFees  || [],   // 👈 這裡就有值
        otherFees:  res.data.otherFees  || []
        // otherFees: [  // 👈 新增機票簽證保險！
        // { 
        //   key: 'air', 
        //   item: '來回機票 (台北-宿霧)', 
        //   people: 1, 
        //   unitPrice: request.AirTicket, 
        //   amount: request.AirTicket,
        //   remark: '可依航班調整'
        // },
        // { 
        //   key: 'visa', 
        //   item: '簽證費 (9A旅遊簽)', 
        //   people: 1, 
        //   unitPrice: request.Visa, 
        //   amount: request.Visa,
        //   remark: '線上申請'
        // },
        // { 
        //   key: 'ins', 
        //   item: '旅平險/醫療險', 
        //   people: 1, 
        //   unitPrice: request.Insurance, 
        //   amount: request.Insurance,
        //   remark: '建議投保3個月'
        // }
        //]
      });

      // ← 加這區塊：存全域快取
      window.quoteCache = {
        request: { ...request, weeks },  // 基本資訊
        quote: res.data,                 // 表格資料
        totals: {
          currentTotalUSD: quote?.courseFees?.reduce((sum, item) => {
                            const p = Number(item.people || 1);
                            const u = Number(item.unitPrice || 0);
                            return sum + (p * u);
                          }, 0) || 0,
          currentTotalNTD: currentTotalUSD * Number(currentRates?.usd?.sellRate),
          currentLocalTotalPeso:quote?.localFees?.reduce((sum, item) => {
                                  const p = Number(item.people || 1);
                                  const u = Number(item.unitPrice || 0);
                                  return sum + p * u;
                                }, 0) || 0,
          currentLocalTotalNTD:currentLocalTotalPeso * Number(currentRates?.php?.sellRate),
          currentOtherTotalNTD:quote?.otherFees?.reduce((sum, item) => {
                                  return sum + Number(item.amount || 0);
                                }, 0) || 0,
          AllTotalNTD: currentTotalNTD + currentLocalTotalNTD + currentOtherTotalNTD
        // ... 其他
        },
        timestamp: new Date().toISOString()
      };
      console.log('💾 全域快取更新:', window.quoteCache);
      
    } catch (error) {
      console.error('API 錯誤：', error); 
      alert('計算失敗：' + error.message);
    }
    setLoading(false);
  };

  const exportExcel = async () => {
    try {
      const res = await axios.post(apiUrl+'/api/quote/export', quote, { 
        responseType: 'blob',
        headers: { 'Content-Type': 'application/json' }
      });
      const url = URL.createObjectURL(res.data);
      const a = document.createElement('a');
      a.href = url;
      a.download = '報價單.xlsx';
      a.click();
    } catch (error) {
      alert('匯出失敗');
    }
  };

  // ==========================================
  // 4. useEffect 載入資料 (這部分已為你修復衝突與錯誤)
  // ==========================================
  useEffect(() => {
    const loadSchools = async () => {
      try {
        const res = await axios.get(apiUrl+'/api/quote/school-list');
        setSchools(res.data);
        console.log('School list loaded:', res.data);
        
        const defaultSchool = request.School || (res.data.length > 0 ? res.data[0] : null);
        
        if (defaultSchool) {
            if(request.School !== defaultSchool){
                setRequest(prev => ({...prev, School: defaultSchool}))
            }
            fetchSchoolDetails(defaultSchool);
        }
      } catch (error) {
        console.error('載入學校失敗', error);
        const fallback = ["PHILINTER", "EV", "PINES", "JIC"]; // 修復變數未定義錯誤
        setSchools(fallback);
        
        if(request.School){
             fetchSchoolDetails(request.School);
        }
      }
    };
    loadSchools();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const [prevSchool, setPrevSchool] = useState(request.School);

  useEffect(() => {
    if (request.School && request.School !== prevSchool) {
        console.log("偵測到學校切換:", prevSchool, "->", request.School);
        fetchSchoolDetails(request.School);
        setPrevSchool(request.School); 
    }
  }, [request.School, prevSchool]);


    // ==========================================
    // 動態計算當前表格的總金額 (單價 * 人數)
    // ==========================================
    const currentTotalUSD = quote?.courseFees?.reduce((sum, item) => {
      const p = Number(item.people || 1);
      const u = Number(item.unitPrice || 0);
      return sum + (p * u);
    }, 0) || 0;

    const currentTotalNTD = currentTotalUSD * Number(currentRates?.usd?.sellRate);

    // 披索合計（當地雜費）
    const currentLocalTotalPeso = quote?.localFees?.reduce((sum, item) => {
      const p = Number(item.people || 1);
      const u = Number(item.unitPrice || 0);
      return sum + p * u;
    }, 0) || 0;

      // 假設你有一個固定的披索對台幣匯率（例如 0.6）
      const pesoToNT = Number(currentRates?.php?.sellRate);
      const currentLocalTotalNTD = currentLocalTotalPeso * Number(currentRates?.php?.sellRate);
// 👇 新增這行：計算其他費用小計 用處：自動算機票簽證總金額，顯示在 Table 底下。
const currentOtherTotalNTD = quote?.otherFees?.reduce((sum, item) => {
  return sum + Number(item.amount || 0);
}, 0) || 0;
    // ==========================================
    // 5. 畫面渲染 Render
    // ==========================================

    // ==========================================
    // 5. 畫面渲染 Render
    // ==========================================
  return (
    <div style={{ padding: 24, maxWidth: 1200, margin: '0 auto' }}>
      <h1 style={{ textAlign: 'center' }}>語宙🌍報價系統📊</h1>
      {/* 3. 把匯率元件放在這裡，並把接收資料的函式傳遞給它 */}
      <ExchangeRate onRatesLoaded={handleRatesLoaded} />
      {/* 下面是你的報價表單其他內容 */}
      <div style={{ border: '1px solid #ccc', padding: '20px', borderRadius: '8px' }}>
        <h2>📝 報價單內容</h2>
        
        {/* 示範：如果在 App.jsx 需要拿匯率來計算 */}
        {currentRates ? (
          <p>
            如果學費是 1000 美金，折合台幣大約是：
            <strong>{Math.round(1000 * parseFloat(Number(currentRates?.usd?.sellRate)))} 元</strong>
          </p>
        ) : (
          <p>等待匯率載入中，暫時無法計算...</p>
        )}
      </div>


      <Card title="基本資訊">
        <Row gutter={16}>
          <Col xs={24} md={12}>
            <div style={{ marginBottom: 16 }}>
              <label style={{ 
                display: 'flex', 
                alignItems: 'center', 
                gap: 8,
                fontWeight: 600, 
                color: '#1677ff',
                marginBottom: 8
              }}>
                <UserOutlined style={{ fontSize: 16 }} />
                學生姓名 <span style={{ color: '#ff4d4f', fontSize: 14 }}>*必填</span>
              </label>
              
              <Input
                value={request.studentName || ''}
                onChange={(e) => setRequest({ ...request, studentName: e.target.value })}
                placeholder="王小明、李小華（逗號分隔多位）"
                allowClear
                maxLength={20}
                status={request.studentName ? '' : 'warning'}
                style={{ width: '100%' }}
              />
              
              <div style={{ 
                display: 'flex', 
                justifyContent: 'space-between',
                fontSize: 12, 
                color: '#8c8c8c', 
                marginTop: 4
              }}>
                <span style={{ cursor: 'help' }} title="0-20字">
                  支援多位學生（逗號分隔）
                </span>
                <span>{(request.studentName || '').length}/20</span>
              </div>
            </div>
          </Col>

          <Col xs={24} md={12}>
            <label style={{ fontWeight: 'bold' }}>📅 預計出發日期與週數:</label>
            <RangePicker 
              value={[dayjs(request.StartDate), dayjs(request.EndDate)]}
              onChange={([s, e]) => setRequest({
                ...request,
                StartDate: s.format('YYYY-MM-DD'),
                EndDate: e.format('YYYY-MM-DD')
              })}
              style={{ width: '100%', marginTop: 8 }}
            />
          </Col>
        </Row>

        <div style={{ margin: '20px 0', borderBottom: '1px solid #f0f0f0' }}></div>

        <Row gutter={16}>
          <Col xs={24} md={12}>
            <label style={{ fontWeight: 'bold' }}>🏫 選擇學校:</label>
            <Select 
              value={request.School} 
              onChange={v => setRequest({ ...request, School: v })}
              loading={schools.length === 0}
              style={{ width: '100%', marginTop: 8 }}
              placeholder="載入中..."
            >
              {schools.map(s => (
                <Option key={s} value={s}>{s}</Option>
              ))}
            </Select>
          </Col>
          
          <Col xs={24} md={12}>
            <label style={{ fontWeight: 'bold', color: '#1890ff' }}>📚 選擇課程:</label>
            <Select 
              value={request.Course || undefined} 
              onChange={v => setRequest({ ...request, Course: v })}
              style={{ width: '100%', marginTop: 8 }}
              loading={loadingDetails}
              disabled={coursesList.length === 0}
              placeholder={coursesList.length === 0 ? "無課程資料" : "請選擇課程"}
              showSearch 
            >
              {coursesList.map(courseName => (
                <Option key={courseName} value={courseName}>
                    {courseName}
                </Option>
              ))}
            </Select>
          </Col>

          <Col xs={24} md={12}>
            <label style={{ fontWeight: 'bold', color: '#52c41a' }}>🛏️ 選擇房型:</label>
            
            <Select 
              value={request.RoomType || undefined} 
              onChange={v => setRequest({ ...request, RoomType: v })}
              style={{ width: '100%', marginTop: 8 }}
              loading={loadingDetails}
              disabled={roomsList.length === 0}
              placeholder={roomsList.length === 0 ? "無房型資料" : "請選擇房型"}
            >
              {roomsList.map(roomName => (
                <Option key={roomName} value={roomName}>
                  {roomName}
                </Option>
              ))}
            </Select>
          </Col>

          <Col xs={24} md={12}>
            <label style={{ fontWeight: 'bold', color: '#52c41a' }}>🛏️ 選擇宿舍:</label>
            
            <Select 
              value={request.Placeofstay || undefined} 
              onChange={v => setRequest({ ...request, Placeofstay: v })}
              style={{ width: '100%', marginTop: 8 }}
              loading={loadingDetails}
              disabled={dormitoryList.length === 0}
              placeholder={dormitoryList.length === 0 ? "無房型資料" : "請選擇房型"}
            >
              {dormitoryList.map(roomName => (
                <Option key={roomName} value={roomName}>
                  {roomName}
                </Option>
              ))}
            </Select>
          </Col>

          {/* 👇 新增這一個 Col 區塊 👇 */}
            <Col span={7}>
              <label style={{ fontWeight: 'bold', color: '#fa8c16' }}>👶 未成年管理費:</label>
              <div style={{ marginTop: 8 }}>
                <Switch 
                  checked={request.NeedGuardianFee} 
                  onChange={checked => setRequest({ ...request, NeedGuardianFee: checked })}
                  checkedChildren="需要" 
                  unCheckedChildren="不需要"
                />
                <span style={{ marginLeft: 10, color: '#888' }}>
                  {request.NeedGuardianFee ? '已啟用' : '未啟用'}
                </span>
              </div>
            </Col>
          {/* 👆 新增結束 👆 */}

          {/* 👇 新增這一個 Col 區塊 👇 */}
            <Col span={5}>
              <label style={{ fontWeight: 'bold', color: '#cf1322' }}>💰 學校折扣 (美金):</label>
              <InputNumber 
                value={request.SchoolDiscount} 
                onChange={v => setRequest({ ...request, SchoolDiscount: v || 0 })}
                style={{ width: '100%', marginTop: 8 }}
                min={0}
                prefix="$"
                placeholder="請輸入折扣金額"
              />
            </Col>
          {/* 👆 新增結束 👆 */}
        </Row>

        {sheetOptions.length > 0 && (
            <Row gutter={16} style={{ marginTop: 20 }}>
                <Col span={8}>
                    <label>📄 報價頁籤:</label>
                    <Select 
                        value={request.SheetName || null} 
                        onChange={v => setRequest({ ...request, SheetName: v })}
                        style={{ width: '100%', marginTop: 8 }}
                        loading={loadingSheets}
                    >
                        {Array.isArray(sheetOptions) && sheetOptions.map(sheet => (
                        <Option key={sheet} value={sheet}>{sheet}</Option>
                        ))}
                    </Select>
                </Col>
            </Row>
        )}

        <div style={{ marginTop: 24, textAlign: 'right' }}>
          <Button 
            type="primary" 
            icon={<CalculatorOutlined />} 
            onClick={calculate}
            loading={loading}
            size="large"
          >
            計算報價
          </Button>

          {/* 👇 查飛機票按鈕加在這裡 👇 */}
          <Button 
            type="default"
            size="large"
            style={{ marginLeft: 12 }}  // 👈 加這行
            onClick={() => setShowFlightSearch(true)}
          >
            ✈️ 查飛機票
          </Button>
        </div>
      </Card>

      {quote && (
        <>
          <Card 
            title="🎓 課程費用項目" 
            style={{ marginTop: 24, borderColor: '#91caff' }} 
            styles={{ header: { backgroundColor: '#e6f7ff' } }} 

            // title={<span><DollarOutlined 
            //   style={{color:'#1890ff', marginRight:8}} 
            //   styles={{ header: { backgroundColor: '#e6f7ff' } }}/>課程費用</span>}
          >
            <Table 
              columns={courseFeeColumns} 
              dataSource={quote.courseFees} 
              pagination={false}
              size="middle"
              bordered
              rowKey="key" 
              scroll={{ x: 'max-content' }}  //{/* 👈 新增這行，讓表格超出寬度時可水平滑動 */}
            />
            
            <div style={{ marginTop: 16 }}>
              <Button type="dashed" onClick={addCourseFeeRow}>
                + 新增報價項目 (如: 代辦折扣、獎學金)
              </Button>
            </div>
            
            <Row gutter={16} style={{ marginTop: 16, textAlign: 'right' }}>
              <Col span={24}>
                <h3 style={{ margin: 0 }}>
                  {/* 👇 改成 currentTotalUSD 👇 */}
                  美金合計: <span style={{ color: '#cf1322' }}>US${currentTotalUSD.toLocaleString()}</span>
                </h3>
                <p style={{ color: '#888', margin: '4px 0' }}>
                  美金：台幣＝1：{currentRates?.usd?.sellRate} (報價當日匯率)
                </p>
                <h3 style={{ margin: 0, color: '#1890ff' }}>
                  {/* 👇 改成 currentTotalNTD 👇 */}
                  台幣換算: NT${currentTotalNTD.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                </h3>

              </Col>
            </Row>
          </Card>

          
          <Card 
              title="🇵🇭 當地雜費項目（披索）" 
              style={{ marginTop: 24, borderColor: '#91caff' }} 
              styles={{ header: { backgroundColor: '#e6f7ff' } }} 
            >
              <Table 
                columns={localFeeColumns} 
                dataSource={quote.localFees || []}
                pagination={false}
                size="middle"
                bordered
                rowKey="key"
                scroll={{ x: 'max-content' }}  //{/* 👈 新增這行，讓表格超出寬度時可水平滑動 */}
              />

              <div style={{ marginTop: 16 }}>
                <Button type="dashed" onClick={addLocalFeeRow}>
                  + 新增當地雜費項目
                </Button>
              </div>

              <Row gutter={16} style={{ marginTop: 16, textAlign: 'right' }}>
                <Col span={24}>
                  <h3 style={{ margin: 0 }}>
                    披索合計: <span style={{ color: '#cf1322' }}>₱{currentLocalTotalPeso.toLocaleString()} </span>
                  </h3>
                  <p style={{ color: '#888', margin: '4px 0' }}>
                    披索：台幣 = 1：{pesoToNT}
                  </p>
                  <h3 style={{ margin: 0, color: '#1890ff' }}>
                    台幣換算: NT${currentLocalTotalNTD.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                  </h3>
                </Col>
              </Row>
            </Card>

            {/* 👇 第三個 Table：機票簽證保險 👇 */}
          <Card 
            title="✈️ 其他費用（機票/簽證/保險）" 
            style={{ marginTop: 24, borderColor: '#91caff' }} 
                      styles={{ header: { backgroundColor: '#e6f7ff' } }} 
          >
            <Table 
              columns={otherFeeColumns}
              dataSource={quote?.otherFees || []}
              pagination={false}
              size="middle"
              bordered
              rowKey="key"
              scroll={{ x: 'max-content' }}  //{/* 👈 新增這行，讓表格超出寬度時可水平滑動 */}
            />
            
            <div style={{ marginTop: 16 }}>
              <Button 
                type="dashed" 
                onClick={addOtherFeeRow}
                icon={<PlusOutlined />}
              >
                + 新增其他費用
              </Button>
            </div>
            
            <Row gutter={16} style={{ marginTop: 16, textAlign: 'right' }}>
              <Col span={24}>
                <h3 style={{ margin: 0, color: '#1890ff' }}>
                  台幣合計：NT${currentOtherTotalNTD.toLocaleString()}
                </h3>
              </Col>
            </Row>
          </Card>

          {/* 👇 總計區塊 👇 */}
            <Card 
              title="💰 報價總計"
              style={{ marginTop: 24, borderColor: '#ff4d4f' }}
              styles={{ header: { backgroundColor: '#fff1f0' } }}
            >
              <Row gutter={16}>
                
                {/* 學費 */}
                <Col xs={24} md={8} style={{ textAlign: 'center', marginBottom: 16 }}>
                  <div style={{ 
                    padding: '16px', 
                    background: '#e6f7ff', 
                    borderRadius: 8,
                    border: '1px solid #91caff'
                  }}>
                    <div style={{ color: '#888', fontSize: 13, marginBottom: 4 }}>
                      🎓 課程費用
                    </div>
                    <div style={{ fontSize: 16, color: '#1890ff' }}>
                      US${currentTotalUSD.toLocaleString()}
                    </div>
                    <div style={{ fontSize: 20, fontWeight: 'bold', color: '#1890ff', marginTop: 4 }}>
                      NT${currentTotalNTD.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                    </div>
                  </div>
                </Col>

                {/* 當地雜費 */}
                <Col xs={24} md={8} style={{ textAlign: 'center', marginBottom: 16 }}>
                  <div style={{ 
                    padding: '16px', 
                    background: '#fffbe6', 
                    borderRadius: 8,
                    border: '1px solid #ffe58f'
                  }}>
                    <div style={{ color: '#888', fontSize: 13, marginBottom: 4 }}>
                      🇵🇭 當地雜費
                    </div>
                    <div style={{ fontSize: 16, color: '#fa8c16' }}>
                      ₱{currentLocalTotalPeso.toLocaleString()}
                    </div>
                    <div style={{ fontSize: 20, fontWeight: 'bold', color: '#fa8c16', marginTop: 4 }}>
                      NT${currentLocalTotalNTD.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                    </div>
                  </div>
                </Col>

                {/* 機票簽證 */}
                <Col xs={24} md={8} style={{ textAlign: 'center', marginBottom: 16 }}>
                  <div style={{ 
                    padding: '16px', 
                    background: '#f6ffed', 
                    borderRadius: 8,
                    border: '1px solid #b7eb8f'
                  }}>
                    <div style={{ color: '#888', fontSize: 13, marginBottom: 4 }}>
                      ✈️ 機票/簽證/保險
                    </div>
                    <div style={{ fontSize: 16, color: '#52c41a' }}>
                      &nbsp;
                    </div>
                    <div style={{ fontSize: 20, fontWeight: 'bold', color: '#52c41a', marginTop: 4 }}>
                      NT${currentOtherTotalNTD.toLocaleString()}
                    </div>
                  </div>
                </Col>
              </Row>

              {/* 分隔線 */}
              <div style={{ 
                margin: '20px 0', 
                borderBottom: '2px dashed #ff4d4f' 
              }} />

              {/* 總計 */}
              <Row>
                <Col span={24} style={{ textAlign: 'right' }}>
                  <div style={{ marginBottom: 8, color: '#888', fontSize: 14 }}>
                    NT${currentTotalNTD.toLocaleString(undefined, { maximumFractionDigits: 0 })} 
                    + NT${currentLocalTotalNTD.toLocaleString(undefined, { maximumFractionDigits: 0 })} 
                    + NT${currentOtherTotalNTD.toLocaleString()}
                  </div>
                  <div style={{ 
                    fontSize: 32, 
                    fontWeight: 'bold', 
                    color: '#cf1322',
                  }}>
                    🧾 總計：NT${(
                      currentTotalNTD + 
                      currentLocalTotalNTD + 
                      currentOtherTotalNTD
                    ).toLocaleString(undefined, { maximumFractionDigits: 0 })}
                  </div>
                  <div style={{ color: '#888', fontSize: 12, marginTop: 4 }}>
                    ＊匯率：US$1 = NT${request.ExchangeRate}　₱1 = NT${pesoToNT}
                  </div>
                </Col>
              </Row>
            </Card>
            {/* 👆 總計區塊結束 👆 */}



            {/* 下面再放「匯出 Excel」那個 Card */}
          <Card title="功能按鈕" style={{ marginTop: 24 }}>
            <div>
              <Button 
                type="primary" 
                icon={<FileExcelOutlined />} 
                onClick={handleSendQuoteData}
                size="large"
                style={{ marginTop: 16 }}
              >
                匯出 Excel
              </Button>

            
              
              {/* <Button onClick={handleSaveJson} type="default" iicon={<DownloadOutlined />}>
                存為 JSON
              </Button>

              <h4>從 JSON 匯出 Excel</h4>
                <input 
                  type="file" 
                  accept=".json" 
                  onChange={handleFileSelect}
                  style={{ marginBottom: 12, width: '100%' }}
                />
                <Button 
                  type="primary" 
                  onClick={handleSendQuoteData}
                  disabled={!selectedFile}
                  icon={<UploadOutlined />}
                  size="large"
                >
                  上傳 JSON → 下載 Excel
                </Button> */}
            </div>

            

            {/* <Button onClick={handleExportExcel} type="primary">匯出 Excel</Button> */}
          </Card>
        </>
      )}

      {/* 👇 步驟4：飛機票 Modal 加在這裡（匯出Card後面）👇 */}
    <Modal 
      title="✈️ 即時查詢飛機票 (TPE-CEB)" 
      open={showFlightSearch}
      onCancel={() => setShowFlightSearch(false)}
      width={1000}
      footer={null}  // 不顯示預設按鈕
    >
      <FlightSearch 
        defaultDeparture="TPE"
        defaultArrival="CEB"
        defaultOutboundDate={request.StartDate}
        defaultReturnDate={request.EndDate}
        defaultAdults={1}
        onSelectFlight={(price, description) => {
          // 自動填入機票表第一行
          setQuote(prev => {
            const newFees = [...(prev?.otherFees || [])];
            newFees[0] = {
              ...newFees[0],
              item: description,        // "華航 CI721"
              unitPrice: price,         // 12800
              amount: price,
              remark: `即時查詢 ${new Date().toLocaleString('zh-TW')}`
            };
            return { ...prev, otherFees: newFees };
          });
          setShowFlightSearch(false);
          message.success(`✅ 已套用 ${description} NT$${price}`);
        }}
      />
    </Modal>
    {/* 👆 Modal 加完 👆 */}


    </div>
  );
}

export default App;
