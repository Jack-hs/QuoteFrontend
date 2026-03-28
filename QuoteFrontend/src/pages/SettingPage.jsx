// src/pages/SettingPage.jsx
import React, { useState, useEffect } from "react";
import {
  Form,
  Input,
  InputNumber,
  Select,
  Row,
  Col,
  Card,
  Button,
  Table,
  Space,
  Popconfirm,
  Divider,
  Collapse,
} from "antd";
import { FileExcelOutlined } from "@ant-design/icons";
import { PlusOutlined, DeleteOutlined, EditOutlined, CheckOutlined, CloseOutlined } from "@ant-design/icons";
import axios from "axios";
import { message } from "antd"; // 👈 這一行要加，否則 message.success 會報錯

const { Panel } = Collapse;
const { Option } = Select;

// 這是「整份」設定的初始結構，也可以從後端 GET 拿來放到 initialData
const initialSettingData = {
  setting: {
    schoolLocation: "宿霧",
    flightDest: "宿霧",
  },
  courseFee: {
    registrationfee: "120",
    registrationfeeInclusive: "no",
    summerSurcharge: "40",
    summerSurchargeStartTime: "7/1",
    summerSurchargeEndTime: "8/31",
    discountInclusive: "0",
    discount: "0.05",
    fixedDiscount: "0",
    legalGuardianFee: "0",
  },
  logo: {
    picLocationRow: "7",
    picLocationCol: "2",
    picLocationLeft: "0",
    picLocationTop: "5",
    picwidth: "330",
    picheight: "110",
  },
  localFee: [
    { code: "1", DocItem: "Special Study Permit SSP", Item: "SSP學生簽證", content: "菲律賓學生就讀許可證", times: "1次", price: "7800", "remark " : "效期6個月" },
    { code: "2", DocItem: "SSP ACR ECARD", Item: "SSP E-CARD", content: "菲律賓學生就讀許可證", times: "1次", price: "4500", "remark " : "-" },
    // 這裡先保留示例，你再依你實際列表進來
  ],
  // 課程群組清單
  courseGroupList: ["IELTS", "ESL & SPEAKING", "TOEIC, BUSINESS", "SPECIAL PROGRAM", "JUNIOR & FAMILY", "PREMIUM"],
  courses: {
    IELTS: [
      { name: "IELTS INTENSIVE", code: "IELTS1", pricePerWeek: 1200, description: "密集 IELTS 課程" },
    ],
    "ESL & SPEAKING": [
      { name: "INTENSIVE ESL", code: "ESL1", pricePerWeek: 1030 },
    ],
  },
  room: [
    { Name: "Single", roomType: "on-campus", code: "Room1", pricePerWeek: 1400, description: "單人房" },
    { Name: "Double", roomType: "on-campus", code: "Room2", pricePerWeek: 970, description: "雙人房" },
  ],
  more4week: { more4week1: "0", more4week2: "0.25", more4week3: "0.5", more4week4: "0.75" },
  less4week: { less4week1: "0.45", less4week2: "0.65", less4week3: "0.85" },
};

const SettingPage = () => {
  const [form] = Form.useForm();
  const [data, setData] = useState(initialSettingData);   // 模擬你從後端拿來的設定
  // 再加一個：用來存「你剛上傳的檔案資料」
const [loadedSettingData, setLoadedSettingData] = useState(null);
  const [editingRow, setEditingRow] = useState(null);     // 哪一筆正在編輯
  const [localFeeEditingKey, setLocalFeeEditingKey] = useState("");
  const [courseEditingKey, setCourseEditingKey] = useState("");
  const [roomEditingKey, setRoomEditingKey] = useState("");
  const [loading, setLoading] = useState(false); // 👈 你自己 add 一個 loading，這樣下面那行就不会错

  // 你自己的 API 端點，從 .env 拿也可以
  const apiUrl = import.meta.env.VITE_API_URL || "https://localhost:7080"; // 👈 這一行要加，否則 handleSaveAll 會炸
  // 假裝你有 API：保存整份設定
  const handleSaveAll = async () => {
    try {
      await axios.post(apiUrl + "/api/setting/save", data); // 假設你有這支 API
      message.success("設定已儲存");
    } catch (err) {
      message.error("儲存失敗：" + err.message);
    }
  };

  // 1. 當你「從檔案讀到一筆設定」後，更新 data
  const handleDataLoaded = (jsonData) => {
    console.log("🟢 使用檔案讀到的設定資料", jsonData);

    // 你可以用 Console 確認 jsonData 結構，再確認要不要全部取代
    setData(jsonData);

    // 也可以只取代某個部分，例如：
    // setData(prev => ({ ...prev, setting: jsonData.setting }));
    // setData(prev => ({ ...prev, courseFee: jsonData.courseFee }));
  };

  //#region 以下是完整的修改版本，新增了「選擇學校」功能：
  // 在 useState 宣告區新增
  const [schoolList, setSchoolList] = useState([]);
  const [selectedSchool, setSelectedSchool] = useState('');
  const [loadingSchools, setLoadingSchools] = useState(false);

    // 在 useEffect 中加入載入學校清單
  useEffect(() => {
    // 載入伺服器上的學校清單
    const loadSchoolList = async () => {
      try {
        setLoadingSchools(true);
        const response = await axios.get(`${apiUrl}/api/quote/school-list`);
        setSchoolList(response.data);
      } catch (error) {
        message.error('載入學校清單失敗');
        console.error('載入學校失敗:', error);
      } finally {
        setLoadingSchools(false);
      }
    };
    loadSchoolList();
  }, []);

  // 新增：載入指定學校的設定
  const loadSchoolSetting = async (schoolName) => {
    try {
      //const response = await axios.get(`${apiUrl}/api/quote/setting/${schoolName}`);
      // 在 SettingPage.jsx 的 loadSchoolSetting 函式中
      const response = await axios.get(`${apiUrl}/api/setting/school/${schoolName}`);

      const schoolData = response.data;
      setData(schoolData);
      setSelectedSchool(schoolName);
      // 👈 新增：重置課程群組到第一個可用群組
    const firstCourseGroup = Object.keys(schoolData.courses || {})[0];
    setCourseGroup(firstCourseGroup || '');
      message.success(`已載入 ${schoolName} 的設定`);
    } catch (error) {
      message.error(`載入 ${schoolName} 設定失敗`);
      console.error(error);
    }
  };

  // // 匯出 JSON 的檔名也用學校名稱
  const handleExportJSON = async () => {
    if (!selectedSchool) {
      message.warning('請先選擇學校');
      return;
    }

    try {
      // 👈 關鍵：先從 Form 抓取最新值，確保包含所有即時修改
      const formValues = form.getFieldsValue();
      const latestData = buildDataFromFormValues(formValues); // 自己寫的轉換函式

      // 1. 先儲存到伺服器
      const response = await axios.post(
        `${apiUrl}/api/setting/school/${selectedSchool}`, 
        latestData  
      );
      
      message.success(`已儲存到伺服器：${response.data.fileName}`);

      // 2. 再下載到本地
      const jsonString = JSON.stringify(latestData, null, 2);
      const blob = new Blob([jsonString], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${selectedSchool}_setting.json`;
      document.body.appendChild(a); // 確保在 DOM 中
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      message.success('已下載到本地電腦');
    } catch (error) {
      console.error('儲存失敗:', error);
      message.error('儲存失敗，請檢查 Console');
    }
  };

  // 👈 新增：從 Form 值建構完整 data 物件
  const buildDataFromFormValues = (formValues) => {
    return {
      ...data, // 保留 courses, localFee, room 等複雜結構
      setting: {
        schoolLocation: formValues.setting__schoolLocation || '',
        flightDest: formValues.setting__flightDest || ''
      },
      courseFee: {
        registrationfee: String(formValues.courseFee__registrationfee || 0),
        summerSurcharge: String(formValues.courseFee__summerSurcharge || 0),
        discount: String(formValues.courseFee__discount || 0),
        // ... 其他 courseFee 欄位
        legalGuardianFee: String(formValues.courseFee__legalGuardianFee || 0),
        
        // 👇 將缺少欄位加入回傳物件
        registrationfeeInclusive: formValues.courseFee__registrationfeeInclusive || "no",
        summerSurchargeStartTime: formValues.courseFee__summerSurchargeStartTime || "",
        summerSurchargeEndTime: formValues.courseFee__summerSurchargeEndTime || "",
        discountInclusive: String(formValues.courseFee__discountInclusive || 0),
        fixedDiscount: String(formValues.courseFee__fixedDiscount || 0),
      },
      logo: {
        picLocationRow: String(formValues.logo__picLocationRow || 0),
        // ... 其他 logo 欄位
      },
      more4week: {
        more4week1: String(formValues.more4week1 || 0),
        // ... 其他 more4week
      },
      less4week: {
        less4week1: String(formValues.less4week1 || 0),
        // ... 其他 less4week
      }
    };
  };
  //#endregion

  // =================== 一、表單基本欄位綁定 ===================

  useEffect(() => {
    // 把 data 轉成 Form 的欄位 layout
    const values = {};
    values.setting__schoolLocation = data.setting.schoolLocation;
    values.setting__flightDest = data.setting.flightDest;

    values.courseFee__registrationfee = Number(data.courseFee.registrationfee);
    values.courseFee__summerSurcharge = Number(data.courseFee.summerSurcharge);
    values.courseFee__discount = Number(data.courseFee.discount);
    values.courseFee__fixedDiscount = Number(data.courseFee.fixedDiscount);
    values.courseFee__legalGuardianFee = Number(data.courseFee.legalGuardianFee);
    values.courseFee__summerSurchargeStartTime = data.courseFee.summerSurchargeStartTime;
    values.courseFee__summerSurchargeEndTime = data.courseFee.summerSurchargeEndTime;

    values.courseFee__registrationfeeInclusive = data.courseFee.registrationfeeInclusive;
    values.courseFee__discountInclusive = data.courseFee.discountInclusive;

    values.logo__picLocationRow = Number(data.logo.picLocationRow);
    values.logo__picLocationCol = Number(data.logo.picLocationCol);
    values.logo__picLocationLeft = Number(data.logo.picLocationLeft);
    values.logo__picLocationTop = Number(data.logo.picLocationTop);
    values.logo__picwidth = Number(data.logo.picwidth);
    values.logo__picheight = Number(data.logo.picheight);

    // more4week / less4week
    values.more4week1 = Number(data.more4week.more4week1);
    values.more4week2 = Number(data.more4week.more4week2);
    values.more4week3 = Number(data.more4week.more4week3);
    values.more4week4 = Number(data.more4week.more4week4);

    values.less4week1 = Number(data.less4week.less4week1);
    values.less4week2 = Number(data.less4week.less4week2);
    values.less4week3 = Number(data.less4week.less4week3);

    form.setFieldsValue(values);
  }, [data, form]);

  const handleFormChange = (changedValues, allValues) => {
    const newData = { ...data };

    // 基本設定
    newData.setting.schoolLocation = allValues.setting__schoolLocation || "台北";
    newData.setting.flightDest = allValues.setting__flightDest || "台北";

    // 課程費率
    newData.courseFee.registrationfee = String(allValues.courseFee__registrationfee || 0);
    newData.courseFee.summerSurcharge = String(allValues.courseFee__summerSurcharge || 0);
    newData.courseFee.discount = String(allValues.courseFee__discount || 0);
    newData.courseFee.fixedDiscount = String(allValues.courseFee__fixedDiscount || 0);
    newData.courseFee.legalGuardianFee = String(allValues.courseFee__legalGuardianFee || 0);
    newData.courseFee.summerSurchargeStartTime = allValues.courseFee__summerSurchargeStartTime || "7/1";
    newData.courseFee.summerSurchargeEndTime = allValues.courseFee__summerSurchargeEndTime || "8/31";
    
    newData.courseFee.registrationfeeInclusive = allValues.courseFee__registrationfeeInclusive || "no";
    newData.courseFee.discountInclusive = String(allValues.courseFee__discountInclusive || 0);

    // Logo 位置
    newData.logo.picLocationRow = String(allValues.logo__picLocationRow || 7);
    newData.logo.picLocationCol = String(allValues.logo__picLocationCol || 2);
    newData.logo.picLocationLeft = String(allValues.logo__picLocationLeft || 0);
    newData.logo.picLocationTop = String(allValues.logo__picLocationTop || 5);
    newData.logo.picwidth = String(allValues.logo__picwidth || 330);
    newData.logo.picheight = String(allValues.logo__picheight || 110);

    // more4week / less4week
    newData.more4week.more4week1 = String(allValues.more4week1 || 0);
    newData.more4week.more4week2 = String(allValues.more4week2 || 0);
    newData.more4week.more4week3 = String(allValues.more4week3 || 0);
    newData.more4week.more4week4 = String(allValues.more4week4 || 0);

    newData.less4week.less4week1 = String(allValues.less4week1 || 0);
    newData.less4week.less4week2 = String(allValues.less4week2 || 0);
    newData.less4week.less4week3 = String(allValues.less4week3 || 0);

    setData(newData);
  };


  // 在 SettingPage 元件裡加入 courseGroupListTable

//   // =================== CourseGroupList 表格 CRUD ===================
// const courseGroupListDataSource = data.courseGroupList.map((groupName, index) => ({
//   key: index,
//   groupName,
// }));

// const courseGroupColumns = [
//   {
//     title: "課程群組名稱",
//     dataIndex: "groupName",
//     key: "groupName",
//     width: 180,
//     render: (text, record) => {
//       const editable = groupEditingKey === record.key;
//       return editable ? (
//         <Input
//           defaultValue={text}
//           onChange={(e) => {
//             const newList = data.courseGroupList.map((g, i) =>
//               i === record.key ? e.target.value : g
//             );
//             setData({ ...data, courseGroupList: newList });
//           }}
//         />
//       ) : (
//         <span>{text}</span>
//       );
//     },
//   },
//   {
//     title: "操作",
//     key: "action",
//     width: 100,
//     render: (_, record) => {
//       const editable = groupEditingKey === record.key;
//       return editable ? (
//         <Space>
//           <Button
//             type="primary"
//             onClick={() => {
//               setGroupEditingKey(null);
//             }}
//             size="small"
//           >
//             儲存
//           </Button>
//           <Button
//             onClick={() => {
//               setGroupEditingKey(null);
//             }}
//             size="small"
//           >
//             取消
//           </Button>
//         </Space>
//       ) : (
//         <Space>
//           <Button
//             icon={<EditOutlined />}
//             onClick={() => {
//               setGroupEditingKey(record.key);
//             }}
//             size="small"
//           >
//             修改
//           </Button>
//           <Popconfirm
//             title="確定刪除這個群組？（含所有課程）"
//             onConfirm={() => {
//               const newList = data.courseGroupList.filter(
//                 (g, i) => i !== record.key
//               );
//               const newCourses = { ...data.courses };
//               delete newCourses[record.groupName]; // 刪除該群組課程
//               setData({ ...data, courseGroupList: newList, courses: newCourses });
//             }}
//           >
//             <Button
//               danger
//               icon={<DeleteOutlined />}
//               size="small"
//             >
//               刪除
//             </Button>
//           </Popconfirm>
//         </Space>
//       );
//     },
//   },
// ];

// const handleAddGroup = () => {
//   const newGroupName = "新課程群組 " + (data.courseGroupList.length + 1);
//   setData({
//     ...data,
//     courseGroupList: [...data.courseGroupList, newGroupName],
//     courses: {
//       ...data.courses,
//       [newGroupName]: [],
//     },
//   });
// };



  // =================== 二、localFee 表格 CRUD ===================

  const localFeeDataSource = (data.localFee || []).map((item) => ({
    ...item,
    key: item.code,
  }));

  const localFeeColumns = [
  { title: "編號", dataIndex: "code", key: "code", width: 80 },
  {
    title: "英文項目",
    dataIndex: "DocItem",
    key: "DocItem",
    width: 180,
    render: (text, record) => {
      const editable = localFeeEditingKey === record.code;
      return editable ? (
        <Input
          defaultValue={text}
          onChange={(e) => {
            const newLocalFee = data.localFee.map((item) =>
              item.code === record.code
                ? { ...item, DocItem: e.target.value }
                : item
            );
            setData({ ...data, localFee: newLocalFee });
          }}
        />
      ) : (
        <span>{text}</span>
      );
    },
  },
  {
    title: "中文項目",
    dataIndex: "Item",
    key: "Item",
    width: 180,
    render: (text, record) => {
      const editable = localFeeEditingKey === record.code;
      return editable ? (
        <Input
          defaultValue={text}
          onChange={(e) => {
            const newLocalFee = data.localFee.map((item) =>
              item.code === record.code
                ? { ...item, Item: e.target.value }
                : item
            );
            setData({ ...data, localFee: newLocalFee });
          }}
        />
      ) : (
        <span>{text}</span>
      );
    },
  },
  {
    title: "內容說明",
    dataIndex: "content",
    key: "content",
    width: 200,
    render: (text, record) => {
      const editable = localFeeEditingKey === record.code;
      return editable ? (
        <Input
          defaultValue={text}
          onChange={(e) => {
            const newLocalFee = data.localFee.map((item) =>
              item.code === record.code
                ? { ...item, content: e.target.value }
                : item
            );
            setData({ ...data, localFee: newLocalFee });
          }}
        />
      ) : (
        <span>{text}</span>
      );
    },
  },
  {
    title: "次數/週數",
    dataIndex: "times",
    key: "times",
    width: 100,
    render: (text, record) => {
      const editable = localFeeEditingKey === record.code;
      return editable ? (
        <Input
          defaultValue={text}
          onChange={(e) => {
            const newLocalFee = data.localFee.map((item) =>
              item.code === record.code
                ? { ...item, times: e.target.value }
                : item
            );
            setData({ ...data, localFee: newLocalFee });
          }}
        />
      ) : (
        <span>{text}</span>
      );
    },
  },
  {
    title: "價格 (披索)",
    dataIndex: "price",
    key: "price",
    width: 120,
    render: (text, record) => {
      const editable = localFeeEditingKey === record.code;
      return editable ? (
        <InputNumber
          defaultValue={Number(text) || 0}
          onChange={(v) => {
            const newLocalFee = data.localFee.map((item) =>
              item.code === record.code
                ? { ...item, price: String(v) }
                : item
            );
            setData({ ...data, localFee: newLocalFee });
          }}
          style={{ width: "100%" }}
          min={0}
        />
      ) : (
        <span>{text === "-" ? "-" : `₱${text}`}</span>
      );
    },
  },
  {
    title: "備註",
    dataIndex: "remark ",
    key: "remark ",
    width: 150,
    render: (text, record) => {
      const editable = localFeeEditingKey === record.code;
      return editable ? (
        <Input
          defaultValue={text || ""}
          onChange={(e) => {
            const newLocalFee = data.localFee.map((item) =>
              item.code === record.code
                ? { ...item, "remark ": e.target.value }
                : item
            );
            setData({ ...data, localFee: newLocalFee });
          }}
        />
      ) : (
        <span>{text}</span>
      );
    },
  },
  {
    title: "操作",
    key: "action",
    width: 120,
    render: (_, record) => {
      const editable = localFeeEditingKey === record.code;
      return editable ? (
        <Space>
          <Button
            type="primary"
            icon={<CheckOutlined />}
            onClick={() => {
              setLocalFeeEditingKey("");
              // 儲存時你已經在 Input/onChange 裡寫回 data，所以不用再做
            }}
            size="small"
          >
            儲存
          </Button>
          <Button
            onClick={() => {
              setLocalFeeEditingKey("");
            }}
            size="small"
          >
            取消
          </Button>
        </Space>
      ) : (
        <Space>
          <Button
            icon={<EditOutlined />}
            onClick={() => {
              setLocalFeeEditingKey(record.code);
            }}
            size="small"
          >
            修改
          </Button>
          <Popconfirm
            title="確定刪除這筆？"
            onConfirm={() => {
              const newLocalFee = data.localFee.filter(
                (item) => item.code !== record.code
              );
              setData({ ...data, localFee: newLocalFee });
            }}
          >
            <Button
              danger
              icon={<DeleteOutlined />}
              size="small"
            >
              刪除
            </Button>
          </Popconfirm>
        </Space>
      );
    },
  },
];


 

  const handleAddLocalFee = () => {
    const newCode = (parseInt(localFeeDataSource.slice(-1)[0]?.code || "0") + 1).toString();
    const newLocalFee = {
      code: newCode,
      DocItem: "New Doc",
      Item: "新項目",
      content: "說明",
      times: "1次",
      price: "0",
      "remark ": "",
    };
    setData({ ...data, localFee: [...data.localFee, newLocalFee] });
    setLocalFeeEditingKey(newCode);
  };


  // =================== 三、courses 表格（可先做 IELTS 這個 group）===================

  const [courseGroup, setCourseGroup] = useState("IELTS");

  const currentCourses = data.courses[courseGroup] || [];

  const courseColumns = [
  {
    title: "名稱",
    dataIndex: "name",
    key: "name",
    width: 180,
    render: (text, record) => {
      const editable = courseEditingKey === record.code;
      return editable ? (
        <Input
          defaultValue={text}
          onChange={(e) => {
            const group = data.courses[courseGroup] || [];
            const newGroup = group.map((item) =>
              item.code === record.code
                ? { ...item, name: e.target.value }
                : item
            );
            setData({
              ...data,
              courses: {
                ...data.courses,
                [courseGroup]: newGroup,
              },
            });
          }}
        />
      ) : (
        <span>{text}</span>
      );
    },
  },
  {
    title: "課程代碼",
    dataIndex: "code",
    key: "code",
    width: 100,
    render: (text, record) => {
      const editable = courseEditingKey === record.code;
      return editable ? (
        <Input
          defaultValue={text}
          onChange={(e) => {
            const group = data.courses[courseGroup] || [];
            const newGroup = group.map((item) =>
              item.code === record.code
                ? { ...item, code: e.target.value }
                : item
            );
            setData({
              ...data,
              courses: {
                ...data.courses,
                [courseGroup]: newGroup,
              },
            });
          }}
        />
      ) : (
        <span>{text}</span>
      );
    },
  },
  {
    title: "每週價格 (美金)",
    dataIndex: "pricePerWeek",
    key: "pricePerWeek",
    width: 120,
    render: (price, record) => {
      const editable = courseEditingKey === record.code;
      return editable ? (
        <InputNumber
          defaultValue={price}
          onChange={(v) => {
            const group = data.courses[courseGroup] || [];
            const newGroup = group.map((item) =>
              item.code === record.code
                ? { ...item, pricePerWeek: v }
                : item
            );
            setData({
              ...data,
              courses: {
                ...data.courses,
                [courseGroup]: newGroup,
              },
            });
          }}
          style={{ width: "100%" }}
          precision={2}
        />
      ) : (
        <span>US${Number(price).toFixed(2)}</span>
      );
    },
  },
  {
    title: "說明",
    dataIndex: "description",
    key: "description",
    width: 200,
    render: (text, record) => {
      const editable = courseEditingKey === record.code;
      return editable ? (
        <Input
          defaultValue={text || ""}
          onChange={(e) => {
            const group = data.courses[courseGroup] || [];
            const newGroup = group.map((item) =>
              item.code === record.code
                ? { ...item, description: e.target.value }
                : item
            );
            setData({
              ...data,
              courses: {
                ...data.courses,
                [courseGroup]: newGroup,
              },
            });
          }}
        />
      ) : (
        <span>{text}</span>
      );
    },
  },
  {
    title: "操作",
    key: "action",
    width: 100,
    render: (_, record) => {
      const editable = courseEditingKey === record.code;
      return editable ? (
        <Space>
          <Button
            type="primary"
            onClick={() => {
              setCourseEditingKey("");
            }}
            size="small"
          >
            儲存
          </Button>
          <Button
            onClick={() => {
              setCourseEditingKey("");
            }}
            size="small"
          >
            取消
          </Button>
        </Space>
      ) : (
        <Space>
          <Button
            icon={<EditOutlined />}
            onClick={() => {
              setCourseEditingKey(record.code);
            }}
            size="small"
          >
            修改
          </Button>
          <Popconfirm
            title="確定刪除這個課程？"
            onConfirm={() => {
              const group = data.courses[courseGroup] || [];
              const newGroup = group.filter(
                (item) => item.code !== record.code
              );
              setData({
                ...data,
                courses: {
                  ...data.courses,
                  [courseGroup]: newGroup,
                },
              });
            }}
          >
            <Button
              danger
              icon={<DeleteOutlined />}
              size="small"
            >
              刪除
            </Button>
          </Popconfirm>
        </Space>
      );
    },
  },
];

  const handleAddCourse = () => {
    const maxCodeNum = currentCourses.reduce((max, item) => {
      const num = Number(item.code.match(/\d+$/)?.[0] || 0);
      return Math.max(max, num);
    }, 0);

    const newCode = courseGroup === "IELTS" ? `IELTS${maxCodeNum + 1}` :
                    courseGroup === "ESL & SPEAKING" ? `ESL${maxCodeNum + 1}` :
                    courseGroup === "TOEIC, BUSINESS" ? `TOEIC${maxCodeNum + 1}` :
                    `${courseGroup.slice(0,1)}${maxCodeNum + 1}`;

    const newCourse = {
      name: "新課程名稱",
      code: newCode,
      pricePerWeek: 1000,
      description: "請填寫說明",
    };
    const group = data.courses[courseGroup] || [];
    setData({
      ...data,
      courses: {
        ...data.courses,
        [courseGroup]: [...group, newCourse],
      },
    });
    setCourseEditingKey(newCode);
  };

  // 你可以用類似方式，再做 `room` 的表格 CRUD，篇幅關係我先省略，原理一樣。


  
  // =================== 三、rooms 表格（可先做 IELTS 這個 group）===================
  // 2. 在這裡寫 roomColumns（在 return 之前，不是 return 裡面）
  const roomColumns = [
    {
      title: "房型名稱",
      dataIndex: "Name",
      key: "Name",
      width: 120,
      render: (text, record) => {
        const editable = roomEditingKey === record.code;
        return editable ? (
          <Input
            defaultValue={text}
            onChange={(e) => {
              const newRoom = (data.room || []).map((item) =>
                item.code === record.code
                  ? { ...item, Name: e.target.value }
                  : item
              );
              setData({ ...data, room: newRoom });
            }}
          />
        ) : (
          <span>{text}</span>
        );
      },
    },
    {
      title: "房型類型",
      dataIndex: "roomType",
      key: "roomType",
      width: 120,
      render: (text, record) => {
        const editable = roomEditingKey === record.code;
        return editable ? (
          <Select
            value={text}
            onChange={(v) => {
              const newRoom = data.room.map((item) =>
                item.code === record.code
                  ? { ...item, roomType: v }
                  : item
              );
              setData({ ...data, room: newRoom });
            }}
            style={{ width: "100%" }}
          >
            <Option value="on-campus">on-campus</Option>
            <Option value="off-campus">off-campus</Option>
          </Select>
        ) : (
          <span>{text}</span>
        );
      },
    },
    {
      title: "代碼",
      dataIndex: "code",
      key: "code",
      width: 100,
      render: (text, record) => {
        const editable = roomEditingKey === record.code;
        return editable ? (
          <Input
            defaultValue={text}
            onChange={(e) => {
              const newRoom = data.room.map((item) =>
                item.code === record.code
                  ? { ...item, code: e.target.value }
                  : item
              );
              setData({ ...data, room: newRoom });
            }}
          />
        ) : (
          <span>{text}</span>
        );
      },
    },
    {
      title: "每週價格 (美金)",
      dataIndex: "pricePerWeek",
      key: "pricePerWeek",
      width: 120,
      render: (price, record) => {
        const editable = roomEditingKey === record.code;
        return editable ? (
          <InputNumber
            defaultValue={price}
            onChange={(v) => {
              const newRoom = data.room.map((item) =>
                item.code === record.code
                  ? { ...item, pricePerWeek: v }
                  : item
              );
              setData({ ...data, room: newRoom });
            }}
            style={{ width: "100%" }}
            precision={2}
          />
        ) : (
          <span>US${price}</span>
        );
      },
    },
    {
      title: "說明",
      dataIndex: "description",
      key: "description",
      width: 180,
      render: (text, record) => {
        const editable = roomEditingKey === record.code;
        return editable ? (
          <Input
            defaultValue={text || ""}
            onChange={(e) => {
              const newRoom = data.room.map((item) =>
                item.code === record.code
                  ? { ...item, description: e.target.value }
                  : item
              );
              setData({ ...data, room: newRoom });
            }}
          />
        ) : (
          <span>{text}</span>
        );
      },
    },
    {
      title: "操作",
      key: "action",
      width: 100,
      render: (_, record) => {
        const editable = roomEditingKey === record.code;
        return editable ? (
          <Space>
            <Button
              type="primary"
              icon={<CheckOutlined />}
              onClick={() => {
                setRoomEditingKey("");
              }}
              size="small"
            >
              儲存
            </Button>
            <Button
              onClick={() => {
                setRoomEditingKey("");
              }}
              size="small"
            >
              取消
            </Button>
          </Space>
        ) : (
          <Space>
            <Button
              icon={<EditOutlined />}
              onClick={() => {
                setRoomEditingKey(record.code);
              }}
              size="small"
            >
              修改
            </Button>
            <Popconfirm
              title="確定刪除這間房型？"
              onConfirm={() => {
                const newRoom = data.room.filter(
                  (r) => r.code !== record.code
                );
                setData({ ...data, room: newRoom });
              }}
            >
              <Button
                danger
                icon={<DeleteOutlined />}
                size="small"
              >
                刪除
              </Button>
            </Popconfirm>
          </Space>
        );
      },
    },
  ];

  

  


  return (
    <div style={{ padding: 24 }}>
      <h1>🔧 系統設定頁面</h1>
      {/* =================== 新增：選擇學校功能 =================== */}
      <div style={{ marginBottom: 24, padding: 16, backgroundColor: '#f0f8ff', borderRadius: 8, border: '1px solid #1890ff' }}>
        <Row gutter={16} align="middle">
          <Col span={24}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <label style={{ fontWeight: 'bold', whiteSpace: 'nowrap' }}>
                選擇學校設定：
              </label>
              <Select
                loading={loadingSchools}
                value={selectedSchool || undefined}
                onChange={loadSchoolSetting}
                style={{ width: 300 }}
                placeholder="選擇學校載入設定..."
                allowClear
                clearIcon={<CloseOutlined />}
                onClear={() => {
                  setSelectedSchool('');
                  setData({ ...initialSettingData });
                }}
              >
                {schoolList.map((school) => (
                  <Option key={school} value={school}>
                    {school}
                  </Option>
                ))}
              </Select>
              {/* 👈 文字直接跟在 Select 旁邊 */}
              <span style={{ 
                color: selectedSchool ? '#1890ff' : '#666', 
                fontSize: '14px',
                fontWeight: selectedSchool ? 500 : 'normal'
              }}>
                {selectedSchool ? `正在編輯 ${selectedSchool}` : '請選擇學校'}
              </span>
            </div>
          </Col>
          {/* <Col pan={16}>
            <span style={{ color: '#666', fontSize: '14px' }}>
              {selectedSchool ? `目前編輯：${selectedSchool}` : '請選擇學校'}
            </span>
          </Col> */}
          </Row>
          <div style={{ margin: '5px 0', borderBottom: '1px solid #f0f0f0' }}></div>
          <Row>
          <Col flex="auto">
            <div style={{ textAlign: 'left',display: 'flex', alignItems: 'center', gap: 12 }}>
              <label style={{ fontWeight: 'bold', whiteSpace: 'nowrap' }}>
                選擇儲存模式：
              </label>
              <Button 
                type="primary" 
                //icon={<FileExcelOutlined />}
                onClick={handleExportJSON}
                loading={loading}  // 可選：加個 loading 狀態
                disabled={!selectedSchool}
              >
                💾 儲存並匯出 {selectedSchool && `${selectedSchool}_setting.json`}
              </Button>
              <Button 
                style={{ marginLeft: 8 }}
                onClick={() => message.info('僅下載，不儲存到伺服器')}
              >
                只下載此電腦中
              </Button>
            </div>
          </Col>
        </Row>
      </div>


      <Divider />

      <Form
        form={form}
        layout="vertical"
        onValuesChange={handleFormChange}
      >
        <Collapse
          defaultActiveKey={[
            "basic",
            "courseFee",
            "logo",
            "more4week",
            "less4week",
            "localFee",
            "courses",
            "room",
          ]}
          items={[
            {
              key: "basic",
              label: "基本設定",
              children: (
                <Row gutter={16}>
                  <Col span={12}>
                    <Form.Item name="setting__schoolLocation" label="學校所在城市">
                      <Input placeholder="如：宿霧" />
                    </Form.Item>
                  </Col>
                  <Col span={12}>
                    <Form.Item name="setting__flightDest" label="機票目的地">
                      <Input placeholder="如：宿霧" />
                    </Form.Item>
                  </Col>
                </Row>
              ),
            },
            {
              key: "courseFee",
              label: "課程費用與折扣",
              children: (
                 <Row gutter={[16, 16]}> {/* 加上直向的 gutter 讓換行好看 */}
                  <Col span={6}>
                    <Form.Item name="courseFee__registrationfee" label="註冊費 (美金)">
                      <InputNumber style={{ width: "100%" }} prefix="US$" />
                    </Form.Item>
                  </Col>
                  {/* 👇 新增：註冊費是否內含 */}
                  <Col span={6}>
                    <Form.Item name="courseFee__registrationfeeInclusive" label="註冊費是否內含">
                      <Select>
                        <Option value="yes">是 (yes)</Option>
                        <Option value="no">否 (no)</Option>
                      </Select>
                    </Form.Item>
                  </Col>

                  <Col span={6}>
                    <Form.Item name="courseFee__summerSurcharge" label="暑期限加價 (美金)">
                      <InputNumber style={{ width: "100%" }} prefix="US$" />
                    </Form.Item>
                  </Col>
                  <Col span={6}>
                    <Form.Item name="courseFee__legalGuardianFee" label="未成年監護費 (美金)">
                      <InputNumber style={{ width: "100%" }} prefix="US$" />
                    </Form.Item>
                  </Col>

                  {/* 👇 新增：暑期加價時間區間 */}
                  <Col span={6}>
                    <Form.Item name="courseFee__summerSurchargeStartTime" label="暑期加價開始日期">
                      <Input placeholder="例如: 7/5" />
                    </Form.Item>
                  </Col>
                  <Col span={6}>
                    <Form.Item name="courseFee__summerSurchargeEndTime" label="暑期加價結束日期">
                      <Input placeholder="例如: 8/29" />
                    </Form.Item>
                  </Col>

                  <Col span={6}>
                    <Form.Item name="courseFee__discount" label="一般折扣率 (0-1)">
                      <InputNumber
                        style={{ width: "100%" }}
                        min={0} max={1} step={0.01}
                        placeholder="0.05 代表 5%"
                      />
                    </Form.Item>
                  </Col>

                  {/* 👇 新增：折扣是否內含與固定折扣金額 */}
                  <Col span={6}>
                    <Form.Item name="courseFee__discountInclusive" label="折扣是否內含(代碼)">
                      <Input placeholder="預設填 0" />
                    </Form.Item>
                  </Col>
                  <Col span={6}>
                    <Form.Item name="courseFee__fixedDiscount" label="固定折扣金額">
                      <InputNumber style={{ width: "100%" }} prefix="US$" />
                    </Form.Item>
                  </Col>
                </Row>
              ),
            },
            {
              key: "logo",
              label: "Logo 印章位置設定 (Excel)",
              children: (
                <Row gutter={16}>
                  <Col xs={12} md={6}>
                    <Form.Item name="logo__picLocationRow" label="起始列">
                      <InputNumber style={{ width: "100%" }} />
                    </Form.Item>
                  </Col>
                  <Col xs={12} md={6}>
                    <Form.Item name="logo__picLocationCol" label="起始欄">
                      <InputNumber style={{ width: "100%" }} />
                    </Form.Item>
                  </Col>
                  <Col xs={12} md={6}>
                    <Form.Item name="logo__picLocationLeft" label="Left 偏移(px)">
                      <InputNumber style={{ width: "100%" }} />
                    </Form.Item>
                  </Col>
                  <Col xs={12} md={6}>
                    <Form.Item name="logo__picLocationTop" label="Top 偏移(px)">
                      <InputNumber style={{ width: "100%" }} />
                    </Form.Item>
                  </Col>
                  <Col span={12}>
                    <Form.Item name="logo__picwidth" label="圖片寬度 (px)">
                      <InputNumber style={{ width: "100%" }} />
                    </Form.Item>
                  </Col>
                  <Col span={12}>
                    <Form.Item name="logo__picheight" label="圖片高度 (px)">
                      <InputNumber style={{ width: "100%" }} />
                    </Form.Item>
                  </Col>
                </Row>
              ),
            },
            {
              key: "more4week",
              label: "長週／短週折扣設定",
              children: (
                <Row gutter={16}>
                  <Col span={6}>
                    <Form.Item name="more4week1" label=">4 週第 1 段折 100%">
                      <InputNumber style={{ width: "100%" }} step={0.01} />
                    </Form.Item>
                  </Col>
                  <Col span={6}>
                    <Form.Item name="more4week2" label="第 2 段折">
                      <InputNumber style={{ width: "100%" }} step={0.01} />
                    </Form.Item>
                  </Col>
                  <Col span={6}>
                    <Form.Item name="more4week3" label="第 3 段折">
                      <InputNumber style={{ width: "100%" }} step={0.01} />
                    </Form.Item>
                  </Col>
                  <Col span={6}>
                    <Form.Item name="more4week4" label="第 4 段折">
                      <InputNumber style={{ width: "100%" }} step={0.01} />
                    </Form.Item>
                  </Col>
                </Row>
              ),
            },
            {
              key: "less4week",
              label: "少於 4 週折扣設定",
              children: (
                <Row gutter={16}>
                  <Col span={8}>
                    <Form.Item name="less4week1" label="<4 週第 1 段折">
                      <InputNumber 
                        style={{ width: "100%" }} 
                        min={0} 
                        precision={2}
                        placeholder="0 = 不折"
                        step={0.01} />
                    </Form.Item>
                  </Col>
                  <Col span={8}>
                    <Form.Item name="less4week2" label="第 2 段折">
                      <InputNumber style={{ width: "100%" }} step={0.01} />
                    </Form.Item>
                  </Col>
                  <Col span={8}>
                    <Form.Item name="less4week3" label="第 3 段折">
                      <InputNumber style={{ width: "100%" }} step={0.01} />
                    </Form.Item>
                  </Col>
                </Row>
              ),
            },
            {
              key: "localFee",
              label: "當地雜費設定 (localFee)",
              children: (
                <>
                  <Table
                    dataSource={localFeeDataSource}
                    columns={localFeeColumns}
                    pagination={false}
                    size="small"
                    bordered
                    style={{ marginBottom: 16 }}
                  />
                  <Button
                    type="dashed"
                    icon={<PlusOutlined />}
                    onClick={handleAddLocalFee}
                  >
                    新增一筆雜費
                  </Button>
                </>
              ),
            },
            {
              key: "courses",
              label: "課程清單設定 (courses)",
              children: (
                <>
                  <Row gutter={16} style={{ marginBottom: 12 }}>
                    <Col span={12}>
                      <label>選擇課程群組</label>
                      <Select
                        value={courseGroup}
                        onChange={setCourseGroup}
                        style={{ width: "100%" }}
                      >
                        {Object.keys(data.courses).map((group) => (
                          <Option key={group} value={group}>
                            {group}
                          </Option>
                        ))}
                      </Select>
                    </Col>
                    <Col span={12}>
                      <Button
                        type="primary"
                        icon={<PlusOutlined />}
                        onClick={handleAddCourse}
                        style={{ marginTop: 26 }}
                      >
                        新增課程
                      </Button>
                    </Col>
                  </Row>
                  <Table
                    dataSource={currentCourses.map((item) => ({ ...item, key: item.code }))}
                    columns={courseColumns}
                    pagination={false}
                    size="small"
                    bordered
                  />
                </>
              ),
            },
            {
              key: "room",
              label: "房型設定 (room)",
              children: (
                <Table
                  dataSource={data.room.map((item) => ({ ...item, key: item.code }))}
                  columns={roomColumns}
                  pagination={false}
                  size="small"
                  bordered
                />
              ),
            },
          ]}
        />
          {/* <Divider />
          <Button
            type="primary"
            onClick={handleSaveAll}
            loading={loading} // 你可以自行加一個 loading state
          >
            儲存所有設定
          </Button>

          <Button
            type="primary"
            icon={<FileExcelOutlined />}
            onClick={handleExportJSON}
            size="large"
            style={{ marginLeft: 12 }}
          >
            匯出設定檔 (JSON)
          </Button> */}

        </Form> 
        {/* 在最下面加一個「測試區」 */}
        <div style={{ marginTop: 32, borderTop: "1px solid #d9d9d9", paddingTop: 16 }}>
          <h2>📂 檔案讀取測試</h2>

          {/* 👉 把 loadedSettingData 傳給 FileUploadTest */}
          <FileUploadTest
          loadedSettingData={loadedSettingData}
          onDataLoaded={handleDataLoaded}
          />
        </div>
    </div>
  );
}

// 👇 把上面寫的 FileUploadTest 也貼在這裡
// 2. 在 FileUploadTest 裡用 props 接，不要直接用 loadedSettingData
function FileUploadTest({ loadedSettingData, onDataLoaded }) {
  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    console.log("📁 檔案名稱:", file.name);
    console.log("🔍 檔案類型:", file.type);
    console.log("📏 檔案大小:", file.size, "bytes");

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target.result;
      console.log("📄 檔案內容：", content);

      try {
        const parsed = JSON.parse(content);
        console.log("✅ 解析後的設定資料:", parsed);

        if (onDataLoaded) {
          onDataLoaded(parsed);
        }

      } catch (err) {
        console.error("❌ JSON 格式錯誤:", err.message);
      }
    };

    reader.readAsText(file);
  };

  return (
    <div>
      <p>請選擇一個設定檔（JSON）測試，會在 Console 顯示內容</p>
      <div style={{ margin: "12px 0" }}>
        <input
          type="file"
          onChange={handleFileChange}
          style={{ padding: "8px" }}
        />
      </div>

      {/* 👉 用 props 傳進來，而不是直接用外部變數 */}
      {loadedSettingData && (
        <div style={{ marginTop: 16, border: "1px solid #d9d9d9", padding: "12px" }}>
          <h3>📊 檔案會套用到以下設定：</h3>
          <p>學校：{loadedSettingData.setting?.schoolLocation}</p>
          <p>機票目的地：{loadedSettingData.setting?.flightDest}</p>
          <p>註冊費：{loadedSettingData.courseFee?.registrationfee} 美金</p>
        </div>
      )}
    </div>
  );
}


// 如果你之後要把這個元件 export 給別人用，也可以多一行
// export default FileUploadTest;  // 但你現在只在 SettingPage 裡用，所以不用也可以
export default SettingPage;








