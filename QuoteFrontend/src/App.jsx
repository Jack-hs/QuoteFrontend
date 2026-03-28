// App.jsx

import React from 'react';
import { Layout, Menu } from 'antd';      // 👈 不要從 antd import Content
const { Sider } = Layout;  // 或直接 destructuring
import { Routes, Route, useNavigate, useLocation } from 'react-router-dom';
import { FileTextOutlined, SettingOutlined, UploadOutlined } from '@ant-design/icons';
import MainQuotePage from './pages/MainQuotePage';
import SettingPage from './pages/SettingPage';
// 引入我們剛寫好的上傳元件 (假設您放在 pages 資料夾下，檔名為 ExcelUploader.jsx)
import ExcelUploader from './pages/ExcelUploader'; 

const { Header } = Layout;

function AppLayout() {
  const navigate = useNavigate();
  const location = useLocation();

  const onMenuClick = (key) => {
    navigate(key);
  };

  const selectedKeys = [location.pathname];

  return (
    <Layout style={{ minHeight: '100vh' }}>
      {/* 左側選單 Sider */}
      <Sider collapsible theme="light">
        <div style={{ height: 32, margin: 16, background: '#fff' }} />
        <Menu
          theme="light"
          mode="inline"
          selectedKeys={selectedKeys}
          items={[
            {
              key: '/quote',
              icon: <FileTextOutlined />,
              label: '報價系統',
              onClick: () => onMenuClick('/quote'),
            },
            {
              key: '/setting',
              icon: <SettingOutlined />,
              label: '系統設定',
              onClick: () => onMenuClick('/setting'),
            },
            // ======== 新增的選單項目 ========
            {
              key: '/upload',
              icon: <UploadOutlined />,
              label: '更新 Excel 資料',
              onClick: () => onMenuClick('/upload'),
            },
            // ================================
          ]}
        />
      </Sider>

      {/* 主要區域，用 Layout.Content 來包 Routes */}
      <Layout>
        <Header style={{ background: '#fff', padding: 0 }} />
        <Layout.Content style={{ margin: 16 }}>
          <Routes>
            <Route path="/quote" element={<MainQuotePage />} />
            <Route path="/setting" element={<SettingPage />} />
            {/* ======== 新增的 Route ======== */}
            <Route path="/upload" element={<ExcelUploader />} />
            {/* ============================== */}
            <Route path="*" element={<MainQuotePage />} />
          </Routes>
        </Layout.Content>
      </Layout>
    </Layout>
  );
}

export default AppLayout;
