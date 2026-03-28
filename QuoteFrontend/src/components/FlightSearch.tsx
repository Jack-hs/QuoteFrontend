import { useState, useCallback } from "react";
import axios from "axios";
import { 
  Button, Input, Select, Badge, Divider, 
  Tag, Spin, Alert, Row, Col, Typography 
} from "antd";
import { 
  SearchOutlined, CheckCircleOutlined, 
  LoadingOutlined, AlertOutlined,
  ArrowRightOutlined
} from "@ant-design/icons";

const { Text, Title } = Typography;
const { Option } = Select;

// ===== 類型定義 =====
interface FlightResult {
  price: number;
  currency: string;
  airline: string;
  airlineLogo: string;
  totalDuration: number;
  stops: number;
  departureTime: string;
  arrivalTime: string;
  flightNumbers: string[];
  bookingToken: string;
  type: string;
}

// ===== 工具函數 =====
function formatDuration(minutes: number): string {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${h}h ${m}m`;
}

function formatTime(dateTimeStr: string): string {
  if (!dateTimeStr) return "—";
  const parts = dateTimeStr.split(" ");
  return parts[1] ?? dateTimeStr;
}

// ===== 機場清單 =====
const airports = [
  { code: 'TPE', city: '台北桃園' },
  { code: 'TSA', city: '台北松山' },
  { code: 'KHH', city: '高雄' },
  { code: 'CEB', city: '宿霧' },
  { code: 'MNL', city: '馬尼拉' },
  { code: 'ILO', city: '伊洛伊洛' },
  { code: 'BCD', city: '巴科洛德' },
];

// ===== 單一航班卡片 =====
interface FlightCardProps {
  flight: FlightResult;
  isBest?: boolean;
  onSelect: (flight: FlightResult) => void;
}

function FlightCard({ flight, isBest, onSelect }: FlightCardProps) {
  return (
    <div style={{
      display: 'flex',
      alignItems: 'center',
      gap: 12,
      padding: '12px 16px',
      borderRadius: 8,
      border: isBest ? '1px solid #91caff' : '1px solid #f0f0f0',
      backgroundColor: isBest ? '#e6f7ff' : '#fafafa',
      marginBottom: 8,
      cursor: 'pointer',
      transition: 'all 0.2s'
    }}>

      {/* 航空公司 Logo */}
      <div style={{
        width: 44, height: 44,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        background: '#fff', borderRadius: 8, border: '1px solid #f0f0f0',
        overflow: 'hidden', flexShrink: 0
      }}>
        {flight.airlineLogo ? (
          <img 
            src={flight.airlineLogo} 
            alt={flight.airline} 
            style={{ width: 36, height: 36, objectFit: 'contain' }} 
          />
        ) : (
          <span style={{ fontSize: 20 }}>✈️</span>
        )}
      </div>

      {/* 航班資訊 */}
      <div style={{ flex: 1 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
          <Text strong>{flight.airline}</Text>
          {isBest && <Tag color="blue">⭐ 最佳選擇</Tag>}
          {flight.stops === 0 
            ? <Tag color="green">直飛</Tag> 
            : <Tag color="orange">{flight.stops} 次轉機</Tag>
          }
        </div>

        <div style={{ 
          display: 'flex', alignItems: 'center', gap: 8, 
          marginTop: 4, color: '#888', fontSize: 13 
        }}>
          <Text type="secondary">{formatTime(flight.departureTime)}</Text>
          <ArrowRightOutlined style={{ fontSize: 11 }} />
          <Text type="secondary">{formatTime(flight.arrivalTime)}</Text>
          <Text type="secondary">🕐 {formatDuration(flight.totalDuration)}</Text>
          {flight.flightNumbers.length > 0 && (
            <Text type="secondary">{flight.flightNumbers.join(", ")}</Text>
          )}
        </div>
      </div>

      {/* 價格 + 按鈕 */}
      <div style={{ textAlign: 'right', flexShrink: 0 }}>
        <div style={{ fontSize: 18, fontWeight: 'bold', color: '#cf1322' }}>
          NT${flight.price.toLocaleString()}
        </div>
        <div style={{ fontSize: 12, color: '#888', marginBottom: 6 }}>
          {flight.type === "Round trip" || flight.type === "來回" ? "來回" : "單程"}
        </div>
        <Button
          size="small"
          type="primary"
          icon={<CheckCircleOutlined />}
          onClick={() => onSelect(flight)}
        >
          帶入報價
        </Button>
      </div>
    </div>
  );
}

// ===== 主元件 =====
interface FlightSearchProps {
  defaultDeparture?: string;
  defaultArrival?: string;
  defaultOutboundDate?: string;
  defaultReturnDate?: string;
  defaultAdults?: number;
  onSelectFlight?: (price: number, description: string) => void;
}

export default function FlightSearch({
  defaultDeparture = "TPE",
  defaultArrival = "CEB",
  defaultOutboundDate = "",
  defaultReturnDate = "",
  defaultAdults = 1,
  onSelectFlight,
}: FlightSearchProps) {
  const [departure, setDeparture] = useState(defaultDeparture);
  const [arrival, setArrival] = useState(defaultArrival);
  const [outboundDate, setOutboundDate] = useState(defaultOutboundDate);
  const [returnDate, setReturnDate] = useState(defaultReturnDate);
  const [adults, setAdults] = useState(defaultAdults);
  const [tripType, setTripType] = useState<"1" | "2">("1");
  const [showOther, setShowOther] = useState(false);

  const [results, setResults] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // ===== 查詢 =====
  const handleSearch = useCallback(async () => {
    if (!departure) { alert("請選擇出發機場"); return; }
    if (!arrival)   { alert("請選擇目的地機場"); return; }
    if (!outboundDate) { alert("請選擇出發日期"); return; }
    if (tripType === "1" && !returnDate) { alert("來回票請選擇回程日期"); return; }

    setIsLoading(true);
    setError(null);
    setResults(null);

    try {
      //const apiUrl = import.meta.env.VITE_API_URL || '';
      const apiUrl = 'https://localhost:7080';
      const res = await axios.get(apiUrl+'/api/flight/search', {
        params: {
          from: departure,
          to: arrival,
          outboundDate,
          returnDate: tripType === "1" ? returnDate : undefined,
          adults,
          tripType
        }
      });
      setResults(res.data);
    } catch (err: any) {
      setError(err.response?.data?.message || err.message || "查詢失敗");
    } finally {
      setIsLoading(false);
    }
  //   // 👇 暫時用假資料（後端 API 還沒好）
  // await new Promise(r => setTimeout(r, 800)); // 模擬等待
  // setResults({
  //   bestFlights: [
  //     {
  //       price: 12800,
  //       currency: 'TWD',
  //       airline: '中華航空 CI',
  //       airlineLogo: '',
  //       totalDuration: 185,
  //       stops: 0,
  //       departureTime: `${outboundDate} 08:00`,
  //       arrivalTime: `${outboundDate} 11:05`,
  //       flightNumbers: ['CI721'],
  //       bookingToken: '',
  //       type: '來回'
  //     },
  //     {
  //       price: 13500,
  //       currency: 'TWD',
  //       airline: '長榮航空 BR',
  //       airlineLogo: '',
  //       totalDuration: 195,
  //       stops: 0,
  //       departureTime: `${outboundDate} 14:30`,
  //       arrivalTime: `${outboundDate} 17:45`,
  //       flightNumbers: ['BR853'],
  //       bookingToken: '',
  //       type: '來回'
  //     }
  //   ],
  //   otherFlights: [
  //     {
  //       price: 15200,
  //       currency: 'TWD',
  //       airline: '菲律賓航空 PR',
  //       airlineLogo: '',
  //       totalDuration: 240,
  //       stops: 1,
  //       departureTime: `${outboundDate} 06:00`,
  //       arrivalTime: `${outboundDate} 10:00`,
  //       flightNumbers: ['PR890'],
  //       bookingToken: '',
  //       type: '來回'
  //     }
  //   ],
  //   priceInsights: {
  //     lowestPrice: 12800,
  //     typicalRangeMin: 11000,
  //     typicalRangeMax: 18000
  //   }
  // });
  // setIsLoading(false);
  }, [departure, arrival, outboundDate, returnDate, adults, tripType]);

  // ===== 選擇航班 =====
  const handleSelectFlight = useCallback((flight: FlightResult) => {
    if (!onSelectFlight) return;
    const tripLabel = tripType === "1" ? "來回" : "單程";
    const description = `${flight.airline} ${flight.flightNumbers.join("/")} ${departure}↔${arrival} ${tripLabel}機票`;
    onSelectFlight(flight.price, description);
  }, [departure, arrival, tripType, onSelectFlight]);

  const bestFlights = results?.bestFlights ?? [];
  const otherFlights = results?.otherFlights ?? [];
  const priceInsights = results?.priceInsights;
  const hasResults = bestFlights.length > 0 || otherFlights.length > 0;

  return (
    <div style={{ padding: '8px 0' }}>

      {/* 查詢表單 */}
      <Row gutter={[12, 12]} style={{ marginBottom: 16 }}>
        <Col span={6}>
          <label style={{ fontSize: 12, fontWeight: 'bold' }}>出發機場</label>
          <Select 
            value={departure} 
            onChange={setDeparture}
            style={{ width: '100%', marginTop: 4 }}
          >
            {airports.filter(a => ['TPE','TSA','KHH'].includes(a.code)).map(a => (
              <Option key={a.code} value={a.code}>{a.code} - {a.city}</Option>
            ))}
          </Select>
        </Col>

        <Col span={6}>
          <label style={{ fontSize: 12, fontWeight: 'bold' }}>目的地機場</label>
          <Select 
            value={arrival} 
            onChange={setArrival}
            style={{ width: '100%', marginTop: 4 }}
          >
            {airports.map(a => (
              <Option key={a.code} value={a.code}>{a.code} - {a.city}</Option>
            ))}
          </Select>
        </Col>

        <Col span={5}>
          <label style={{ fontSize: 12, fontWeight: 'bold' }}>去程日期</label>
          <Input
            type="date"
            value={outboundDate}
            onChange={e => setOutboundDate(e.target.value)}
            style={{ marginTop: 4 }}
          />
        </Col>

        <Col span={5}>
          <label style={{ fontSize: 12, fontWeight: 'bold' }}>回程日期</label>
          <Input
            type="date"
            value={returnDate}
            onChange={e => setReturnDate(e.target.value)}
            disabled={tripType === "2"}
            style={{ marginTop: 4 }}
          />
        </Col>

        <Col span={2}>
          <label style={{ fontSize: 12, fontWeight: 'bold' }}>類型</label>
          <Select 
            value={tripType} 
            onChange={(v) => setTripType(v)}
            style={{ width: '100%', marginTop: 4 }}
          >
            <Option value="1">來回</Option>
            <Option value="2">單程</Option>
          </Select>
        </Col>

        <Col span={24} style={{ textAlign: 'right' }}>
          <Button
            type="primary"
            icon={isLoading ? <LoadingOutlined /> : <SearchOutlined />}
            onClick={handleSearch}
            loading={isLoading}
            size="large"
          >
            {isLoading ? "查詢中..." : "🔍 查詢機票"}
          </Button>
        </Col>
      </Row>

      {/* 錯誤訊息 */}
      {error && (
        <Alert
          type="error"
          message="查詢失敗"
          description={error}
          showIcon
          style={{ marginBottom: 16 }}
        />
      )}

      {/* 載入中 */}
      {isLoading && (
        <div style={{ textAlign: 'center', padding: '32px 0' }}>
          <Spin indicator={<LoadingOutlined style={{ fontSize: 24 }} spin />} />
          <p style={{ marginTop: 12, color: '#888' }}>
            正在查詢 Google Flights 最新機票資訊...
          </p>
        </div>
      )}

      {/* 查詢結果 */}
      {!isLoading && hasResults && (
        <div>
          {/* 價格洞察 */}
          {priceInsights?.lowestPrice > 0 && (
            <Alert
              type="success"
              message={
                <span>
                  💰 最低價格 <strong>NT${priceInsights.lowestPrice.toLocaleString()}</strong>，
                  一般區間 NT${priceInsights.typicalRangeMin.toLocaleString()} – NT${priceInsights.typicalRangeMax.toLocaleString()}
                </span>
              }
              style={{ marginBottom: 16 }}
            />
          )}

          {/* 最佳航班 */}
          {bestFlights.length > 0 && (
            <>
              <Text strong style={{ color: '#888', fontSize: 12 }}>⭐ 最佳航班</Text>
              <div style={{ marginTop: 8 }}>
                {bestFlights.map((flight: FlightResult, idx: number) => (
                  <FlightCard
                    key={`best-${idx}`}
                    flight={flight}
                    isBest={idx === 0}
                    onSelect={handleSelectFlight}
                  />
                ))}
              </div>
            </>
          )}

          {/* 其他航班 */}
          {otherFlights.length > 0 && (
            <>
              <Divider />
              <Button 
                type="link" 
                onClick={() => setShowOther(p => !p)}
                style={{ padding: 0, marginBottom: 8 }}
              >
                {showOther ? '▲' : '▼'} 其他航班（{otherFlights.length} 個選項）
              </Button>
              {showOther && otherFlights.map((flight: FlightResult, idx: number) => (
                <FlightCard
                  key={`other-${idx}`}
                  flight={flight}
                  onSelect={handleSelectFlight}
                />
              ))}
            </>
          )}

          <Text type="secondary" style={{ fontSize: 11, display: 'block', textAlign: 'right', marginTop: 12 }}>
            資料來源：Google Flights（via SerpAPI）· 價格僅供參考
          </Text>
        </div>
      )}

      {/* 無結果 */}
      {!isLoading && results && !hasResults && !error && (
        <div style={{ textAlign: 'center', padding: '32px 0', color: '#888' }}>
          <span style={{ fontSize: 32 }}>✈️</span>
          <p style={{ marginTop: 8 }}>查無符合條件的航班，請調整日期或機場</p>
        </div>
      )}
    </div>
  );
}
