import Head from 'next/head';
import { useState } from 'react';

export default function Home() {
  const [date, setDate] = useState('');
  const [loading, setLoading] = useState(false);

  const handleScrape = async () => {
    if (!date) return alert('날짜를 선택하세요.');
    setLoading(true);
    try {
      const res = await fetch(`/api/scrape?date=${date}`);
      if (!res.ok) {
        const text = await res.text();
        throw new Error(`스크랩 실패: ${res.status} ${text}`);
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `GS칼텍스뉴스_${date}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch (err) {
      console.error(err);
      const msg = err instanceof Error ? err.message : String(err);
      alert(`오류 발생: ${msg}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-gray-100">
      <Head>
        <title>GS칼텍스 뉴스 스크래퍼</title>
      </Head>
      <h1 className="text-2xl font-bold mb-4">GS칼텍스 주요 뉴스 스크랩</h1>
      <input
        type="date"
        value={date}
        onChange={e => setDate(e.target.value)}
        className="border p-2 mb-4"
      />
      <button
        onClick={handleScrape}
        disabled={loading}
        className="bg-blue-500 text-white px-4 py-2 rounded"
      >
        {loading ? '스크랩 중...' : 'GS칼텍스 주요 뉴스 스크랩하기'}
      </button>
    </div>
  );
}
