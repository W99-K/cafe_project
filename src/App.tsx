/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useEffect, useState } from 'react';
import { supabase } from './lib/supabase';

export default function App() {
  const [status, setStatus] = useState<'loading' | 'connected' | 'error'>('loading');
  const [errorMsg, setErrorMsg] = useState('');

  useEffect(() => {
    async function testConnection() {
      try {
        // 간단한 쿼리로 연결 테스트
        const { error } = await supabase.from('orders').select('count').limit(1);
        if (error && error.code !== 'PGRST116' && error.code !== '42P01') {
          // PGRST116: 테이블 없음, 42P01: relation does not exist → 연결은 성공
          setErrorMsg(error.message);
          setStatus('error');
        } else {
          setStatus('connected');
        }
      } catch (e: unknown) {
        setErrorMsg(e instanceof Error ? e.message : '알 수 없는 오류');
        setStatus('error');
      }
    }
    testConnection();
  }, []);

  return (
    <div style={{ fontFamily: 'sans-serif', padding: '2rem' }}>
      <h1>☕ 바이브 카페 주문서</h1>
      <p>
        Supabase 연결 상태:{' '}
        {status === 'loading' && '⏳ 연결 중...'}
        {status === 'connected' && '✅ 연결 성공!'}
        {status === 'error' && `❌ 연결 실패: ${errorMsg}`}
      </p>
    </div>
  );
}
