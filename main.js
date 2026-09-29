/**
 * ==========================================================================
 * 바이브 카페 (Vibe Cafe) 주문서 & 실시간 주문 게시판 메인 스크립트 (main.js)
 * 
 * [주요 기능 목록]
 * 1. 탭 네비게이션: '주문서 작성' 탭과 '실시간 주문 게시판' 탭 간의 매끄러운 전환
 * 2. 실시간 주문 계산: 음료, 사이즈, 추가 옵션, 수량 변경 시 예상 금액 즉각 반영
 * 3. 주문 접수 처리: 유효성 검사, 주문 확인 메시지 표시, 색종이 폭죽 애니메이션
 * 4. 실시간 주문 게시판:
 *    - 주문 상태 흐름: 접수 완료(received) -> 제조 중(making) -> 제조 완료(completed)
 *    - 실시간 통계 요약 (전체 주문 수, 상태별 주문 수, 실시간 누적 매출)
 *    - 상태별 필터링 (전체, 접수, 제조중, 완료) 및 실시간 검색(고객명, 연락처, 음료명)
 *    - 상대적 시간 표기 ('방금 전', '1분 전' 등) 및 주기적 갱신
 *    - Supabase Realtime 구독으로 여러 기기 간 실시간 동기화
 * 5. 엑셀(Excel .csv) 내역 다운로드 (UTF-8 BOM 지원)
 * ==========================================================================
 */

import { createClient } from '@supabase/supabase-js';

// Supabase 클라이언트 초기화 (Vite 환경변수 사용)
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

document.addEventListener('DOMContentLoaded', () => {
  // ------------------------------------------------------------------------
  // 1. 주요 DOM 엘리먼트 참조
  // ------------------------------------------------------------------------

  // 탭 네비게이션 엘리먼트
  const tabBtnOrder = document.getElementById('tab-btn-order');
  const tabBtnBoard = document.getElementById('tab-btn-board');
  const panelOrder = document.getElementById('panel-order');
  const panelBoard = document.getElementById('panel-board');
  const boardLiveBadge = document.getElementById('board-live-badge');
  const btnGotoBoard = document.getElementById('btn-goto-board');

  // 주문 폼 관련 엘리먼트
  const orderForm = document.getElementById('order-form');
  const customerNameInput = document.getElementById('customer-name');
  const customerPhoneInput = document.getElementById('customer-phone');
  const diningTypeRadios = document.querySelectorAll('input[name="diningType"]');
  const beverageSelect = document.getElementById('beverage-select');
  const sizeRadios = document.querySelectorAll('input[name="size"]');
  const optionCheckboxes = document.querySelectorAll('input[name="option"]');
  const quantityInput = document.getElementById('quantity');
  const requestsTextarea = document.getElementById('requests');

  // 주문서 화면 버튼 및 안내 영역 엘리먼트
  const estimatedPriceDisplay = document.getElementById('estimated-price');
  const confirmationBox = document.getElementById('order-confirmation');
  const confirmationText = document.getElementById('confirmation-text');
  const alertBox = document.getElementById('alert-box');
  const btnReset = document.getElementById('btn-reset');
  const btnExcel = document.getElementById('btn-download-excel');
  const orderCountBadge = document.getElementById('order-count-badge');
  const recentOrdersList = document.getElementById('recent-orders-list');
  const confettiCanvas = document.getElementById('confetti-canvas');

  // 주문 게시판 관련 엘리먼트
  const boardOrdersContainer = document.getElementById('board-orders-container');
  const boardTotalRevenue = document.getElementById('board-total-revenue');
  const statValTotal = document.getElementById('stat-val-total');
  const statValReceived = document.getElementById('stat-val-received');
  const statValMaking = document.getElementById('stat-val-making');
  const statValCompleted = document.getElementById('stat-val-completed');
  const filterBtns = document.querySelectorAll('.filter-btn');
  const filterCountAll = document.getElementById('filter-count-all');
  const filterCountReceived = document.getElementById('filter-count-received');
  const filterCountMaking = document.getElementById('filter-count-making');
  const filterCountCompleted = document.getElementById('filter-count-completed');
  const boardSearchInput = document.getElementById('board-search-input');
  const btnClearSearch = document.getElementById('btn-clear-search');
  const btnBoardExcel = document.getElementById('btn-board-excel');
  const btnAddSampleOrder = document.getElementById('btn-add-sample-order');
  const btnClearAllOrders = document.getElementById('btn-clear-all-orders');

  // ------------------------------------------------------------------------
  // 2. 주문 데이터 관리 (Supabase + 로컬 스토리지 병행)
  // ------------------------------------------------------------------------
  const STORAGE_KEY = 'vibe_cafe_orders';
  let orderHistory = [];
  let currentFilterStatus = 'all'; // 'all' | 'received' | 'making' | 'completed'
  let currentSearchQuery = '';

  // DB 행 → 앱 객체 변환 (cafe_menu 컬럼 → camelCase)
  function rowToOrder(row) {
    return {
      id: row.id,                                               // UUID
      createdAtMs: row.created_at ? new Date(row.created_at).getTime() : Date.now(),
      timestamp: row.created_at
        ? new Date(row.created_at).toLocaleString('ko-KR')
        : '',
      customerName: row.customer_name || '',
      customerPhone: row.customer_phone || '',
      diningTypeName: row.dining_type || '포장',
      diningTypeValue: row.dining_type === '매장' ? 'store' : 'takeout',
      beverageName: row.beverage_name || '',
      beveragePrice: row.beverage_price || 0,
      sizeName: row.size || 'M',
      sizePrice: row.size_extra_price || 0,
      // options 컬럼은 ARRAY → 콤마 구분 문자열로 변환
      optionsText: Array.isArray(row.options) ? row.options.join(', ') : (row.options || ''),
      optionsPrice: row.options_extra_price || 0,
      quantity: row.quantity || 1,
      totalPrice: row.total_price || 0,
      requests: row.requests || '',
      status: row.order_status || 'received',
    };
  }

  // 앱 객체 → DB 행 변환 (camelCase → cafe_menu 컬럼)
  function orderToRow(order) {
    return {
      id: order.id,                                             // UUID (crypto.randomUUID()로 생성)
      customer_name: order.customerName,
      customer_phone: order.customerPhone || '',
      dining_type: order.diningTypeName,
      beverage_name: order.beverageName,
      beverage_price: order.beveragePrice,
      size: order.sizeName,
      size_extra_price: order.sizePrice,
      // 문자열 → ARRAY 변환 (빈 문자열이면 빈 배열)
      options: order.optionsText
        ? order.optionsText.split(', ').filter(Boolean)
        : [],
      options_extra_price: order.optionsPrice,
      quantity: order.quantity,
      total_price: order.totalPrice,
      requests: order.requests || '',
      order_status: order.status || 'received',
    };
  }

  // 로컬 스토리지 캐시 저장
  function saveToLocalCache() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(orderHistory));
    } catch (e) {
      console.error('로컬 캐시 저장 실패:', e);
    }
  }

  // 로컬 스토리지 캐시에서 불러오기 (오프라인 폴백)
  function loadFromLocalCache() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        orderHistory = JSON.parse(saved).map((order) => {
          if (!order.status) order.status = 'received';
          if (!order.createdAtMs) order.createdAtMs = order.id;
          return order;
        });
      } else {
        orderHistory = [];
      }
    } catch (e) {
      console.error('로컬 캐시 불러오기 오류:', e);
      orderHistory = [];
    }
  }

  // Supabase에서 전체 주문 불러오기 (초기 로드)
  async function loadOrdersFromSupabase() {
    try {
      const { data, error } = await supabase
        .from('cafe_menu')
        .select('*')
        .order('created_at', { ascending: true });

      if (error) {
        console.error('Supabase 불러오기 오류:', error.message);
        loadFromLocalCache(); // 오류 시 로컬 캐시 사용
      } else {
        orderHistory = (data || []).map(rowToOrder);
        saveToLocalCache();
      }
    } catch (e) {
      console.error('Supabase 연결 실패, 로컬 캐시 사용:', e);
      loadFromLocalCache();
    }
    updateAllUI();
  }

  // Supabase에 신규 주문 저장
  async function saveOrderToSupabase(order) {
    try {
      const { error } = await supabase.from('cafe_menu').insert([orderToRow(order)]);
      if (error) console.error('Supabase 주문 저장 실패:', error.message);
    } catch (e) {
      console.error('Supabase 저장 오류:', e);
    }
    saveToLocalCache();
  }

  // Supabase에서 주문 상태 업데이트
  async function updateOrderStatusInSupabase(orderId, newStatus) {
    try {
      const { error } = await supabase
        .from('cafe_menu')
        .update({ order_status: newStatus })
        .eq('id', orderId);
      if (error) console.error('Supabase 상태 업데이트 실패:', error.message);
    } catch (e) {
      console.error('Supabase 상태 업데이트 오류:', e);
    }
    saveToLocalCache();
  }

  // Supabase에서 주문 삭제
  async function deleteOrderFromSupabase(orderId) {
    try {
      const { error } = await supabase.from('cafe_menu').delete().eq('id', orderId);
      if (error) console.error('Supabase 삭제 실패:', error.message);
    } catch (e) {
      console.error('Supabase 삭제 오류:', e);
    }
    saveToLocalCache();
  }

  // Supabase에서 전체 주문 삭제
  async function deleteAllOrdersFromSupabase() {
    try {
      // UUID 기반이므로 created_at 조건으로 전체 삭제
      const { error } = await supabase
        .from('cafe_menu')
        .delete()
        .gte('created_at', '1970-01-01');
      if (error) console.error('Supabase 전체 삭제 실패:', error.message);
    } catch (e) {
      console.error('Supabase 전체 삭제 오류:', e);
    }
    saveToLocalCache();
  }

  // 초기 주문 데이터 로드 (로컬 캐시 먼저 표시 후 Supabase 동기화)
  loadFromLocalCache();
  updateAllUI(); // 로컬 캐시로 즉시 렌더링
  loadOrdersFromSupabase(); // 비동기로 Supabase 동기화

  // Supabase Realtime 구독 (다른 기기/탭의 변경 사항 실시간 반영)
  supabase
    .channel('orders-realtime')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'cafe_menu' }, () => {
      // 변경 발생 시 Supabase에서 전체 데이터 재로드
      loadOrdersFromSupabase();
    })
    .subscribe();

  // ------------------------------------------------------------------------
  // 3. 탭 전환 처리 (주문서 작성 vs 실시간 주문 게시판)
  // ------------------------------------------------------------------------
  function switchTab(targetTab) {
    if (targetTab === 'order') {
      // 1) 주문서 작성 탭 활성화
      tabBtnOrder.classList.add('active');
      tabBtnOrder.setAttribute('aria-selected', 'true');
      tabBtnOrder.setAttribute('tabindex', '0');

      tabBtnBoard.classList.remove('active');
      tabBtnBoard.setAttribute('aria-selected', 'false');
      tabBtnBoard.setAttribute('tabindex', '-1');

      panelOrder.classList.add('active');
      panelOrder.removeAttribute('hidden');

      panelBoard.classList.remove('active');
      panelBoard.setAttribute('hidden', '');
    } else if (targetTab === 'board') {
      // 2) 실시간 주문 게시판 탭 활성화
      tabBtnBoard.classList.add('active');
      tabBtnBoard.setAttribute('aria-selected', 'true');
      tabBtnBoard.setAttribute('tabindex', '0');

      tabBtnOrder.classList.remove('active');
      tabBtnOrder.setAttribute('aria-selected', 'false');
      tabBtnOrder.setAttribute('tabindex', '-1');

      panelBoard.classList.add('active');
      panelBoard.removeAttribute('hidden');

      panelOrder.classList.remove('active');
      panelOrder.setAttribute('hidden', '');

      // 게시판 화면 갱신
      renderOrderBoard();
    }
  }

  tabBtnOrder.addEventListener('click', () => switchTab('order'));
  tabBtnBoard.addEventListener('click', () => switchTab('board'));

  // 주문 완료 알림창에서 '게시판에서 확인하기' 버튼 클릭 시 게시판으로 이동
  if (btnGotoBoard) {
    btnGotoBoard.addEventListener('click', () => {
      switchTab('board');
      panelBoard.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }

  // ------------------------------------------------------------------------
  // 4. 실시간 예상 금액 계산 함수
  // ------------------------------------------------------------------------
  function calculateTotalPrice() {
    // 1) 음료 기본 가격 (선택되지 않은 경우 0원)
    const selectedBeveragePrice = beverageSelect.value ? parseInt(beverageSelect.value, 10) : 0;

    if (selectedBeveragePrice === 0) {
      estimatedPriceDisplay.textContent = '예상 금액: 0원';
      return 0;
    }

    // 2) 사이즈 추가 금액 (S: 0원, M: +500원, L: +1,000원)
    let sizePrice = 0;
    sizeRadios.forEach((radio) => {
      if (radio.checked) {
        sizePrice = parseInt(radio.value, 10) || 0;
      }
    });

    // 3) 추가 옵션 금액 합산
    let optionsPrice = 0;
    optionCheckboxes.forEach((checkbox) => {
      if (checkbox.checked) {
        optionsPrice += parseInt(checkbox.value, 10) || 0;
      }
    });

    // 4) 수량 (최소 1잔, 최대 10잔)
    let quantity = parseInt(quantityInput.value, 10);
    if (isNaN(quantity) || quantity < 1) {
      quantity = 1;
    } else if (quantity > 10) {
      quantity = 10;
      quantityInput.value = 10;
    }

    // 1잔당 단가 = (기본 가격 + 사이즈 추가금 + 옵션 추가금)
    const singleItemPrice = selectedBeveragePrice + sizePrice + optionsPrice;

    // 총 예상 금액 = 1잔당 단가 * 수량
    const total = singleItemPrice * quantity;

    // 천 단위 콤마 포맷(toLocaleString) 적용 후 화면에 출력
    estimatedPriceDisplay.textContent = `예상 금액: ${total.toLocaleString('ko-KR')}원`;

    return total;
  }

  // ------------------------------------------------------------------------
  // 5. 알림 메시지 표시 함수
  // ------------------------------------------------------------------------
  function showAlert(message) {
    alertBox.textContent = `⚠️ ${message}`;
    alertBox.classList.add('show');

    // 4초 후 알림 박스 자동 숨김
    setTimeout(() => {
      alertBox.classList.remove('show');
    }, 4000);
  }

  function hideAlert() {
    alertBox.classList.remove('show');
    alertBox.textContent = '';
  }

  // ------------------------------------------------------------------------
  // 6. 경과 시간 계산 함수 (실시간 주문 게시판용)
  // ------------------------------------------------------------------------
  function getRelativeTimeString(createdAtMs) {
    if (!createdAtMs) return '방금 전';
    const now = Date.now();
    const diffSeconds = Math.max(0, Math.floor((now - createdAtMs) / 1000));

    if (diffSeconds < 30) {
      return '방금 전';
    } else if (diffSeconds < 60) {
      return `${diffSeconds}초 전`;
    } else if (diffSeconds < 3600) {
      const minutes = Math.floor(diffSeconds / 60);
      return `${minutes}분 전`;
    } else if (diffSeconds < 86400) {
      const hours = Math.floor(diffSeconds / 3600);
      return `${hours}시간 전`;
    } else {
      const days = Math.floor(diffSeconds / 86400);
      return `${days}일 전`;
    }
  }

  // ------------------------------------------------------------------------
  // 7. 실시간 주문 게시판 렌더링 & 통계 계산
  // ------------------------------------------------------------------------
  function renderOrderBoard() {
    if (!boardOrdersContainer) return;

    // 1) 전체 및 상태별 카운트 계산
    const totalCount = orderHistory.length;
    const receivedCount = orderHistory.filter((o) => o.status === 'received').length;
    const makingCount = orderHistory.filter((o) => o.status === 'making').length;
    const completedCount = orderHistory.filter((o) => o.status === 'completed').length;

    // 실시간 누적 매출 합산 (천 단위 콤마)
    const totalRevenue = orderHistory.reduce((sum, o) => sum + (o.totalPrice || 0), 0);

    // 상단 통계 카드 수치 업데이트
    if (statValTotal) statValTotal.textContent = totalCount;
    if (statValReceived) statValReceived.textContent = receivedCount;
    if (statValMaking) statValMaking.textContent = makingCount;
    if (statValCompleted) statValCompleted.textContent = completedCount;
    if (boardTotalRevenue) {
      boardTotalRevenue.textContent = `${totalRevenue.toLocaleString('ko-KR')}원`;
    }

    // 탭 뱃지 및 필터 버튼 내 카운트 수치 업데이트
    if (boardLiveBadge) {
      // 대기/제조 중인 미완료 건수를 뱃지에 강조하거나 전체 건수 표시
      boardLiveBadge.textContent = totalCount;
    }
    if (filterCountAll) filterCountAll.textContent = totalCount;
    if (filterCountReceived) filterCountReceived.textContent = receivedCount;
    if (filterCountMaking) filterCountMaking.textContent = makingCount;
    if (filterCountCompleted) filterCountCompleted.textContent = completedCount;

    // 2) 필터링 및 검색 적용
    let filteredOrders = orderHistory.slice();

    // 상태 필터 적용
    if (currentFilterStatus !== 'all') {
      filteredOrders = filteredOrders.filter((order) => order.status === currentFilterStatus);
    }

    // 검색어 필터 적용 (고객명, 연락처, 음료명)
    if (currentSearchQuery.trim()) {
      const query = currentSearchQuery.trim().toLowerCase();
      filteredOrders = filteredOrders.filter((order) => {
        const nameMatch = (order.customerName || '').toLowerCase().includes(query);
        const phoneMatch = (order.customerPhone || '').toLowerCase().includes(query);
        const beverageMatch = (order.beverageName || '').toLowerCase().includes(query);
        return nameMatch || phoneMatch || beverageMatch;
      });
    }

    // 최신 주문이 상단에 위치하도록 역순 정렬
    filteredOrders.reverse();

    // 3) 주문이 없을 때의 빈 상태 화면 렌더링
    if (filteredOrders.length === 0) {
      let emptyMsg = '아직 접수된 주문 내역이 없습니다.';
      if (currentSearchQuery.trim()) {
        emptyMsg = `"${escapeHtml(currentSearchQuery)}"에 대한 검색 결과가 없습니다.`;
      } else if (currentFilterStatus !== 'all') {
        const statusMap = { received: '접수 완료', making: '제조 중', completed: '제조 완료' };
        emptyMsg = `'${statusMap[currentFilterStatus]}' 상태의 주문이 없습니다.`;
      }

      boardOrdersContainer.innerHTML = `
        <div class="board-empty-state">
          <div class="empty-state-icon" aria-hidden="true">☕</div>
          <div class="empty-state-title">주문 내역이 비어있습니다</div>
          <div class="empty-state-desc">${emptyMsg}</div>
          <button type="button" class="btn-empty-create" id="btn-empty-goto-order">
            📝 첫 주문서 작성하러 가기
          </button>
        </div>
      `;

      const btnEmptyGoto = document.getElementById('btn-empty-goto-order');
      if (btnEmptyGoto) {
        btnEmptyGoto.addEventListener('click', () => switchTab('order'));
      }
      return;
    }

    // 4) 주문 목록 카드 HTML 생성
    const cardsHtml = filteredOrders
      .map((order, index) => {
        const optDisplay = order.optionsText ? `추가옵션: ${escapeHtml(order.optionsText)}` : '추가옵션 없음';
        const isTakeout = order.diningTypeName === '포장';
        const diningClass = isTakeout ? 'takeout' : 'store';
        const diningText = isTakeout ? '🛍️ 포장' : '☕ 매장';
        const relTime = getRelativeTimeString(order.createdAtMs);

        // 상태별 뱃지 및 액션 버튼 구성
        let statusBadgeHtml = '';
        let actionButtonsHtml = '';

        if (order.status === 'received') {
          statusBadgeHtml = '<span class="order-status-badge badge-received">🟡 접수 완료</span>';
          actionButtonsHtml = `
            <button type="button" class="btn-card-action btn-action-start" data-action="make" data-id="${order.id}">
              <span>☕ 제조 시작</span>
            </button>
            <button type="button" class="btn-card-action btn-action-delete" data-action="delete" data-id="${order.id}" title="주문 취소/삭제">
              <span>✕</span>
            </button>
          `;
        } else if (order.status === 'making') {
          statusBadgeHtml = '<span class="order-status-badge badge-making">🔵 제조 중</span>';
          actionButtonsHtml = `
            <button type="button" class="btn-card-action btn-action-finish" data-action="complete" data-id="${order.id}">
              <span>🔔 제조 완료</span>
            </button>
            <button type="button" class="btn-card-action btn-action-delete" data-action="delete" data-id="${order.id}" title="주문 취소/삭제">
              <span>✕</span>
            </button>
          `;
        } else {
          statusBadgeHtml = '<span class="order-status-badge badge-completed">🟢 제조 완료</span>';
          actionButtonsHtml = `
            <button type="button" class="btn-card-action btn-action-done" data-action="archive" data-id="${order.id}">
              <span>✓ 픽업 완료</span>
            </button>
            <button type="button" class="btn-card-action btn-action-delete" data-action="delete" data-id="${order.id}" title="내역 삭제">
              <span>✕</span>
            </button>
          `;
        }

        // 고객 요청사항 영역 (요청사항이 있을 때만 렌더링)
        const requestBlock = order.requests
          ? `<div class="order-request-box">💬 요청: ${escapeHtml(order.requests)}</div>`
          : '';

        return `
          <div class="board-order-card status-${order.status}" id="card-order-${order.id}">
            <!-- 카드 상단: 주문 번호, 상대 시간, 상태 뱃지 -->
            <div class="order-card-header">
              <div class="order-num-time">
                <span class="order-number">#주문-${order.id.toString().slice(-4)}</span>
                <span class="order-rel-time" title="${order.timestamp}">${relTime}</span>
              </div>
              <div class="order-status-wrapper">
                ${statusBadgeHtml}
              </div>
            </div>

            <!-- 고객 정보 및 이용 방식 -->
            <div class="order-customer-row">
              <div class="order-customer-info">
                <span class="customer-name-bold">${escapeHtml(order.customerName)}님</span>
                <span class="customer-phone-sub">(${escapeHtml(order.customerPhone || '연락처 미기재')})</span>
              </div>
              <span class="dining-pill ${diningClass}">${diningText}</span>
            </div>

            <!-- 주문 음료 및 세부 옵션 -->
            <div class="order-beverage-row">
              <div class="order-beverage-details">
                <div class="beverage-main">
                  <span>${escapeHtml(order.beverageName)}</span>
                  <span class="size-pill">${order.sizeName}</span>
                </div>
                <div class="options-sub">${optDisplay}</div>
              </div>
              <span class="quantity-badge">${order.quantity}잔</span>
            </div>

            <!-- 고객 요청사항 -->
            ${requestBlock}

            <!-- 카드 하단: 결제 금액 및 실시간 상태 변경 버튼 -->
            <div class="order-card-footer">
              <span class="order-price-val">${(order.totalPrice || 0).toLocaleString('ko-KR')}원</span>
              <div class="order-status-actions">
                ${actionButtonsHtml}
              </div>
            </div>
          </div>
        `;
      })
      .join('');

    boardOrdersContainer.innerHTML = cardsHtml;
  }

  // ------------------------------------------------------------------------
  // 8. 주문 상태 변경 및 삭제 핸들러 (이벤트 위임)
  // ------------------------------------------------------------------------
  if (boardOrdersContainer) {
    boardOrdersContainer.addEventListener('click', (e) => {
      const actionBtn = e.target.closest('[data-action]');
      if (!actionBtn) return;

      const action = actionBtn.getAttribute('data-action');
      const orderId = parseInt(actionBtn.getAttribute('data-id'), 10);
      if (!orderId) return;

      const targetOrder = orderHistory.find((o) => o.id === orderId);
      if (!targetOrder) return;

      if (action === 'make') {
        // [제조 시작] -> 상태를 'making'으로 변경
        targetOrder.status = 'making';
        updateAllUI();
        updateOrderStatusInSupabase(orderId, 'making');
      } else if (action === 'complete') {
        // [제조 완료] -> 상태를 'completed'로 변경
        targetOrder.status = 'completed';
        updateAllUI();
        updateOrderStatusInSupabase(orderId, 'completed');
      } else if (action === 'archive') {
        // [픽업 완료]
        targetOrder.status = 'completed';
        updateAllUI();
        updateOrderStatusInSupabase(orderId, 'completed');
      } else if (action === 'delete') {
        // 주문 삭제
        if (confirm(`${targetOrder.customerName}님의 주문(#${targetOrder.id.toString().slice(-4)})을 목록에서 삭제하시겠습니까?`)) {
          orderHistory = orderHistory.filter((o) => o.id !== orderId);
          updateAllUI();
          deleteOrderFromSupabase(orderId);
        }
      }
    });
  }

  // ------------------------------------------------------------------------
  // 9. 주문서 화면의 최근 주문 목록 및 뱃지 업데이트
  // ------------------------------------------------------------------------
  function updateOrderFormSummaryUI() {
    if (orderCountBadge) {
      orderCountBadge.textContent = `${orderHistory.length}건`;
    }

    if (btnExcel) {
      btnExcel.disabled = orderHistory.length === 0;
    }

    if (recentOrdersList) {
      if (orderHistory.length === 0) {
        recentOrdersList.innerHTML = `
          <div style="padding: 16px; text-align: center; color: #888888; font-size: 13px;">
            아직 접수된 주문 내역이 없습니다.
          </div>
        `;
        return;
      }

      // 최신 주문이 상단에 오도록 역순 정렬하여 렌더링
      const listHtml = orderHistory
        .slice()
        .reverse()
        .slice(0, 5) // 최근 5건만 간편 표시
        .map((order) => {
          const optStr = order.optionsText ? ` (${order.optionsText})` : '';
          const diningBadge = order.diningTypeName === '포장' ? '🛍️ [포장]' : '☕ [매장]';
          return `
            <div class="recent-order-item" id="order-item-${order.id}">
              <div class="recent-order-details">
                <span class="recent-order-name">${escapeHtml(order.customerName)}님 - ${diningBadge} ${escapeHtml(order.beverageName)} ${order.sizeName}${escapeHtml(optStr)} ${order.quantity}잔</span>
                <span class="recent-order-desc">${order.timestamp} | ${getRelativeTimeString(order.createdAtMs)}</span>
              </div>
              <span class="recent-order-price">${(order.totalPrice || 0).toLocaleString('ko-KR')}원</span>
            </div>
          `;
        })
        .join('');

      recentOrdersList.innerHTML = listHtml;
    }
  }

  // 전체 UI 일괄 동기화 함수
  function updateAllUI() {
    updateOrderFormSummaryUI();
    renderOrderBoard();
  }

  // HTML 특수문자 이스케이프 함수
  function escapeHtml(text) {
    if (!text) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // ------------------------------------------------------------------------
  // 10. 색종이 폭죽(Confetti) 애니메이션 시스템
  // ------------------------------------------------------------------------
  let confettiAnimId = null;

  function triggerConfetti() {
    if (!confettiCanvas) return;

    const ctx = confettiCanvas.getContext('2d');
    if (!ctx) return;

    const width = (confettiCanvas.width = window.innerWidth);
    const height = (confettiCanvas.height = window.innerHeight);

    if (confettiAnimId) {
      cancelAnimationFrame(confettiAnimId);
    }

    const colors = [
      '#1ed760', // Spotify Green
      '#ffffff', // White
      '#1db954', // Spotify Green Variant
      '#539df5', // Blue
      '#ffa42b', // Orange
      '#f3727f', // Red
      '#b3b3b3', // Silver
    ];

    const particleCount = 100;
    const particles = [];

    for (let i = 0; i < particleCount; i++) {
      particles.push({
        x: width * (0.2 + Math.random() * 0.6),
        y: height * 0.45,
        w: Math.random() * 8 + 6,
        h: Math.random() * 5 + 4,
        color: colors[Math.floor(Math.random() * colors.length)],
        vx: (Math.random() - 0.5) * 14,
        vy: -(Math.random() * 12 + 6),
        gravity: 0.35,
        rotation: Math.random() * 360,
        rotationSpeed: (Math.random() - 0.5) * 12,
        opacity: 1,
        decay: Math.random() * 0.005 + 0.008,
      });
    }

    function renderConfetti() {
      ctx.clearRect(0, 0, width, height);
      let activeParticles = 0;

      for (let i = 0; i < particles.length; i++) {
        const p = particles[i];
        if (p.opacity <= 0 || p.y > height + 20) continue;

        activeParticles++;
        p.vy += p.gravity;
        p.x += p.vx;
        p.y += p.vy;
        p.rotation += p.rotationSpeed;
        p.opacity -= p.decay;

        ctx.save();
        ctx.translate(p.x, p.y);
        ctx.rotate((p.rotation * Math.PI) / 180);
        ctx.globalAlpha = Math.max(0, p.opacity);
        ctx.fillStyle = p.color;
        ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
        ctx.restore();
      }

      if (activeParticles > 0) {
        confettiAnimId = requestAnimationFrame(renderConfetti);
      } else {
        ctx.clearRect(0, 0, width, height);
        confettiAnimId = null;
      }
    }

    renderConfetti();
  }

  window.addEventListener('resize', () => {
    if (confettiCanvas) {
      confettiCanvas.width = window.innerWidth;
      confettiCanvas.height = window.innerHeight;
    }
  });

  // ------------------------------------------------------------------------
  // 11. 주문 폼 실시간 이벤트 리스너 등록
  // ------------------------------------------------------------------------
  beverageSelect.addEventListener('change', () => {
    hideAlert();
    calculateTotalPrice();
  });

  sizeRadios.forEach((radio) => {
    radio.addEventListener('change', calculateTotalPrice);
  });

  optionCheckboxes.forEach((checkbox) => {
    checkbox.addEventListener('change', calculateTotalPrice);
  });

  quantityInput.addEventListener('input', calculateTotalPrice);
  quantityInput.addEventListener('change', calculateTotalPrice);
  customerNameInput.addEventListener('input', hideAlert);

  // ------------------------------------------------------------------------
  // 12. 주문하기 (Submit) 처리
  // ------------------------------------------------------------------------
  orderForm.addEventListener('submit', (e) => {
    e.preventDefault();

    const customerName = customerNameInput.value.trim();
    const customerPhone = customerPhoneInput.value.trim();
    const selectedBeverageValue = beverageSelect.value;

    // [유효성 검사 1] 이름이 비어있으면 알림
    if (!customerName) {
      showAlert('이름을 입력해주세요');
      customerNameInput.focus();
      return;
    }

    // [유효성 검사 2] 음료를 선택하지 않았으면 알림
    if (!selectedBeverageValue) {
      showAlert('음료를 선택해주세요');
      beverageSelect.focus();
      return;
    }

    hideAlert();

    // 1) 포장/매장 선택 여부 수집
    let diningTypeName = '포장';
    let diningTypeValue = 'takeout';
    diningTypeRadios.forEach((radio) => {
      if (radio.checked) {
        diningTypeName = radio.getAttribute('data-name') || '포장';
        diningTypeValue = radio.value;
      }
    });

    // 2) 선택된 음료 정보 수집
    const selectedOption = beverageSelect.options[beverageSelect.selectedIndex];
    const selectedBeverageText = selectedOption.getAttribute('data-name');
    const beveragePrice = parseInt(beverageSelect.value, 10);

    // 3) 선택된 사이즈 확인
    let selectedSizeName = 'M';
    let sizePrice = 500;
    sizeRadios.forEach((radio) => {
      if (radio.checked) {
        selectedSizeName = radio.getAttribute('data-name');
        sizePrice = parseInt(radio.value, 10) || 0;
      }
    });

    // 4) 선택된 옵션 목록 및 가격 수집
    const selectedOptionNames = [];
    let optionsPrice = 0;
    optionCheckboxes.forEach((checkbox) => {
      if (checkbox.checked) {
        selectedOptionNames.push(checkbox.getAttribute('data-name'));
        optionsPrice += parseInt(checkbox.value, 10) || 0;
      }
    });

    const quantity = parseInt(quantityInput.value, 10) || 1;
    const totalPrice = calculateTotalPrice();
    const requests = requestsTextarea.value.trim();
    const optionsText = selectedOptionNames.length > 0 ? selectedOptionNames.join(', ') : '';
    const optionsDisplay = optionsText ? ` (${optionsText})` : '';

    // 주문 확인 메시지 조합
    const formattedMessage = `${customerName}님 [${diningTypeName}], ${selectedBeverageText} ${selectedSizeName}사이즈${optionsDisplay} ${quantity}잔, 총 ${totalPrice.toLocaleString('ko-KR')}원 주문이 접수되었습니다!`;

    // 화면 하단 주문 확인 박스 표시
    confirmationText.textContent = formattedMessage;
    confirmationBox.classList.add('show');

    // 색종이 폭죽 애니메이션 발사
    triggerConfetti();

    // 주문 확인 영역으로 부드럽게 스크롤
    confirmationBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    // 신규 주문 객체 생성 (UUID id, 실시간 상태 'received' 부여)
    const now = new Date();
    const timestamp = now.toLocaleString('ko-KR');

    const newOrder = {
      id: crypto.randomUUID(),  // UUID (cafe_menu.id 타입과 일치)
      createdAtMs: now.getTime(),
      timestamp,
      customerName,
      customerPhone,
      diningTypeName,
      diningTypeValue,
      beverageName: selectedBeverageText,
      beveragePrice,
      sizeName: selectedSizeName,
      sizePrice,
      optionsText,
      optionsPrice,
      quantity,
      totalPrice,
      requests,
      status: 'received',
    };

    orderHistory.push(newOrder);
    updateAllUI();
    saveOrderToSupabase(newOrder);
  });

  // ------------------------------------------------------------------------
  // 13. 다시 작성 (초기화) 버튼 처리
  // ------------------------------------------------------------------------
  btnReset.addEventListener('click', () => {
    orderForm.reset();

    // 포장 기본 선택
    const defaultDiningRadio = document.getElementById('dining-takeout');
    if (defaultDiningRadio) defaultDiningRadio.checked = true;

    // 음료 기본값: 선택 안함
    beverageSelect.selectedIndex = 0;

    // 사이즈 기본값: M사이즈 (+500원)
    const defaultSizeRadio = document.getElementById('size-m');
    if (defaultSizeRadio) defaultSizeRadio.checked = true;

    // 옵션 체크박스 전부 해제
    optionCheckboxes.forEach((checkbox) => {
      checkbox.checked = false;
    });

    // 수량 기본값 1
    quantityInput.value = 1;

    calculateTotalPrice();
    hideAlert();
    confirmationBox.classList.remove('show');
    confirmationText.textContent = '';
    customerNameInput.focus();
  });

  // ------------------------------------------------------------------------
  // 14. 실시간 게시판 필터 및 검색 이벤트
  // ------------------------------------------------------------------------
  filterBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      filterBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      currentFilterStatus = btn.getAttribute('data-status') || 'all';
      renderOrderBoard();
    });
  });

  if (boardSearchInput) {
    boardSearchInput.addEventListener('input', (e) => {
      currentSearchQuery = e.target.value;
      if (btnClearSearch) {
        btnClearSearch.hidden = !currentSearchQuery;
      }
      renderOrderBoard();
    });
  }

  if (btnClearSearch) {
    btnClearSearch.addEventListener('click', () => {
      if (boardSearchInput) {
        boardSearchInput.value = '';
        currentSearchQuery = '';
        btnClearSearch.hidden = true;
        renderOrderBoard();
        boardSearchInput.focus();
      }
    });
  }

  // ------------------------------------------------------------------------
  // 15. 실시간 테스트 주문 추가 기능 (편의 기능)
  // ------------------------------------------------------------------------
  if (btnAddSampleOrder) {
    btnAddSampleOrder.addEventListener('click', () => {
      const sampleNames = ['김서연', '이준호', '박민지', '정우성', '최수빈', '강하늘'];
      const sampleBeverages = [
        { name: '아메리카노', price: 3500 },
        { name: '카페라떼', price: 4000 },
        { name: '바닐라라떼', price: 4500 },
        { name: '녹차라떼', price: 4500 },
      ];
      const sampleSizes = [
        { name: 'S', price: 0 },
        { name: 'M', price: 500 },
        { name: 'L', price: 1000 },
      ];
      const sampleOptions = ['샷 추가', '크림 추가', '시럽 추가', '디카페인'];
      const sampleDining = [
        { name: '포장', val: 'takeout' },
        { name: '매장', val: 'store' },
      ];

      const randomName = sampleNames[Math.floor(Math.random() * sampleNames.length)];
      const randomBev = sampleBeverages[Math.floor(Math.random() * sampleBeverages.length)];
      const randomSize = sampleSizes[Math.floor(Math.random() * sampleSizes.length)];
      const randomDining = sampleDining[Math.floor(Math.random() * sampleDining.length)];
      const randomOpt = Math.random() > 0.4 ? sampleOptions[Math.floor(Math.random() * sampleOptions.length)] : '';
      const optPrice = randomOpt === '시럽 추가' ? 300 : randomOpt ? 500 : 0;
      const quantity = Math.floor(Math.random() * 2) + 1;
      const totalPrice = (randomBev.price + randomSize.price + optPrice) * quantity;

      const now = new Date();
      const timestamp = now.toLocaleString('ko-KR');

      const sampleOrder = {
        id: crypto.randomUUID(),
        createdAtMs: now.getTime(),
        timestamp,
        customerName: randomName,
        customerPhone: `010-${Math.floor(1000 + Math.random() * 9000)}-${Math.floor(1000 + Math.random() * 9000)}`,
        diningTypeName: randomDining.name,
        diningTypeValue: randomDining.val,
        beverageName: randomBev.name,
        beveragePrice: randomBev.price,
        sizeName: randomSize.name,
        sizePrice: randomSize.price,
        optionsText: randomOpt,
        optionsPrice: optPrice,
        quantity,
        totalPrice,
        requests: Math.random() > 0.5 ? '덜 달게 부탁드려요.' : '',
        status: 'received',
      };

      orderHistory.push(sampleOrder);
      updateAllUI();
      saveOrderToSupabase(sampleOrder);
    });
  }

  // ------------------------------------------------------------------------
  // 16. 전체 주문 내역 비우기
  // ------------------------------------------------------------------------
  if (btnClearAllOrders) {
    btnClearAllOrders.addEventListener('click', () => {
      if (orderHistory.length === 0) {
        alert('삭제할 주문 내역이 없습니다.');
        return;
      }
      if (confirm('저장된 모든 주문 내역을 초기화하시겠습니까? (이 작업은 되돌릴 수 없습니다)')) {
        orderHistory = [];
        updateAllUI();
        deleteAllOrdersFromSupabase();
      }
    });
  }

  // ------------------------------------------------------------------------
  // 17. 엑셀로 주문 내역 저장받기 (UTF-8 BOM CSV 파일 다운로드)
  // ------------------------------------------------------------------------
  function downloadOrdersAsExcel() {
    if (orderHistory.length === 0) {
      alert('저장할 주문 내역이 없습니다. 먼저 주문을 접수해주세요!');
      return;
    }

    const BOM = '\uFEFF';
    const statusLabelMap = {
      received: '접수 완료',
      making: '제조 중',
      completed: '제조 완료',
    };

    // CSV 헤더 컬럼 정의 (진행상태 컬럼 포함)
    const headers = [
      '주문번호',
      '주문일시',
      '진행상태',
      '고객명',
      '전화번호',
      '이용방식',
      '주문음료',
      '사이즈',
      '추가옵션',
      '수량',
      '총금액(원)',
      '요청사항',
    ];

    // 행 데이터 생성 (쉼표 및 큰따옴표 이스케이프 처리)
    const rows = orderHistory.map((order, index) => {
      const statusKorean = statusLabelMap[order.status] || '접수 완료';
      return [
        index + 1,
        `"${order.timestamp}"`,
        `"${statusKorean}"`,
        `"${(order.customerName || '').replace(/"/g, '""')}"`,
        `"${(order.customerPhone || '').replace(/"/g, '""')}"`,
        `"${order.diningTypeName || '포장'}"`,
        `"${(order.beverageName || '').replace(/"/g, '""')}"`,
        `"${order.sizeName || 'M'}"`,
        `"${order.optionsText ? order.optionsText.replace(/"/g, '""') : '없음'}"`,
        order.quantity || 1,
        order.totalPrice || 0,
        `"${(order.requests || '').replace(/"/g, '""')}"`,
      ].join(',');
    });

    const csvContent = BOM + [headers.join(','), ...rows].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);

    const downloadLink = document.createElement('a');
    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}`;

    downloadLink.href = url;
    downloadLink.setAttribute('download', `바이브카페_주문내역_${dateStr}.csv`);
    document.body.appendChild(downloadLink);
    downloadLink.click();

    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(url);
  }

  if (btnExcel) {
    btnExcel.addEventListener('click', downloadOrdersAsExcel);
  }

  if (btnBoardExcel) {
    btnBoardExcel.addEventListener('click', downloadOrdersAsExcel);
  }

  // ------------------------------------------------------------------------
  // 18. 경과 시간 주기적 자동 갱신 (15초 주기)
  // ------------------------------------------------------------------------
  setInterval(() => {
    // 게시판 패널이 열려있을 때 상대 시간 자동 갱신
    if (!panelBoard.hasAttribute('hidden')) {
      renderOrderBoard();
    }
  }, 15000);

  // ------------------------------------------------------------------------
  // 19. 초기 구동 실행
  // ------------------------------------------------------------------------
  calculateTotalPrice(); // 초기 예상 금액 (0원)
  updateAllUI();          // 저장된 주문 데이터 화면 반영
});
