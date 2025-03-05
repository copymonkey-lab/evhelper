// ==UserScript==
// @name         evhelper
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  evhelper
// @author       victor
// @match        https://envycube.com/order_make_pk_list.php
// @grant        none
// @require      https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js
// @require      https://apis.google.com/js/api.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jsrsasign/8.0.20/jsrsasign-all-min.js
// @license      MIT
// ==/UserScript==

(function() {
    'use strict';

    // 전역 변수 선언
    let patternMap = new Map();
    let orderDataCache = new Map();
    const sellerInfoMap = new Map();
    let sortedPatterns = [];
    let processedWaybills = new Set();
    let lastProcessedDate = '';

    // 구글 시트 API 설정
    const GOOGLE_API_CONFIG = {
        apiKey: 'AIzaSyBwMx9Lw_RrGKWBJAD7Lo5AH6YhkVbPuXk',  // 웹 API 키 (Google Cloud Console에서 생성)
        clientId: '113068596238740059939',
        serviceAccount: {
            private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDVABvLDtVTMdxw\ne8Qk8TFUVjWl3K1nrs/EV1L+gwMb3Qzkw2cRHYMuEqdZTKys1C4XVXLFpLqBWk3C\nzalvZAb6DJEv25ae27pERy2Nx6mogN85aH558Gxdh+2IoKKTGef+hqfYHfKHajtF\n4wpHUDqnKlQyK3by8S7l7WO4VsEZFMfUyT1vNXPrDa82PuPfqFYwujzoYkVO/yji\n76h9vaU/kJyvFUWQx827uqvWW3DkZ7WpYuJiTBZjLYhEwJi0diJSf5t9xV0dRVM3\nvAuzRKBkLc22XIaq3VnhyGoyK9aEfB3NMMblz+T+ShyungGAxOanBq7A76p9bw0O\ng/HwdKwvAgMBAAECggEAIe60/snbqTfJxu+opTv5YFalkElSJLDgL4a71Yj4j1FN\nUwpgGoVplwbouxywa44X06bMtHjUL3Q77BtIcVLtm5sx6/5fBeq6R1NRigMzX4E8\nTpB7iaCIGvRjHn98ttOLNmysQ40tzG3biHwtcIPy/BuNszpiZjyO/JkvaDgF5iEz\ncSNao0tsSWC2c29g+0BnVBERVoR+1jwVqqfa7cZGhqWXtmlmC2W1657wSYpBPZlr\nVMlAazoi22yVQvO1r78L+PMkgJ365h7SYsVPzwEQWgT1DQTaaX04kz/gnI3psvtd\ntLZIu8NmHdVIdKbeTJrQAPhHIu9GvpZq0t3nktpmTQKBgQDvwm97a+dfFdv6w6GD\nR99/b/VKP2IiEOQYa0aam4VbL0aIhlEU0i/JgIwCUlyu1EY4yePIScendel7upfq\nJku1vI10qBCJI85pNlUr9VUcPRHVLaHHsOKG0eZv8kyn6roh5JUQ3qxxQKOrKqA0\nAxsOoOdBX+RX8tI/j20dVTX2PQKBgQDjbafMRVXp2rLqq3sG6H6dfM1+03kcMyKO\nvS+eC5R26ZTxAoZdU6G7Gs6k/gFQQeofoY6kfJHFIyZWrLuQnhw1ov/7ans0s/g+\nf6/Odb1p0PzgvPWPBPlEEA42/8z7Z9iMy3dyjEc96fjr2NKDlqvHO634mQzAoHr0\nhnCGlw1+2wKBgHyco/CT3ocvB0xIDVP8MQ89E1Hpq4llGggPCX0lw6Pm6FPg65dU\nvv2N0DcMs5syPOUbGUZqAljpEdb63iYWjVcBjsvI5f9BGvDYCmB0fC3XF8Oimej9\n6F6GDay1VF4Zw3AGK+u+sAWUwPwfhXBDBPcPbeIugrGrRNdAJkgOl6NJAoGBALKb\n8rT1GwTekbbE14jUXGO4mPZqhGnGKvSo1VWsyHse9K7WicmPnauA4RsotMVgDsuq\nqIi7oAuPkFNvsppf4c2p5pl/xaTdVi9XPi3Jv+jzjTW+kKcyg8SVS2ScPlKO+r2Q\nKY3XZzfToX8vuBxJ3zxHvVhIcoBxSD7zujmpNZsTAoGASnUJ3gZW0o37HKxxT2Hd\n+TOyCumyIQTaE+w0MImw+En0XpdSF2tgXR69Sr2oQR0VdjjhFuqHuIZhQdnPz0dy\nMRGX2ptUvSE+DgEPI4MjUN5mLq1sMs/h8yzLExgXdMS/YCD8e7BmPlFpavB+btoH\nw99FVQUR1gGDLosbhmBqZGA=\n-----END PRIVATE KEY-----\n",  // -----BEGIN PRIVATE KEY----- 로 시작하는 긴 문자열
            client_email: "id-196@secure-potion-408002.iam.gserviceaccount.com"
        },
        spreadsheetId: '1tYTAvi5WoqvYWzIw2gAhwzIIo4VuI3pSC7-xJ0YpxAA',
        scope: 'https://www.googleapis.com/auth/spreadsheets',
        dailySpreadsheetId: '1shS_nn0pJZf8gwb8oAWPAPKuVh3OobTEG9Y4LfrG_lo',
    };

    // 프로그레스 바 UI 생성
    function createProgressBar() {
        const progressContainer = document.createElement('div');
        progressContainer.id = 'progress-container';
        progressContainer.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 300px;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            z-index: 10000;
            display: none;
        `;

        const progressText = document.createElement('div');
        progressText.id = 'progress-text';
        progressText.style.cssText = `
            margin-bottom: 10px;
            text-align: center;
            font-size: 14px;
        `;

        const progressBarOuter = document.createElement('div');
        progressBarOuter.style.cssText = `
            width: 100%;
            height: 20px;
            background: #f0f0f0;
            border-radius: 10px;
            overflow: hidden;
        `;

        const progressBarInner = document.createElement('div');
        progressBarInner.id = 'progress-bar';
        progressBarInner.style.cssText = `
            width: 0%;
            height: 100%;
            background: #4CAF50;
            transition: width 0.3s ease;
        `;

        progressBarOuter.appendChild(progressBarInner);
        progressContainer.appendChild(progressText);
        progressContainer.appendChild(progressBarOuter);
        document.body.appendChild(progressContainer);

        return {
            container: progressContainer,
            text: progressText,
            bar: progressBarInner,
            show: () => progressContainer.style.display = 'block',
            hide: () => progressContainer.style.display = 'none',
            update: (progress, text) => {
                progressBarInner.style.width = `${progress}%`;
                progressText.textContent = text;
            }
        };
    }

    // 출고건 확인 버튼 추가
    function addCheckOrderButton() {
        const button = document.createElement('button');
        button.textContent = '출고데이터 확인'
        button.style.cssText = `
            position: fixed;
            top: 7px;
            right: 250px;
            z-index: 9999;
            padding: 8px 16px;
            background-color: #2196F3;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        `;

        button.addEventListener('click', showOrderCount);
        document.body.appendChild(button);
    }

    // 주문건수 확인 창 생성
    function createOrderCountDialog() {
        const dialog = document.createElement('div');
        dialog.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            z-index: 10000;
            width: 600px;
        `;

        dialog.innerHTML = `
            <h3 style="margin-bottom: 15px; text-align: center;">판매처별 출고건수</h3>
            <div style="margin-bottom: 10px;">
                <label style="display: flex; align-items: center; gap: 5px;">
                    <input type="checkbox" id="select-all-sellers">
                    <span>전체 선택</span>
                </label>
            </div>
            <div id="order-count-list" style="
                max-height: 400px;
                overflow-y: auto;
                margin-bottom: 20px;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 10px;
            ">
                <div style="text-align: center; color: #666;">데이터 로딩 중...</div>
            </div>
            <div id="total-count" style="
                text-align: right;
                padding: 10px;
                border-top: 1px solid #ddd;
                font-weight: bold;
            ">
                총 주문건수: 0건
            </div>
            <div style="text-align: center; margin-top: 20px;">
                <button id="process-all-orders" style="
                    padding: 10px 20px;
                    background: #4CAF50;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                    margin-right: 10px;
                ">출고(전체)</button>
                <button id="process-selected-orders" style="
                    padding: 10px 20px;
                    background: #2196F3;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                    margin-right: 10px;
                ">출고(선택)</button>
                <button id="close-order-count" style="
                    padding: 10px 20px;
                    background: #f44336;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                ">취소</button>
            </div>
        `;

        return dialog;
    }

    // 주문건수 조회 및 표시
    async function showOrderCount() {
        const dialog = createOrderCountDialog();
        document.body.appendChild(dialog);
        const progressBar = createProgressBar();

        try {
            progressBar.show();
            progressBar.update(0, '판매처 목록 조회 중...');

            // 검색할 페이지 범위 지정
            const pagesToLoad = [1, 2];  // 판매처 목록의 페이지 범위
            let allMembers = [];

            // 모든 페이지의 판매처 정보 수집
            for (const currentPage of pagesToLoad) {
                const pageParam = currentPage === 1 ? 1 : currentPage.toString();
                const memberResponse = await fetch('https://envycube.com/web2/ajax/proc/loadOrder/get_member_list.php', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({
                        action: 'get_member_list',
                        page: pageParam,
                        search_txt: ''
                    })
                });

                const memberData = await memberResponse.json();
                if (memberData.MSG === 'OK') {
                    allMembers = allMembers.concat(memberData.DAT.res_result);
                }
                await new Promise(resolve => setTimeout(resolve, 100));
            }

            const orderCountList = document.getElementById('order-count-list');
            let totalOrderCount = 0;
            let orderCountHtml = '';

            orderDataCache.clear();
            const totalMembers = allMembers.length;

            for (let i = 0; i < totalMembers; i++) {
                const member = allMembers[i];
                progressBar.update(
                    (i / totalMembers) * 100,
                    `판매처 주문 조회 중... (${i + 1}/${totalMembers})`
                );

                const now = new Date();
                const koreaTime = new Date(now.getTime() + (9 * 60 * 60 * 1000));
                const today = koreaTime.toISOString().split('T')[0];

                // 30일 전 날짜 계산
                const thirtyDaysAgo = new Date(koreaTime);
                thirtyDaysAgo.setDate(koreaTime.getDate() - 30);
                const startDate = thirtyDaysAgo.toISOString().split('T')[0];

                const orderResponse = await fetch('https://envycube.com/web2/ajax/proc/loadOrder/loadOrder.php', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({
                        action: 'loadOrderList',
                        s_date: startDate,  // 30일 전 날짜
                        e_date: today,      // 오늘 날짜
                        e_hour: '23',
                        e_minute: '59',
                        e_second: '59',
                        tradecode: member.idx,
                        tradename: `${member.MBI_CODE} ${member.MBI_0032} ${member.MBI_0033}`,
                        item_type: '1',
                        limit_select: '10000',
                        page: '1',
                        data_sort: 'cnt_asc',
                        searchType: 'OSI_0015',
                        printYN: 'N',
                        packingClear: 'N'
                    })
                });

                const orderData = await orderResponse.json();
                const orderCount = orderData.MSG === 'OK' ? parseInt(orderData.res_page.total_count) : 0;

                if (orderCount > 0) {
                    orderDataCache.set(member.MBI_CODE, {
                        memberInfo: member,
                        orders: orderData.resResult
                    });

                    totalOrderCount += orderCount;
                    orderCountHtml += generateOrderListItem(member, orderCount);
                }

                await new Promise(resolve => setTimeout(resolve, 100));
            }

            orderCountList.innerHTML = orderCountHtml || '<div style="text-align: center; color: #666;">출고 대기 건이 없습니다.</div>';
            document.getElementById('total-count').textContent = `총 주문건수: ${totalOrderCount}건`;
        } catch (error) {
            console.error('주문건수 조회 중 오류:', error);
            orderCountList.innerHTML = '<div style="text-align: center; color: red;">데이터 로드 실패</div>';
        } finally {
            progressBar.hide();
            setupOrderDialogHandlers(dialog);
        }
    }

    // 주문 목록 아이템 HTML 생성
    function generateOrderListItem(member, orderCount) {
        return `
            <div style="
                padding: 10px;
                border-bottom: 1px solid #eee;
                display: flex;
                justify-content: space-between;
                align-items: center;
            ">
                <label style="display: flex; align-items: center; gap: 10px;">
                    <input type="checkbox" class="seller-checkbox"
                        data-mbi-code="${member.MBI_CODE}"
                        data-company="${member.MBI_0032}">
                    <span style="flex: 1;">${member.MBI_0032}</span>
                </label>
                <span style="
                    background: #4CAF50;
                    color: white;
                    padding: 4px 8px;
                    border-radius: 12px;
                    font-size: 12px;
                ">${orderCount}건</span>
            </div>
        `;
    }

    // localStorage에서 데이터 로드하는 함수 추가
    function loadProcessedWaybills() {
        const currentDate = new Date().toDateString();
        const storedDate = localStorage.getItem('lastProcessedDate');
        const storedWaybills = localStorage.getItem('processedWaybills');

        if (storedDate === currentDate && storedWaybills) {
            processedWaybills = new Set(JSON.parse(storedWaybills));
            lastProcessedDate = storedDate;
        } else {
            processedWaybills = new Set();
            lastProcessedDate = currentDate;
            localStorage.setItem('lastProcessedDate', currentDate);
            localStorage.setItem('processedWaybills', JSON.stringify([]));
        }
    }

    // 구글 시트에 데이터 백업
    let cachedToken = null;
    let tokenExpireTime = 0;
    const BATCH_SIZE = 1000; // 배치 크기 설정
    let lastKnownRow = null;

    // base64URLEncode 함수 추가
    function base64URLEncode(str) {
        return btoa(str)
            .replace(/\+/g, '-')
            .replace(/\//g, '_')
            .replace(/=/g, '');
    }

    async function backupToGoogleSheets(orders) {
        try {
            // 현재 날짜 확인 및 데이터 로드
            const currentDate = new Date().toDateString();
            if (currentDate !== lastProcessedDate) {
                processedWaybills.clear();
                lastProcessedDate = currentDate;
                localStorage.setItem('lastProcessedDate', currentDate);
                localStorage.setItem('processedWaybills', JSON.stringify([]));
            }

            // 중복되지 않은 주문만 필터링
            const uniqueOrders = orders.filter(order => !processedWaybills.has(order.FWD_0006));

            if (uniqueOrders.length === 0) {
                console.log('모든 주문이 이미 오늘 백업되었습니다.');
                return;
            }

            // 액세스 토큰 가져오기
            const accessToken = await getAccessToken();

            // 전체 시트 데이터를 가져옴
            const existingDataResponse = await fetch(
                `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_API_CONFIG.spreadsheetId}/values/백업!A:J`,
                {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`
                    }
                }
            );

            const existingData = await existingDataResponse.json();
            const existingValues = existingData.values || [];

            // 운송장번호를 키로 하는 Map 생성 (행 번호 저장)
            const waybillRowMap = new Map();
            existingValues.forEach((row, index) => {
                if (row[4]) { // E열(운송장번호)
                    if (!waybillRowMap.has(row[4])) {
                        waybillRowMap.set(row[4], []);
                    }
                    waybillRowMap.get(row[4]).push(index + 1); // 1-based index
                }
            });

            // 현재 시간
            const now = new Date().toLocaleString('ko-KR', {
                timeZone: 'Asia/Seoul',
                year: 'numeric',
                month: 'numeric',
                day: 'numeric',
                hour: 'numeric',
                minute: 'numeric',
                second: 'numeric',
                hour12: true
            });

            // 데이터 배치 처리 준비
            const updateBatches = [];
            const newBatches = [];
            let currentRow = existingValues.length + 1;

            // 데이터 변환 및 배치 구성
            uniqueOrders.forEach(order => {
                const waybillNumber = order.FWD_0006;

                order.itemInfos.forEach(item => {
                    const sellerInfo = orderDataCache.get(order.MBI);
                    const sellerName = sellerInfo ? sellerInfo.memberInfo.MBI_0032 : order.MBI;
                    const row = [
                        now,
                        sellerName,
                        order.MBI,
                        order.SignDate,
                        waybillNumber,
                        order.idx,
                        order.OSI_0015,
                        item.prdCode,
                        item.prdName,
                        item.itemCnt
                    ];

                    // 기존 데이터가 있는 경우 시간만 업데이트
                    if (waybillRowMap.has(waybillNumber)) {
                        const existingRows = waybillRowMap.get(waybillNumber);
                        existingRows.forEach(rowNumber => {
                            updateBatches.push({
                                range: `백업!A${rowNumber}`,
                                values: [[now]]
                            });
                        });
                    } else {
                        // 새로운 데이터 추가
                        newBatches.push({
                            range: `백업!A${currentRow}:J${currentRow}`,
                            values: [row]
                        });
                        currentRow++;
                    }
                });

                // 처리된 운송장 번호 기록
                processedWaybills.add(waybillNumber);
            });

            // 업데이트 배치 처리 (기존 데이터 시간 업데이트)
            if (updateBatches.length > 0) {
                for (let i = 0; i < updateBatches.length; i += BATCH_SIZE) {
                    const batch = updateBatches.slice(i, i + BATCH_SIZE);
                    const batchUpdateRequest = {
                        data: batch,
                        valueInputOption: 'USER_ENTERED'
                    };

                    const updateResponse = await fetch(
                        `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_API_CONFIG.spreadsheetId}/values:batchUpdate`,
                        {
                            method: 'POST',
                            headers: {
                                'Authorization': `Bearer ${accessToken}`,
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify(batchUpdateRequest)
                        }
                    );

                    if (!updateResponse.ok) {
                        throw new Error('기존 데이터 업데이트 실패');
                    }

                    // API 레이트 리밋을 위한 짧은 대기
                    if (i + BATCH_SIZE < updateBatches.length) {
                        await new Promise(resolve => setTimeout(resolve, 100));
                    }
                }
            }

            // 새로운 데이터 배치 처리
            if (newBatches.length > 0) {
                for (let i = 0; i < newBatches.length; i += BATCH_SIZE) {
                    const batch = newBatches.slice(i, i + BATCH_SIZE);
                    const batchUpdateRequest = {
                        data: batch,
                        valueInputOption: 'USER_ENTERED'
                    };

                    const newDataResponse = await fetch(
                        `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_API_CONFIG.spreadsheetId}/values:batchUpdate`,
                        {
                            method: 'POST',
                            headers: {
                                'Authorization': `Bearer ${accessToken}`,
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify(batchUpdateRequest)
                        }
                    );

                    if (!newDataResponse.ok) {
                        throw new Error('새로운 데이터 추가 실패');
                    }

                    // API 레이트 리밋을 위한 짧은 대기
                    if (i + BATCH_SIZE < newBatches.length) {
                        await new Promise(resolve => setTimeout(resolve, 100));
                    }
                }
            }

            // 로컬 스토리지 업데이트
            localStorage.setItem('processedWaybills', JSON.stringify([...processedWaybills]));

            // 마지막 행 번호 업데이트
            lastKnownRow = currentRow - 1;

            console.log('구글 시트 백업 완료');

            // 일별 출고내역 데이터 처리
            await processDailySummary(uniqueOrders, accessToken);

        } catch (error) {
            console.error('구글 시트 백업 중 오류:', error);
            throw error;
        }
    }

    // 토큰 관리 함수
    async function getAccessToken() {
        const currentTime = Math.floor(Date.now() / 1000);

        // 캐시된 토큰이 있고 아직 유효한 경우
        if (cachedToken && tokenExpireTime > currentTime + 300) { // 5분 여유
            return cachedToken;
        }

        // 새 토큰 발급
        const header = {
            alg: 'RS256',
            typ: 'JWT'
        };

        const claim = {
            iss: GOOGLE_API_CONFIG.serviceAccount.client_email,
            scope: 'https://www.googleapis.com/auth/spreadsheets',
            aud: 'https://oauth2.googleapis.com/token',
            exp: currentTime + 3600,
            iat: currentTime
        };

        const base64Header = base64URLEncode(JSON.stringify(header));
        const base64Claim = base64URLEncode(JSON.stringify(claim));
        const signatureInput = `${base64Header}.${base64Claim}`;

        const privateKey = GOOGLE_API_CONFIG.serviceAccount.private_key.replace(/\\n/g, '\n');
        const key = KEYUTIL.getKey(privateKey);
        const sig = new KJUR.crypto.Signature({alg: 'SHA256withRSA'});
        sig.init(key);
        sig.updateString(signatureInput);
        const signatureBytes = sig.sign();
        const signature = base64URLEncode(hexToBase64(signatureBytes));
        const jwt = `${base64Header}.${base64Claim}.${signature}`;

        const tokenResponse = await fetch('https://oauth2.googleapis.com/token', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: new URLSearchParams({
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt
            })
        });

        if (!tokenResponse.ok) {
            throw new Error('토큰 발급 실패');
        }

        const tokenData = await tokenResponse.json();
        cachedToken = tokenData.access_token;
        tokenExpireTime = currentTime + 3600; // 1시간

        return cachedToken;
    }

    // 마지막 행 번호 조회 함수
    async function getLastRowNumber(accessToken) {
        const response = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_API_CONFIG.spreadsheetId}/values/백업!A:A`,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            }
        );

        if (!response.ok) {
            throw new Error('시트 데이터 조회 실패');
        }

        const data = await response.json();
        return data.values ? data.values.length : 1;
    }

    // 일별 출고내역 처리 함수
    async function processDailySummary(orders, accessToken) {
        const dailySummary = new Map();

        orders.flatMap(order => order.itemInfos).forEach(item => {
            const key = `${item.prdCode}|${item.prdName}`;
            if (dailySummary.has(key)) {
                dailySummary.get(key).quantity += parseInt(item.itemCnt || '0', 10);
            } else {
                dailySummary.set(key, {
                    code: item.prdCode,
                    name: item.prdName,
                    quantity: parseInt(item.itemCnt || '0', 10)
                });
            }
        });

        const summaryData = Array.from(dailySummary.values()).map(item =>
                                                                  ['', '', '', '', item.code, item.name, item.quantity]
                                                                 );

        if (summaryData.length === 0) return;

        const response = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_API_CONFIG.dailySpreadsheetId}/values/일별 출고내역!A2:G2:append?valueInputOption=USER_ENTERED`,
            {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    values: summaryData
                })
            }
        );

        if (!response.ok) {
            throw new Error('일별 출고내역 업데이트 실패');
        }
    }

    // base64URLEncode 함수 추가
    function base64URLEncode(str) {
        return btoa(str)
            .replace(/\+/g, '-')
            .replace(/\//g, '_')
            .replace(/=/g, '');
    }

    // 16진수 문자열을 Base64로 변환하는 함수
    function hexToBase64(hexstring) {
        const bytes = [];
        for (let i = 0; i < hexstring.length - 1; i += 2) {
            bytes.push(parseInt(hexstring.substr(i, 2), 16));
        }
        return String.fromCharCode.apply(String, bytes);
    }

    // 패턴 분석 버튼 클릭 핸들러
    function showPatternDialog() {
        // 기존 UI 요소들 모두 제거
        const existingDialog = document.querySelector('#pattern-dialog');
        if (existingDialog) existingDialog.remove();

        const existingFilterUI = document.getElementById('filter-ui');
        if (existingFilterUI) existingFilterUI.remove();

        const existingResultContainer = document.getElementById('result-container');
        if (existingResultContainer) existingResultContainer.remove();

        const dialog = document.createElement('div');
        dialog.id = 'pattern-dialog';  // ID 추가
        dialog.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            z-index: 10000;
            width: 80%;
            max-height: 80vh;
            overflow-y: auto;
        `;

        dialog.innerHTML = `
            <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                <h3>패턴 분석</h3>
                <button id="close-pattern" style="
                    padding: 5px 10px;
                    background: #f44336;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                ">닫기</button>
            </div>
            <div style="margin-bottom: 20px;">
                <input type="text" id="pattern-product-filter"
                    placeholder="상품코드/상품명 검색"
                    style="width: 200px; padding: 8px; margin-right: 10px;">
                <input type="number" id="pattern-min-count"
                    placeholder="최소 반복 횟수"
                    style="width: 150px; padding: 8px;">
            </div>
            <table style="width: 100%; border-collapse: collapse;">
                <thead>
                    <tr>
                        <th style="padding: 12px; border: 1px solid #ddd; background: #f5f5f5;">반복 횟수</th>
                        <th style="padding: 12px; border: 1px solid #ddd; background: #f5f5f5;">상품 구성</th>
                        <th style="padding: 12px; border: 1px solid #ddd; background: #f5f5f5;">주문번호</th>
                    </tr>
                </thead>
                <tbody id="pattern-table-body"></tbody>
            </table>
        `;

        document.body.appendChild(dialog);

        // 닫기 버튼 이벤트
        document.getElementById('close-pattern').addEventListener('click', () => {
            document.body.removeChild(dialog);
        });

        // 필터 이벤트
        document.getElementById('pattern-product-filter').addEventListener('input', updatePatternTable);
        document.getElementById('pattern-min-count').addEventListener('input', updatePatternTable);

        // 패턴 분석 실행
        processPattern();
    }

    // 패턴 분석 함수
    async function processPattern(orders) {
        try {
            console.log('입력된 주문 데이터:', orders);

            // 판매처 정보 로드
            await loadSellerInfo();

            // 패턴 맵 초기화
            patternMap = new Map();
            const trackingPatterns = new Map();

            // 주문 데이터 처리
            for (const order of orders) {
                console.log('처리중인 주문:', order);
                const items = new Map();

                order.itemInfos.forEach(item => {
                    const key = item.prdCode;
                    if (items.has(key)) {
                        // 기존 수량에 더하기
                        const existingItem = items.get(key);
                        existingItem.quantity += parseInt(item.itemCnt || '0', 10);
                    } else {
                        // 새로운 아이템 추가
                        items.set(key, {
                            quantity: parseInt(item.itemCnt || '0', 10),
                            productName: item.prdName
                        });
                    }
                });

                trackingPatterns.set(order.FWD_0006, {
                    items,
                    sellerCode: order.MBI,
                    waybills: new Set([order.FWD_0006])
                });
            }

            console.log('trackingPatterns:', trackingPatterns);

            // 패턴 분석
            trackingPatterns.forEach((patternData, trackingNo) => {
                if (patternData.items.size > 0) {
                    const sortedItems = Array.from(patternData.items.entries())
                        .sort((a, b) => a[0].localeCompare(b[0]))
                        .map(([barcode, itemData]) => [barcode, itemData.quantity, itemData.productName]);

                    const pattern = sortedItems.map(([barcode, qty]) => `${barcode}:${qty}`).join('|');

                    if (patternMap.has(pattern)) {
                        const existing = patternMap.get(pattern);
                        existing.count += 1;
                        existing.waybills.add(trackingNo);
                        if (!existing.sellerCodes.includes(patternData.sellerCode)) {
                            existing.sellerCodes.push(patternData.sellerCode);
                        }
                    } else {
                        patternMap.set(pattern, {
                            items: sortedItems,
                            count: 1,
                            sellerCodes: [patternData.sellerCode],
                            waybills: new Set([trackingNo]),
                            printStatus: 'not-printed'
                        });
                    }
                }
            });

            // 최소 반복수 필터 적용 및 짜투리 처리
            const minRepeatFilter = parseInt(document.getElementById('min-repeat-filter')?.value, 10) || 0;
            let totalRemnantCount = 0;
            const allRemnantItems = [];
            const allRemnantWaybills = new Set();

            window.sortedPatterns = Array.from(patternMap.entries())
                .sort((a, b) => b[1].count - a[1].count)
                .filter(([pattern, data]) => {
                    if (minRepeatFilter > 0 && data.count < minRepeatFilter) {
                        totalRemnantCount++;
                        data.items.forEach(item => allRemnantItems.push(item));
                        data.waybills.forEach(waybill => allRemnantWaybills.add(waybill));
                        return false;
                    }
                    return true;
                })
                .map(([pattern, data]) => {
                    const sellerInfo = sellerInfoMap.get(data.sellerCodes[0]);
                    const sellerName = sellerInfo ? sellerInfo.name.replace(/주식회사\s*/g, '') : data.sellerCodes[0];

                    return [pattern, {
                        ...data,
                        sellerName
                    }];
                });

            // 짜투리가 있는 경우 별도의 행으로 추가
            if (totalRemnantCount > 0) {
                window.sortedPatterns.push(['remnant', {
                    items: [],
                    count: totalRemnantCount,
                    sellerName: '짜투리',
                    waybills: allRemnantWaybills,
                    printStatus: 'not-printed',
                    isRemnant: true,
                    remnantItems: allRemnantItems
                }]);
            }

            // 재고 기반 주문 할당 처리
            window.sortedPatterns = await allocateOrdersByInventory(window.sortedPatterns);

            // 결과 테이블 생성
            const table = createResultTable();

            // 테이블 내용 업데이트
            updateTableContent(window.sortedPatterns);

            // 필터 이벤트 리스너 추가
            document.getElementById('seller-filter')?.addEventListener('change', () => applyFilters(window.sortedPatterns));
            document.getElementById('barcode-filter')?.addEventListener('input', () => applyFilters(window.sortedPatterns));
            document.getElementById('min-repeat-filter')?.addEventListener('input', () => applyFilters(window.sortedPatterns));
            document.getElementById('print-status-filter')?.addEventListener('change', () => applyFilters(window.sortedPatterns));

        } catch (error) {
            console.error('패턴 분석 중 오류:', error);
            throw error;
        }
    }

    // 테이블 내용 업데이트 함수
    function updateTableContent(patterns) {
        const tbody = document.getElementById('pattern-table-body');
        if (!tbody) return;

        // 판매처 필터 업데이트
        const sellerFilter = document.getElementById('seller-filter');
        if (sellerFilter) {
            const uniqueSellers = new Set();
            patterns.forEach(([_, data]) => {
                if (data.sellerName) {
                    uniqueSellers.add(data.sellerName);
                }
            });

            // 기존 옵션 제거
            while (sellerFilter.firstChild) {
                sellerFilter.removeChild(sellerFilter.firstChild);
            }

            // 전체 판매처 옵션 추가
            const allOption = document.createElement('option');
            allOption.value = '';
            allOption.textContent = '전체 판매처';
            sellerFilter.appendChild(allOption);

            // 판매처 옵션 추가
            Array.from(uniqueSellers)
                .sort((a, b) => a.localeCompare(b))
                .forEach(seller => {
                    const option = document.createElement('option');
                    option.value = seller;
                    option.textContent = seller;
                    sellerFilter.appendChild(option);
                });
        }

        tbody.innerHTML = patterns.map(([pattern, data]) => {
            const printStatus = data.printStatus || 'not-printed';
            const printTime = data.printTime || '';

            // 패턴 상세 내용 생성
            let itemsHtml = '';

            if (data.isOutOfStock) {
                // 재고 부족 패턴 표시
                itemsHtml = `
                    <div style="margin: 5px 0;">
                        <div style="color: #f44336; font-weight: bold; margin-bottom: 8px;">
                            ▼ 재고 부족으로 출고 불가한 주문 (${data.count}건)
                        </div>
                        <div style="margin: 5px 0; color: #f44336;">
                            해당 주문들은 재고가 부족하여 출고할 수 없습니다.
                        </div>
                    </div>`;
            } else if (data.isRemnant) {
                itemsHtml = `
                    <div style="margin: 5px 0;">
                        <div style="color: #666; font-weight: bold; margin-bottom: 8px;">
                            ▼ 짜투리 패턴 (${data.count}건)
                        </div>
                        ${data.remnantItems.map(([barcode, qty, productName]) =>
                            `<div style="margin: 5px 0;">
                                <span style="font-weight: bold;">${barcode}</span>
                                <span style="color: #666; margin-left: 8px;">${productName || ''}</span>
                                <span style="float: right;">${qty}개</span>
                            </div>`
                        ).join('')}
                    </div>`;
            } else {
                itemsHtml = data.items.map(([barcode, qty, productName]) =>
                    `<div style="margin: 5px 0;">
                        <span style="font-weight: bold;">${barcode}</span>
                        <span style="color: #666; margin-left: 8px;">${productName || ''}</span>
                        <span style="float: right;">${qty}개</span>
                    </div>`
                ).join('');
            }

            // 배경색 지정 (재고 부족인 경우 연한 빨간색)
            const rowStyle = data.isOutOfStock ? 'background-color: #ffebee;' : '';

            return `
                <tr style="${rowStyle}">
                    <td style="border: 1px solid #ddd; padding: 12px; text-align: center;">
                        ${data.sellerName}
                    </td>
                    <td style="border: 1px solid #ddd; padding: 12px;">
                        ${itemsHtml}
                    </td>
                    <td style="border: 1px solid #ddd; padding: 12px; text-align: center;">
                        ${data.count}회
                    </td>
                    <td style="border: 1px solid #ddd; padding: 12px; text-align: center;">
                        <button class="pattern-search-button"
                            data-pattern="${pattern}"
                            data-waybills="${Array.from(data.waybills).join(',')}"
                            style="padding: 8px 15px; background: #2196F3; color: white; border: none; border-radius: 4px; cursor: pointer;">
                            조회
                        </button>
                    </td>
                    <td style="border: 1px solid #ddd; padding: 12px; text-align: center;">
                        <button class="print-status-button"
                            data-pattern="${pattern}"
                            data-waybills="${Array.from(data.waybills).join(',')}"
                            style="padding: 8px 15px; background: ${data.isOutOfStock ? '#f44336' : (printStatus === 'printed' ? '#4CAF50' : '#FF9800')}; color: white; border: none; border-radius: 4px; cursor: pointer;">
                            ${data.isOutOfStock ? '출고 불가' : (printStatus === 'printed' ? '출력 완' : '출력 전')}
                        </button>
                        ${printTime ? `<div style="margin-top: 8px; font-size: 12px; color: #666;">${printTime}</div>` : ''}
                    </td>
                </tr>
            `;
        }).join('');

        // 버튼 이벤트 리스너 즉시 추가
        addButtonEventListeners();
    }

    // 결과 테이블 생성 함수 수정
    function createResultTable() {
        // 기존 UI 요소들 제거
        const existingElements = document.querySelectorAll('#filter-ui, #result-container, #pattern-dialog');
        existingElements.forEach(el => el.remove());

        // 결과 컨테이너 생성
        const resultContainer = document.createElement('div');
        resultContainer.id = 'result-container';
        resultContainer.style.cssText = `
            margin: 20px;
            padding: 20px;
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        `;

        // 필터 UI 생성
        const filterUI = document.createElement('div');
        filterUI.id = 'filter-ui';
        filterUI.style.cssText = `
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 10px;
        `;

        filterUI.innerHTML = `
            <select id="seller-filter" style="padding: 5px; flex: 1; border: 1px solid #ddd; border-radius: 4px;">
                <option value="">전체 판매처</option>
            </select>
            <input type="text" id="barcode-filter" placeholder="바코드 필터"
                style="padding: 5px; flex: 1; border: 1px solid #ddd; border-radius: 4px;">
            <input type="number" id="min-repeat-filter" placeholder="최소 반복수"
                style="padding: 5px; flex: 1; border: 1px solid #ddd; border-radius: 4px;">
            <select id="print-status-filter" style="padding: 5px; flex: 1; border: 1px solid #ddd; border-radius: 4px;">
                <option value="not-printed" selected>출력 전</option>
                <option value="printed">출력 완</option>
                <option value="all">전체</option>
            </select>
        `;

        // 테이블 생성
        const table = document.createElement('table');
        table.id = 'pattern-table';
        table.style.cssText = `
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        `;

        table.innerHTML = `
            <thead>
                <tr>
                    <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; text-align: center;">판매처</th>
                    <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; width: 45%;">패턴 상세</th>
                    <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; text-align: center; width: 10%;">반복수</th>
                    <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; text-align: center; width: 15%;">조회하기</th>
                    <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; text-align: center; width: 15%;">출력여부</th>
                </tr>
            </thead>
            <tbody id="pattern-table-body">
            </tbody>
        `;

        // 요소들을 컨테이너에 추가
        resultContainer.appendChild(filterUI);
        resultContainer.appendChild(table);
        document.body.appendChild(resultContainer);

        // 필터 이벤트 리스너 추가
        document.getElementById('seller-filter')?.addEventListener('change', () => applyFilters(window.sortedPatterns));
        document.getElementById('barcode-filter')?.addEventListener('input', () => applyFilters(window.sortedPatterns));
        document.getElementById('min-repeat-filter')?.addEventListener('input', () => applyFilters(window.sortedPatterns));
        document.getElementById('print-status-filter')?.addEventListener('change', () => applyFilters(window.sortedPatterns));

        // 초기 데이터 표시
        if (window.sortedPatterns) {
            updateTableContent(window.sortedPatterns);
        }

        // 초기 필터 적용
        setTimeout(() => applyFilters(window.sortedPatterns), 0);

        return table;
    }

    function applyFilters(sortedPatterns) {
        if (!sortedPatterns) return;

        const sellerFilter = document.getElementById('seller-filter')?.value || '';
        const barcodeFilter = document.getElementById('barcode-filter')?.value.toLowerCase() || '';
        const minRepeatFilter = parseInt(document.getElementById('min-repeat-filter')?.value, 10) || 0;
        const printStatusFilter = document.getElementById('print-status-filter')?.value || 'all';

        // 짜투리 데이터와 재고 부족 데이터를 저장할 변수
        let totalRemnantCount = 0;
        let remnantWaybills = new Set();
        let totalOutOfStockCount = 0;
        let outOfStockWaybills = new Set();
        const filteredPatterns = [];

        // 패턴 필터링 및 짜투리/재고부족 수집
        sortedPatterns.forEach(([pattern, data]) => {
            // 재고 부족 항목은 항상 표시하되, 판매처 필터는 적용
            if (data.isOutOfStock) {
                const matchesSeller = !sellerFilter || data.sellerName.toLowerCase().includes(sellerFilter.toLowerCase());
                if (matchesSeller) {
                    filteredPatterns.push([pattern, data]);
                } else {
                    // 필터에 맞지 않는 경우 카운트만 증가
                    totalOutOfStockCount += data.count;
                    data.waybills.forEach(waybill => outOfStockWaybills.add(waybill));
                }
                return;
            }

            const matchesSeller = !sellerFilter || data.sellerName.toLowerCase().includes(sellerFilter.toLowerCase());
            const matchesBarcode = !barcodeFilter ||
                (data.items && data.items.some(([barcode]) => barcode.toLowerCase().includes(barcodeFilter)));
            const printStatus = data.printStatus || 'not-printed';
            const matchesPrintStatus = printStatusFilter === 'all' || printStatusFilter === printStatus;

            if (matchesSeller && matchesBarcode && matchesPrintStatus) {
                if (data.isRemnant || data.count >= minRepeatFilter) {
                    // 정상 패턴 또는 이미 짜투리로 표시된 항목
                    filteredPatterns.push([pattern, data]);
                } else {
                    // 짜투리로 카운트
                    totalRemnantCount += data.count;
                    data.waybills.forEach(waybill => remnantWaybills.add(waybill));
                }
            }
        });

        // 필터링된 재고 부족 항목이 있는 경우 통합하여 추가
        if (totalOutOfStockCount > 0) {
            filteredPatterns.push(['out-of-stock-filtered', {
                sellerName: '재고 부족 (필터링됨)',
                items: [],
                count: totalOutOfStockCount,
                waybills: outOfStockWaybills,
                printStatus: 'not-printed',
                isOutOfStock: true
            }]);
        }

        // 짜투리가 있는 경우 마지막에 추가
        if (totalRemnantCount > 0) {
            filteredPatterns.push(['remnant', {
                sellerName: '짜투리',
                items: [['짜투리', totalRemnantCount, '']],  // 패턴상세도 단순하게 표시
                count: totalRemnantCount,
                waybills: remnantWaybills,
                printStatus: 'not-printed',
                isRemnant: true
            }]);
        }

        // 결과 정렬 (짜투리와 재고 부족은 맨 아래로)
        filteredPatterns.sort((a, b) => {
            if (a[1].isRemnant && b[1].isOutOfStock) return 1;  // 짜투리가 재고 부족 아래로
            if (a[1].isOutOfStock && b[1].isRemnant) return -1; // 재고 부족이 짜투리 위로
            if (a[1].isRemnant || a[1].isOutOfStock) return 1;  // 둘 다 특수 항목이면 재고 부족이 위로
            if (b[1].isRemnant || b[1].isOutOfStock) return -1;
            return b[1].count - a[1].count;
        });

        // 테이블 본문 업데이트
        const tbody = document.getElementById('pattern-table-body');
        if (tbody) {
            tbody.innerHTML = filteredPatterns.map(([pattern, data]) => {
                const printStatus = data.printStatus || 'not-printed';
                const printTime = data.printTime || '';
                const sellerDisplay = data.isRemnant
                    ? `${data.sellerName} (짜투리)`
                    : data.sellerName;

                return `
                    <tr>
                        <td style="border: 1px solid #ddd; padding: 12px; text-align: center; font-size: 14px;">
                            ${sellerDisplay}
                        </td>
                        <td style="border: 1px solid #ddd; padding: 12px; font-size: 14px;">
                            ${data.items.map(([barcode, qty, productName]) =>
                                `<div style="margin: 8px 0;">
                                    <span style="font-weight: bold; font-size: 14px;">${barcode}</span>
                                    <span style="color: #666; margin-left: 8px; font-size: 14px;">${productName || ''}</span>
                                    <span style="float: right;">${qty}개</span>
                                </div>`
                            ).join('')}
                        </td>
                        <td style="border: 1px solid #ddd; padding: 12px; text-align: center; font-size: 14px;">
                            ${data.count}회
                        </td>
                        <td style="border: 1px solid #ddd; padding: 12px; text-align: center;">
                            <button class="pattern-search-button"
                                data-pattern="${pattern}"
                                data-seller-code="${data.sellerCode}"
                                data-waybills="${Array.from(data.waybills).join(',')}"
                                style="
                                    padding: 8px 15px;
                                    background: #2196F3;
                                    color: white;
                                    border: none;
                                    border-radius: 4px;
                                    cursor: pointer;
                                    font-size: 14px;
                                ">조회</button>
                        </td>
                        <td style="border: 1px solid #ddd; padding: 12px; text-align: center;">
                            <button class="print-status-button"
                                data-pattern="${pattern}"
                                data-waybills="${Array.from(data.waybills).join(',')}"
                                style="
                                    padding: 8px 15px;
                                    background: ${printStatus === 'printed' ? '#4CAF50' : '#FF9800'};
                                    color: white;
                                    border: none;
                                    border-radius: 4px;
                                    cursor: pointer;
                                    font-size: 14px;
                                "
                            >${printStatus === 'printed' ? '출력 완' : '출력 전'}</button>
                            ${printTime ? `<div style="margin-top: 8px; font-size: 13px; color: #666;">${printTime}</div>` : ''}
                        </td>
                    </tr>
                `;
            }).join('');

            // 버튼 이벤트 리스너 다시 추가
            addButtonEventListeners();
        }
    }

    // 버튼 이벤트 리스너 함수 수정
    function addButtonEventListeners() {
        document.querySelectorAll('.pattern-search-button').forEach(button => {
            button.addEventListener('click', async function() {
                const waybillNumbers = this.dataset.waybills.split(',');
                const pattern = this.dataset.pattern;

                // 운송장번호 입력
                const textarea = document.querySelector('textarea[name="waybillList"]');
                if (textarea) {
                    textarea.value = Array.from(new Set(waybillNumbers)).join('\n');

                    // 조회 버튼 클릭
                    const searchBtn = document.getElementById('btn_tradecode');
                    if (searchBtn) {
                        searchBtn.click();

                        // 판매처 목록이 로드될 때까지 대기
                        await new Promise(resolve => setTimeout(resolve, 50));

                        // 판매처 선택 로직
                        if (pattern === 'remnant') {
                            // 짜투리인 경우 첫 번째 판매처 선택
                            const firstSellerBtn = document.querySelector('.user_list_add_btn');
                            if (firstSellerBtn) {
                                firstSellerBtn.click();
                            }
                        } else {
                            // 일반 패턴인 경우 해당 판매처 선택
                            const patternData = window.sortedPatterns.find(([p]) => p === pattern);
                            if (patternData) {
                                const sellerCode = patternData[1].sellerCodes[0];
                                const sellerBtn = document.querySelector(`button[data-mbi-code="${sellerCode}"]`);
                                if (sellerBtn) {
                                    sellerBtn.click();
                                }
                            }
                        }

                        // 주문수집 버튼 클릭
                        setTimeout(() => {
                            const collectBtn = document.getElementById('form_search_submit');
                            if (collectBtn) {
                                collectBtn.click();
                            }
                        }, 50);
                    }
                }
            });
        });

        document.querySelectorAll('.print-status-button').forEach(button => {
            button.addEventListener('click', function() {
                const pattern = this.dataset.pattern;
                const currentStatus = this.textContent.trim() === '출력 전' ? 'not-printed' : 'printed';
                const newStatus = currentStatus === 'printed' ? 'not-printed' : 'printed';
                const now = new Date();
                const timeString = now.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });

                if (pattern === 'remnant') {
                    // 짜투리 패턴 처리 수정
                    const remnantWaybills = new Set(this.dataset.waybills.split(','));

                    // 상태 업데이트
                    window.sortedPatterns = window.sortedPatterns.map(([p, data]) => {
                        if (p === 'remnant' || Array.from(data.waybills).some(w => remnantWaybills.has(w))) {
                            return [p, {
                                ...data,
                                printStatus: newStatus,
                                printTime: newStatus === 'printed' ? timeString : ''
                            }];
                        }
                        return [p, data];
                    });

                    // 즉시 UI 업데이트
                    updateButtonStatus(this, newStatus, timeString);

                    // 현재 필터와 새로운 상태가 일치하지 않으면 행 제거
                    const currentFilter = document.getElementById('print-status-filter').value;
                    if ((currentFilter === 'not-printed' && newStatus === 'printed') ||
                        (currentFilter === 'printed' && newStatus === 'not-printed')) {
                        const row = this.closest('tr');
                        if (row) {
                            row.remove();
                        }
                    }
                } else {
                    // 일반 패턴 처리
                    window.sortedPatterns = window.sortedPatterns.map(([p, data]) => {
                        if (p === pattern) {
                            return [p, {
                                ...data,
                                printStatus: newStatus,
                                printTime: newStatus === 'printed' ? timeString : ''
                            }];
                        }
                        return [p, data];
                    });

                    // 즉시 UI 업데이트
                    updateButtonStatus(this, newStatus, timeString);

                    // 현재 필터와 새로운 상태가 일치하지 않으면 행 제거
                    const currentFilter = document.getElementById('print-status-filter').value;
                    if ((currentFilter === 'not-printed' && newStatus === 'printed') ||
                        (currentFilter === 'printed' && newStatus === 'not-printed')) {
                        const row = this.closest('tr');
                        if (row) {
                            row.remove();
                        }
                    }
                }
            });
        });
    }

    // 버튼 상태 업데이트 헬퍼 함수 추가
    function updateButtonStatus(button, status, timeString) {
        button.textContent = status === 'printed' ? '출력 완' : '출력 전';
        button.style.background = status === 'printed' ? '#4CAF50' : '#FF9800';

        let timeDiv = button.parentElement.querySelector('div');
        if (!timeDiv) {
            timeDiv = document.createElement('div');
            timeDiv.style.cssText = 'margin-top: 8px; font-size: 13px; color: #666;';
            button.parentElement.appendChild(timeDiv);
        }

        if (status === 'printed') {
            timeDiv.textContent = timeString;
            timeDiv.style.display = 'block';
        } else {
            timeDiv.style.display = 'none';
        }
    }

    // UI 생성 함수 분리
    function createPatternUI() {
        // 이전 UI 제거
        const oldFilterUI = document.getElementById('filter-ui');
        if (oldFilterUI) oldFilterUI.remove();
        const oldResultTable = document.getElementById('result-table');
        if (oldResultTable) oldResultTable.remove();

        // 필터링 UI 생성
        const filterUI = document.createElement('div');
        filterUI.id = 'filter-ui';
        filterUI.style.cssText = `
            width: 80%;
            margin: 20px auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
        `;
        filterUI.innerHTML = `
            <select id="seller-filter" style="padding: 5px; width: 20%;">
                <option value="">전체 판매처</option>
            </select>
            <input type="text" id="barcode-filter" placeholder="바코드 필터" style="padding: 5px; width: 20%;">
            <input type="number" id="min-repeat-filter" placeholder="최소 반복수" style="padding: 5px; width: 20%;">
            <select id="print-status-filter" style="padding: 5px; width: 20%;">
                <option value="not-printed" selected>출력 전</option>
                <option value="printed">출력 완</option>
                <option value="all">전체</option>
            </select>
        `;

        // 결과 테이블 생성
        const resultTable = document.createElement('div');
        resultTable.id = 'result-table';
        resultTable.style.cssText = `
            width: 80%;
            margin: 20px auto;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        `;

        resultTable.innerHTML = `
            <h3 style="margin-bottom: 15px; font-size: 18px;">패턴 분석 결과</h3>
            <table style="width: 100%; border-collapse: collapse;">
                <thead>
                    <tr>
                        <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; text-align: center; width: 15%;">판매처</th>
                        <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; width: 45%;">패턴 상세</th>
                        <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; text-align: center; width: 10%;">반복수</th>
                        <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; text-align: center; width: 15%;">조회하기</th>
                        <th style="border: 1px solid #ddd; padding: 12px; background: #f5f5f5; text-align: center; width: 15%;">출력여부</th>
                    </tr>
                </thead>
                <tbody id="pattern-table-body">
                </tbody>
            </table>
        `;

        // UI 요소들을 페이지에 추가
        document.body.appendChild(filterUI);
        document.body.appendChild(resultTable);

        // 판매처 필터 옵션 추가
        const uniqueSellers = new Set();
        window.sortedPatterns.forEach(([_, data]) => {
            uniqueSellers.add(data.sellerName);
        });

        const sellerFilter = document.getElementById('seller-filter');
        Array.from(uniqueSellers)
            .sort((a, b) => a.localeCompare(b))
            .forEach(seller => {
                const option = document.createElement('option');
                option.value = seller;
                option.textContent = seller;
                sellerFilter.appendChild(option);
            });

        // 필터 이벤트 리스너 추가
        document.getElementById('seller-filter')?.addEventListener('change', () => applyFilters(window.sortedPatterns));
        document.getElementById('barcode-filter')?.addEventListener('input', () => applyFilters(window.sortedPatterns));
        document.getElementById('min-repeat-filter')?.addEventListener('input', () => applyFilters(window.sortedPatterns));
        document.getElementById('print-status-filter')?.addEventListener('change', () => applyFilters(window.sortedPatterns));
    }

    // 패턴 분석 버튼 추가
    function addPatternButton() {
        const button = document.createElement('button');
        button.textContent = '패턴 분석';
        button.style.cssText = `
            position: fixed;
            top: 7px;
            right: 360px;
            z-index: 9999;
            padding: 8px 16px;
            background-color: #FF9800;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        `;

        button.addEventListener('click', showPatternDialog);
        document.body.appendChild(button);
    }

    // 초기화 시 버튼 추가
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initialize);
    } else {
        initialize();
    }

    // 초기화 버튼 추가 함수
    function addResetButton() {
        const button = document.createElement('button');
        button.textContent = '백업 초기화';
        button.style.cssText = `
        position: fixed;
        top: 7px;
        right: 470px;
        z-index: 9999;
        padding: 8px 16px;
        background-color: #f44336;
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 13px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    `;
        button.addEventListener('click', resetProcessedOrders);
        document.body.appendChild(button);
    }

    // 초기화 함수
    function resetProcessedOrders() {
        if (confirm('오늘 백업된 모든 주문 기록을 초기화하시겠습니까?\n초기화하면 동일한 주문을 다시 백업할 수 있습니다.')) {
            processedWaybills.clear();
            localStorage.setItem('processedWaybills', JSON.stringify([]));
            alert('백업 기록이 초기화되었습니다.\n이제 주문을 다시 수집할 수 있습니다.');
        }
    }

    function initialize() {
        // 저장된 운송장 데이터 로드
        loadProcessedWaybills();
        // 출고건 확인 버튼 추가
        addCheckOrderButton();
        // 초기화 버튼 추가
        addResetButton();
    }

    // 주문 처리 다이얼로그 이벤트 핸들러 설정
    function setupOrderDialogHandlers(dialog) {
        // 전체 선택 체크박스 이벤트 리스너
        const selectAllCheckbox = document.getElementById('select-all-sellers');
        if (selectAllCheckbox) {
            selectAllCheckbox.addEventListener('change', (e) => {
                const checkboxes = document.querySelectorAll('.seller-checkbox');
                checkboxes.forEach(cb => cb.checked = e.target.checked);
            });
        }

        // 전체 출고 버튼
        const processAllBtn = document.getElementById('process-all-orders');
        if (processAllBtn) {
            processAllBtn.addEventListener('click', async () => {
                try {
                    await processOrders(false);
                    if (dialog && dialog.parentNode) {
                        dialog.parentNode.removeChild(dialog);
                    }
                } catch (error) {
                    console.error('처리 중 오류:', error);
                }
            });
        }

        // 선택 출고 버튼
        const processSelectedBtn = document.getElementById('process-selected-orders');
        if (processSelectedBtn) {
            processSelectedBtn.addEventListener('click', async () => {
                try {
                    await processOrders(true);
                    if (dialog && dialog.parentNode) {
                        dialog.parentNode.removeChild(dialog);
                    }
                } catch (error) {
                    console.error('처리 중 오류:', error);
                }
            });
        }

        // 취소 버튼
        const closeBtn = document.getElementById('close-order-count');
        if (closeBtn) {
            closeBtn.addEventListener('click', () => {
                if (dialog && dialog.parentNode) {
                    dialog.parentNode.removeChild(dialog);
                }
            });
        }
    }

    // 주문 처리 함수 수정
    async function processOrders(selectedOnly) {
        // 기존 UI 요소들 모두 제거
        const existingElements = document.querySelectorAll('#filter-ui, #result-container, #pattern-dialog');
        existingElements.forEach(el => el.remove());

        const selectedSellers = selectedOnly ?
            Array.from(document.querySelectorAll('.seller-checkbox:checked')).map(cb => cb.dataset.mbiCode) :
            Array.from(orderDataCache.keys());

        if (selectedSellers.length === 0) {
            alert('처리할 주문이 없습니다.');
            return;
        }

        const ordersToProcess = [];
        selectedSellers.forEach(mbiCode => {
            const data = orderDataCache.get(mbiCode);
            if (data) {
                ordersToProcess.push(...data.orders);
            }
        });

        try {
            // 패턴 분석 실행
            await processPattern(ordersToProcess);

            // 테이블 생성 및 데이터 표시
            createResultTable();

            // 구글 시트 백업 - 출고 가능한 주문만 백업
            const ordersToBackup = [];

            for (const [pattern, data] of window.sortedPatterns) {
                // 재고 부족이나 짜투리가 아닌 정상 패턴만 백업 대상에 포함
                if (!data.isOutOfStock && !data.isRemnant) {
                    for (const waybill of data.waybills) {
                        const matchingOrder = ordersToProcess.find(order => order.FWD_0006 === waybill);
                        if (matchingOrder) {
                            ordersToBackup.push(matchingOrder);
                        }
                    }
                }
            }

            if (ordersToBackup.length > 0) {
                try {
                    await backupToGoogleSheets(ordersToBackup);
                    console.log(`출고 가능한 ${ordersToBackup.length}건의 주문이 백업되었습니다.`);
                    alert(`출고 가능한 ${ordersToBackup.length}건의 주문이 백업되었습니다.`);
                } catch (backupError) {
                    console.error('구글 시트 백업 중 오류:', backupError);
                    alert('주문 백업 중 오류가 발생했습니다. 자세한 내용은 콘솔을 확인하세요.');
                }
            } else {
                console.log('출고 가능한 주문이 없어 백업을 진행하지 않았습니다.');
                alert('출고 가능한 주문이 없어 백업을 진행하지 않았습니다.');
            }

        } catch (error) {
            console.error('주문 처리 중 오류:', error);
            alert('주문 처리 중 오류가 발생했습니다.');
        }
    }

    // 판매처 정보 로드 함수
    async function loadSellerInfo() {
        try {
            // 검색할 페이지 범위 지정 (여기서 원하는 페이지 번호를 배열로 지정)
            const pagesToLoad = [1, 2];  // 예: [1]로 변경하면 1페이지만 검색

            for (const currentPage of pagesToLoad) {
                const pageParam = currentPage === 1 ? 1 : currentPage.toString(); // 1페이지는 숫자로, 나머지는 문자열로

                const response = await fetch('https://envycube.com/web2/ajax/proc/loadOrder/get_member_list.php', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                    },
                    body: new URLSearchParams({
                        action: 'get_member_list',
                        page: pageParam,
                        search_txt: ''
                    })
                });

                const data = await response.json();
                if (data.MSG === 'OK') {
                    data.DAT.res_result.forEach(member => {
                        sellerInfoMap.set(member.MBI_CODE, {
                            name: member.MBI_0032,
                            code: member.MBI_CODE,
                            ceo: member.MBI_0033
                        });
                    });

                    // API 호출 간 짧은 딜레이 추가
                    await new Promise(resolve => setTimeout(resolve, 100));
                }
            }

            console.log(`총 ${pagesToLoad.length}페이지의 판매처 정보가 로드되었습니다.`);
        } catch (error) {
            console.error('판매처 정보 로드 실패:', error);
        }
    }

    // 구글 시트에서 재고 데이터를 가져오는 함수
    async function fetchInventoryData() {
        try {
            console.log('재고 데이터를 구글 시트에서 가져오는 중...');

            // 액세스 토큰 가져오기
            const accessToken = await getAccessToken();

            // 재고 현황 시트에서 데이터 가져오기
            const response = await fetch(
                `https://sheets.googleapis.com/v4/spreadsheets/1eHVar4VJD-jpoHNuLr9SVncEght6Z-fLDIE7WkVbyOY/values/재고현황!B:G`,
                {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`
                    }
                }
            );

            const data = await response.json();
            const inventoryRows = data.values || [];

            // 바코드(B열)와 재고량(G열) 매핑
            const inventoryMap = new Map();

            // 첫 행은 헤더이므로 건너뛰기
            for (let i = 1; i < inventoryRows.length; i++) {
                const row = inventoryRows[i];
                if (row && row.length >= 6) {  // B열(인덱스 0)과 G열(인덱스 5)이 존재하는지 확인
                    const barcode = row[0];
                    const quantity = parseInt(row[5], 10) || 0;
                    if (barcode) {
                        // 이미 존재하는 바코드면 수량 합산, 아니면 새로 추가
                        const currentQty = inventoryMap.get(barcode) || 0;
                        inventoryMap.set(barcode, currentQty + quantity);
                    }
                }
            }

            // 디버깅을 위해 각 바코드별 재고량 로그 추가
            for (const [barcode, qty] of inventoryMap.entries()) {
                if (qty > 100) {  // 수량이 많은 아이템만 로그 출력 (필요에 따라 조정)
                    console.log(`바코드 ${barcode}의 총 재고량: ${qty}`);
                }
            }

            console.log(`구글 시트에서 ${inventoryMap.size}개의 재고 데이터를 불러왔습니다.`);
            return inventoryMap;
        } catch (error) {
            console.error('재고 데이터 불러오기 실패:', error);
            alert('재고 데이터를 불러오는 중 오류가 발생했습니다.');
            return new Map();
        }
    }

    // 패턴별로 재고 확인 및 주문 할당
    async function allocateOrdersByInventory(patterns) {
        try {
            // 재고 데이터 가져오기
            const inventoryMap = await fetchInventoryData();
            if (inventoryMap.size === 0) {
                alert('재고 데이터를 불러올 수 없습니다. 재고 확인 없이 처리합니다.');
                return patterns;
            }

            // 작업용 재고 맵 (할당 진행 중 재고를 차감하기 위함)
            const workingInventory = new Map(inventoryMap);

            // 디버그: 패턴 내 주요 품목 재고 확인
            for (const [pattern, data] of patterns) {
                if (!data.isRemnant) {
                    console.log(`패턴 ${pattern} 필요 품목:`);
                    for (const [barcode, qty] of data.items) {
                        const availableQty = workingInventory.get(barcode) || 0;
                        console.log(`- 품목 ${barcode}: 주문당 필요량 ${qty}, 총 재고 ${availableQty}`);
                    }
                }
            }

            // 할당 결과를 저장할 변수들
            const allocatedPatterns = [];
            const unallocatedOrders = new Set();
            const allocationResults = {
                allocated: 0,
                unallocated: 0,
                patterns: []
            };

            // 패턴 내 주문 개수에 따라 내림차순 정렬 (패턴 우선순위 적용)
            const sortedPatterns = Array.from(patterns).sort((a, b) => {
                // 짜투리 패턴은 항상 마지막으로
                if (a[1].isRemnant) return 1;
                if (b[1].isRemnant) return -1;

                // 주문 건수가 많은 패턴을 우선 처리
                return b[1].waybills.size - a[1].waybills.size;
            });

            console.log('패턴 우선순위 정렬 결과:');
            sortedPatterns.forEach(([pattern, data]) => {
                if (!data.isRemnant) {
                    console.log(`패턴: ${pattern}, 주문건수: ${data.waybills.size}`);
                }
            });

            // 패턴을 순서대로 처리
            for (const [pattern, data] of sortedPatterns) {
                // 짜투리 패턴은 그대로 유지
                if (data.isRemnant) {
                    allocatedPatterns.push([pattern, data]);
                    continue;
                }

                const allocatedWaybills = new Set();
                const unallocatedWaybills = new Set();

                console.log(`패턴 ${pattern} 처리 중 (주문 ${data.waybills.size}건)`);

                // 패턴 내 각 운송장별로 재고 확인 및 할당
                for (const waybill of data.waybills) {
                    let canAllocate = true;
                    const requiredItems = new Map();

                    // 패턴의 각 아이템에 대해 재고 확인
                    for (const [barcode, qty] of data.items) {
                        if (!requiredItems.has(barcode)) {
                            requiredItems.set(barcode, 0);
                        }
                        requiredItems.set(barcode, requiredItems.get(barcode) + qty);
                    }

                    // 모든 아이템에 대해 재고가 충분한지 확인
                    for (const [barcode, requiredQty] of requiredItems) {
                        const inventoryQty = workingInventory.get(barcode) || 0;
                        if (inventoryQty < requiredQty) {
                            canAllocate = false;
                            console.log(`운송장 ${waybill}: 아이템 ${barcode} 재고 부족 (필요: ${requiredQty}, 가용: ${inventoryQty})`);
                            break;
                        }
                    }

                    if (canAllocate) {
                        // 재고 차감 및 할당 성공 처리
                        for (const [barcode, requiredQty] of requiredItems) {
                            const newQty = workingInventory.get(barcode) - requiredQty;
                            workingInventory.set(barcode, newQty);
                            // 재고 변동 로깅 (필요 시 주석 해제)
                            // console.log(`아이템 ${barcode} 재고 차감: ${workingInventory.get(barcode) + requiredQty} -> ${newQty}`);
                        }
                        allocatedWaybills.add(waybill);
                        allocationResults.allocated++;
                    } else {
                        unallocatedWaybills.add(waybill);
                        unallocatedOrders.add(waybill);
                        allocationResults.unallocated++;
                    }
                }

                // 할당된 주문이 있는 경우 패턴 추가
                if (allocatedWaybills.size > 0) {
                    const newPatternData = {
                        ...data,
                        waybills: allocatedWaybills,
                        count: allocatedWaybills.size,
                    };
                    allocatedPatterns.push([pattern, newPatternData]);

                    // 결과 기록
                    allocationResults.patterns.push({
                        pattern,
                        allocated: allocatedWaybills.size,
                        unallocated: unallocatedWaybills.size
                    });

                    console.log(`패턴 ${pattern}: 할당 성공 ${allocatedWaybills.size}건, 할당 실패 ${unallocatedWaybills.size}건`);
                }

                // 각 패턴 처리 후 주요 품목의 남은 재고 확인
                for (const [barcode, _] of data.items) {
                    const remainingQty = workingInventory.get(barcode) || 0;
                    console.log(`패턴 ${pattern} 처리 후 품목 ${barcode}의 남은 재고: ${remainingQty}`);
                }
            }

            // 할당되지 않은 주문이 있는 경우 '재고 부족' 패턴으로 추가
            if (unallocatedOrders.size > 0) {
                allocatedPatterns.push(['out-of-stock', {
                    items: [],
                    count: unallocatedOrders.size,
                    sellerName: '재고 부족',
                    waybills: unallocatedOrders,
                    printStatus: 'not-printed',
                    isOutOfStock: true
                }]);
            }

            // 할당 결과 로그 출력
            console.log('재고 할당 결과:', allocationResults);
            alert(`재고 할당 결과:\n- 할당 성공: ${allocationResults.allocated}건\n- 재고 부족: ${allocationResults.unallocated}건`);

            return allocatedPatterns;
        } catch (error) {
            console.error('재고 기반 할당 중 오류:', error);
            alert('재고 기반 할당 처리 중 오류가 발생했습니다.');
            return patterns;
        }
    }
})();