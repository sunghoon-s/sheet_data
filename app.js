// 분석 버튼에 이벤트 리스너 추가
analyzeBtn.addEventListener('click', openFileDialog);
downloadBtn.addEventListener('click', downloadCSV);

const analyzeBtn = document.getElementById('analyze-btn');
const downloadBtn = document.getElementById('download-btn');

analyzeBtn.addEventListener('click', fetchAndParseCSV);
downloadBtn.addEventListener('click', downloadCSV);

// 성능 최적화 변수들
let currentData = [];
let displayedRows = 0;
const BATCH_SIZE = 500; // 한 번에 표시할 행 수
const RENDER_DELAY = 10; // 렌더링 간 딜레이 (ms)


function fetchAndParseCSV() {
    // 다운로드 버튼 숨기기
    downloadBtn.style.display = 'none';
    showProgress(0, 'CSV 파일을 불러오는 중...');
    analyzeBtn.disabled = true;
    analyzeBtn.textContent = '처리 중...';

    fetch('data_erp.csv')
        .then(response => {
            if (!response.ok) throw new Error('CSV 파일을 찾을 수 없습니다.');
            return response.text();
        })
        .then(csvText => {
            showProgress(20, 'CSV 데이터를 파싱하는 중...');
            setTimeout(() => {
                parseCSVData(csvText);
            }, 100);
        })
        .catch(error => {
            handleError(error);
        });
}

function parseCSVData(csvText) {
    // CSV 파싱 (쉼표 구분, 첫 행 헤더)
    const rows = csvText.trim().split(/\r?\n/).map(row => row.split(','));
    if (rows.length < 2) {
        handleError(new Error('CSV 데이터가 비어있거나 잘못되었습니다.'));
        return;
    }
    // 빈 헤더(공백) 컬럼 제거
    const headers = rows[0].map(h => h.trim());
    const validIndexes = headers.map((h, i) => h ? i : -1).filter(i => i !== -1);
    const filteredHeaders = validIndexes.map(i => headers[i]);
    const data = rows.slice(1).map(row => {
        const obj = {};
        validIndexes.forEach((colIdx, i) => {
            obj[filteredHeaders[i]] = (row[colIdx] || '').trim();
        });
        return obj;
    });
    // "규격"이 없는 행 제거
    const filteredData = data.filter(row => row['규격'] && row['규격'].trim() !== '');
    showProgress(60, `${filteredData.length}개 행을 처리하는 중...`);
    setTimeout(() => {
        processDataAsync(filteredData);
    }, 100);
}

function handleError(error) {
    hideProgress();
    resetButton();
    console.error('데이터 처리 오류:', error);
    alert('파일 처리 중 오류가 발생했습니다: ' + error.message);
}

async function processDataAsync(jsonData) {
    try {
        showProgress(70, '빈 열을 제거하는 중...');
        await delay(30);
        
        // 빈 헤더 열 제거
        const cleanedData = removeEmptyColumns(jsonData);
        console.log('빈 열 제거된 데이터:', cleanedData);

        showProgress(80, '데이터를 필터링하는 중...');
        await delay(30);
        
        // "규격" 헤더가 없거나 빈 행 필터링
        const filteredData = filterDataBySpec(cleanedData);
        console.log('필터링된 데이터:', filteredData);

        showProgress(90, '날짜와 시간을 포맷팅하는 중...');
        await delay(30);
        
        // 배치 단위로 날짜 및 시간 포맷팅
        const formattedData = await formatDataInBatches(filteredData);
        console.log('포맷팅된 데이터:', formattedData);

        showProgress(95, '테이블을 생성하는 중...');
        await delay(30);
        
        // 전역 변수에 데이터 저장
        window.excelData = formattedData;
        currentData = formattedData;
        displayedRows = 0;

        // 가상 스크롤링으로 테이블 표시
        await displayDataWithVirtualScrolling(formattedData);
        
        showProgress(100, `완료! ${formattedData.length}개의 데이터를 성공적으로 처리했습니다.`);
        
        // 다운로드 버튼 활성화
        downloadBtn.style.display = 'inline-block';
        
        setTimeout(() => {
            hideProgress();
            resetButton();
        }, 2000);
        
    } catch (error) {
        handleError(error);
    }
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function showProgress(percentage, message) {
    const progressContainer = document.getElementById('progress-container');
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    
    progressContainer.style.display = 'block';
    progressFill.style.width = percentage + '%';
    progressText.textContent = message;
}

function hideProgress() {
    const progressContainer = document.getElementById('progress-container');
    progressContainer.style.display = 'none';
}

function resetButton() {
    const analyzeBtn = document.getElementById('analyze-btn');
    analyzeBtn.disabled = false;
    analyzeBtn.textContent = 'CSV파일열기';
}

// 배치 단위로 데이터 포맷팅
async function formatDataInBatches(data) {
    if (data.length === 0) return data;
    
    const hasProductionDate = data.length > 0 && '생산일자' in data[0];
    const hasStartTime = data.length > 0 && '시작시간' in data[0];
    const hasEndTime = data.length > 0 && '종료시간' in data[0];
    
    const result = [];
    const batchSize = 1000; // 1000개씩 처리
    
    for (let i = 0; i < data.length; i += batchSize) {
        const batch = data.slice(i, i + batchSize);
        
        const formattedBatch = batch.map(row => {
            const newRow = { ...row };
            
            if (hasProductionDate && newRow['생산일자']) {
                newRow['생산일자'] = formatDate(newRow['생산일자']);
            }
            if (hasStartTime && newRow['시작시간']) {
                newRow['시작시간'] = formatTime(newRow['시작시간']);
            }
            if (hasEndTime && newRow['종료시간']) {
                newRow['종료시간'] = formatTime(newRow['종료시간']);
            }
            
            return newRow;
        });
        
        result.push(...formattedBatch);
        
        // 배치 처리 후 잠시 대기 (UI 블로킹 방지)
        if (i + batchSize < data.length) {
            await delay(5);
        }
    }
    
    return result;
}

// 가상 스크롤링으로 테이블 표시
async function displayDataWithVirtualScrolling(data) {
    const tableContainer = document.getElementById('data-table');
    
    if (data.length === 0) {
        tableContainer.innerHTML = '<p>표시할 데이터가 없습니다.</p>';
        return;
    }

    // 초기 컨테이너 설정
    let html = `<div class="data-info">총 ${data.length}개의 유효한 데이터가 있습니다.</div>`;
    html += '<div class="table-container">';
    html += '<table class="data-table">';
    
    // 헤더 생성
    const headers = Object.keys(data[0]);
    html += '<thead><tr>';
    headers.forEach(header => {
        html += `<th>${header}</th>`;
    });
    html += '</tr></thead>';
    html += '<tbody id="table-body"></tbody>';
    html += '</table>';
    
    // 더보기 버튼 추가
    if (data.length > BATCH_SIZE) {
        html += `<div class="load-more-container">
            <button id="load-more-btn" class="btn btn-secondary">
                더 보기 (${Math.min(BATCH_SIZE, data.length - displayedRows)}개 더 로드)
            </button>
        </div>`;
    }
    
    html += '</div>';
    tableContainer.innerHTML = html;
    
    // 초기 배치 로드
    await loadMoreRows();
    
    // 더보기 버튼 이벤트 연결
    const loadMoreBtn = document.getElementById('load-more-btn');
    if (loadMoreBtn) {
        loadMoreBtn.addEventListener('click', loadMoreRows);
    }
}

// 추가 행 로드
async function loadMoreRows() {
    const tableBody = document.getElementById('table-body');
    const loadMoreBtn = document.getElementById('load-more-btn');
    
    if (!tableBody || displayedRows >= currentData.length) return;
    
    // 로딩 상태 표시
    if (loadMoreBtn) {
        loadMoreBtn.textContent = '로딩 중...';
        loadMoreBtn.disabled = true;
    }
    
    const headers = Object.keys(currentData[0]);
    const endIndex = Math.min(displayedRows + BATCH_SIZE, currentData.length);
    const batch = currentData.slice(displayedRows, endIndex);
    
    // 배치 단위로 DOM 업데이트
    const fragment = document.createDocumentFragment();
    
    batch.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.setAttribute('data-row-index', displayedRows + index);
        
        headers.forEach(header => {
            const td = document.createElement('td');
            td.textContent = row[header] || '';
            tr.appendChild(td);
        });
        
        fragment.appendChild(tr);
    });
    
    tableBody.appendChild(fragment);
    displayedRows = endIndex;
    
    // 더보기 버튼 업데이트
    if (loadMoreBtn) {
        if (displayedRows >= currentData.length) {
            loadMoreBtn.style.display = 'none';
        } else {
            loadMoreBtn.textContent = `더 보기 (${Math.min(BATCH_SIZE, currentData.length - displayedRows)}개 더 로드)`;
            loadMoreBtn.disabled = false;
        }
    }
    
    // 브라우저가 렌더링할 시간 제공
    await delay(RENDER_DELAY);
}

function removeEmptyColumns(data) {
    if (data.length === 0) return data;
    
    // 빈 헤더 또는 공백만 있는 헤더 찾기 (최적화: Set 사용)
    const headers = Object.keys(data[0]);
    const validHeaders = headers.filter(header => {
        return header && 
               header.trim() !== '' && 
               header !== 'undefined' &&
               header !== '__EMPTY' &&
               !header.startsWith('__EMPTY_');
    });
    
    console.log(`총 ${headers.length}개 열 중 ${validHeaders.length}개의 유효한 열이 있습니다.`);
    console.log('제거된 열:', headers.filter(h => !validHeaders.includes(h)));
    
    // 유효한 헤더만 포함된 새로운 데이터 생성 (최적화: map 사용)
    const cleanedData = data.map(row => {
        const newRow = {};
        validHeaders.forEach(header => {
            newRow[header] = row[header];
        });
        return newRow;
    });
    
    return cleanedData;
}

function formatDateAndTime(data) {
    if (data.length === 0) return data;
    
    // 최적화: 한 번만 키 검사
    const hasProductionDate = data.length > 0 && '생산일자' in data[0];
    const hasStartTime = data.length > 0 && '시작시간' in data[0];
    const hasEndTime = data.length > 0 && '종료시간' in data[0];
    
    return data.map(row => {
        const newRow = { ...row };
        
        // 필요한 필드만 처리 (성능 최적화)
        if (hasProductionDate && newRow['생산일자']) {
            newRow['생산일자'] = formatDate(newRow['생산일자']);
        }
        if (hasStartTime && newRow['시작시간']) {
            newRow['시작시간'] = formatTime(newRow['시작시간']);
        }
        if (hasEndTime && newRow['종료시간']) {
            newRow['종료시간'] = formatTime(newRow['종료시간']);
        }
        
        return newRow;
    });
}

function formatDate(dateValue) {
    try {
        // 이미 YYYY-MM-DD 형태인 경우
        if (typeof dateValue === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateValue)) {
            return dateValue;
        }
        
        let date;
        
        // 엑셀 날짜 시리얼 번호인 경우
        if (typeof dateValue === 'number') {
            // 엑셀 날짜를 JavaScript Date로 변환
            date = new Date((dateValue - 25569) * 86400 * 1000);
        } else {
            // 문자열인 경우 Date 객체로 변환
            date = new Date(dateValue);
        }
        
        // 유효한 날짜인지 확인
        if (isNaN(date.getTime())) {
            return dateValue; // 변환 실패시 원본 값 반환
        }
        
        // YYYY-MM-DD 형태로 포맷팅
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        
        return `${year}-${month}-${day}`;
    } catch (error) {
        console.warn('날짜 포맷팅 오류:', dateValue, error);
        return dateValue; // 오류 발생시 원본 값 반환
    }
}

function formatTime(timeValue) {
    try {
        // 이미 HH:MM 형태인 경우
        if (typeof timeValue === 'string' && /^\d{1,2}:\d{2}$/.test(timeValue)) {
            return timeValue;
        }
        
        let hours, minutes;
        
        // 엑셀 시간 시리얼 번호인 경우 (0~1 사이의 소수)
        if (typeof timeValue === 'number' && timeValue >= 0 && timeValue <= 1) {
            const totalMinutes = Math.round(timeValue * 24 * 60);
            hours = Math.floor(totalMinutes / 60);
            minutes = totalMinutes % 60;
        }
        // 시간이 숫자로 저장된 경우 (예: 1720 -> 17:20)
        else if (typeof timeValue === 'number' && timeValue > 1) {
            const timeStr = String(timeValue).padStart(4, '0');
            hours = parseInt(timeStr.substring(0, 2));
            minutes = parseInt(timeStr.substring(2, 4));
        }
        // 문자열인 경우
        else if (typeof timeValue === 'string') {
            // 콜론이 있는 경우
            if (timeValue.includes(':')) {
                const parts = timeValue.split(':');
                hours = parseInt(parts[0]);
                minutes = parseInt(parts[1]);
            }
            // 4자리 숫자 문자열인 경우
            else if (/^\d{4}$/.test(timeValue)) {
                hours = parseInt(timeValue.substring(0, 2));
                minutes = parseInt(timeValue.substring(2, 4));
            }
            else {
                return timeValue; // 변환 불가능한 경우 원본 반환
            }
        }
        else {
            return timeValue; // 변환 불가능한 경우 원본 반환
        }
        
        // 유효성 검사
        if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
            return timeValue; // 유효하지 않은 시간인 경우 원본 반환
        }
        
        // HH:MM 형태로 포맷팅
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    } catch (error) {
        console.warn('시간 포맷팅 오류:', timeValue, error);
        return timeValue; // 오류 발생시 원본 값 반환
    }
}

function filterDataBySpec(data) {
    // "규격" 컬럼이 있는지 확인
    if (data.length === 0) return data;
    
    const headers = Object.keys(data[0]);
    const hasSpecColumn = headers.includes('규격');
    
    if (!hasSpecColumn) {
        console.warn('데이터에 "규격" 컬럼이 없습니다.');
        return data; // 규격 컬럼이 없으면 원본 데이터 반환
    }
    
    // "규격" 값이 있는 행만 필터링 (null, undefined, 빈 문자열 제외)
    const filteredData = data.filter(row => {
        const specValue = row['규격'];
        return specValue !== null && 
               specValue !== undefined && 
               specValue !== '' && 
               String(specValue).trim() !== '';
    });
    
    console.log(`총 ${data.length}행 중 ${filteredData.length}행이 유효한 규격을 가지고 있습니다.`);
    
    return filteredData;
}

function displayDataAsTable(data) {
    const tableContainer = document.getElementById('data-table');
    
    if (data.length === 0) {
        tableContainer.innerHTML = '<p>표시할 데이터가 없습니다.</p>';
        return;
    }

    // 데이터 개수 정보 추가
    let tableHTML = `<div class="data-info">총 ${data.length}개의 유효한 데이터가 있습니다.</div>`;
    
    // 테이블 생성
    tableHTML += '<table class="data-table">';
    
    // 헤더 생성 (첫 번째 행의 키들을 헤더로 사용)
    const headers = Object.keys(data[0]);
    tableHTML += '<thead><tr>';
    headers.forEach(header => {
        tableHTML += `<th>${header}</th>`;
    });
    tableHTML += '</tr></thead>';
    
    // 데이터 행 생성
    tableHTML += '<tbody>';
    data.forEach((row, index) => {
        tableHTML += `<tr data-row-index="${index}">`;
        headers.forEach(header => {
            const cellValue = row[header] || ''; // undefined 값 처리
            tableHTML += `<td>${cellValue}</td>`;
        });
        tableHTML += '</tr>';
    });
    tableHTML += '</tbody></table>';
    
    // 테이블을 DOM에 추가
    tableContainer.innerHTML = tableHTML;
}

// 데이터 접근을 위한 헬퍼 함수들
function getExcelData() {
    return window.excelData || [];
}

function getDataBySpec(specValue) {
    const data = getExcelData();
    return data.filter(row => row['규격'] === specValue);
}

function getAllSpecs() {
    const data = getExcelData();
    const specs = data.map(row => row['규격']).filter(spec => spec);
    return [...new Set(specs)]; // 중복 제거
}

// 콘솔에서 사용할 수 있도록 전역 함수로 등록
window.getExcelData = getExcelData;
window.getDataBySpec = getDataBySpec;
window.getAllSpecs = getAllSpecs;

// CSV 다운로드 함수
function downloadCSV() {
    if (!currentData || currentData.length === 0) {
        alert('다운로드할 데이터가 없습니다. 먼저 엑셀 파일을 업로드해주세요.');
        return;
    }

    try {
        // CSV 문자열 생성
        const csvContent = convertToCSV(currentData);
        
        // Blob 생성 (UTF-8 BOM 추가로 한글 깨짐 방지)
        const BOM = '\uFEFF';
        const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
        
        // 파일명 생성 (현재 날짜와 시간 포함)
        const now = new Date();
        const timestamp = now.toISOString().slice(0, 19).replace(/[:-]/g, '').replace('T', '_');
        const filename = `output_${timestamp}.csv`;
        
        // 다운로드 실행
        downloadFile(blob, filename);
        
        // 사용자에게 알림
        showNotification(`CSV 파일이 다운로드되었습니다: ${filename}`, 'success');
        
    } catch (error) {
        console.error('CSV 다운로드 오류:', error);
        alert('CSV 파일 생성 중 오류가 발생했습니다: ' + error.message);
    }
}

// 데이터를 CSV 형식으로 변환
function convertToCSV(data) {
    if (data.length === 0) return '';
    
    const headers = Object.keys(data[0]);
    
    // 헤더 행 생성
    const csvRows = [];
    csvRows.push(headers.map(header => `"${header}"`).join(','));
    
    // 데이터 행 생성
    data.forEach(row => {
        const values = headers.map(header => {
            const value = row[header] || '';
            // CSV 이스케이프 처리 (쌍따옴표와 쉼표 처리)
            const escapedValue = String(value).replace(/"/g, '""');
            return `"${escapedValue}"`;
        });
        csvRows.push(values.join(','));
    });
    
    return csvRows.join('\n');
}

// 파일 다운로드 실행
function downloadFile(blob, filename) {
    // 최신 브라우저의 경우
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
        // IE 지원
        window.navigator.msSaveOrOpenBlob(blob, filename);
    } else {
        // 다른 브라우저
        const link = document.createElement('a');
        const url = window.URL.createObjectURL(blob);
        
        link.href = url;
        link.download = filename;
        link.style.display = 'none';
        
        document.body.appendChild(link);
        link.click();
        
        // 정리
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
    }
}

// 알림 메시지 함수
function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.textContent = message;
    
    // 알림 스타일
    notification.style.position = 'fixed';
    notification.style.top = '20px';
    notification.style.right = '20px';
    notification.style.padding = '15px 20px';
    notification.style.borderRadius = '5px';
    notification.style.color = 'white';
    notification.style.fontWeight = 'bold';
    notification.style.zIndex = '1000';
    notification.style.minWidth = '250px';
    notification.style.boxShadow = '0 4px 6px rgba(0, 0, 0, 0.1)';
    
    // 타입별 색상
    switch(type) {
        case 'success':
            notification.style.backgroundColor = '#2ecc71';
            break;
        case 'error':
            notification.style.backgroundColor = '#e74c3c';
            break;
        case 'info':
            notification.style.backgroundColor = '#3498db';
            break;
        default:
            notification.style.backgroundColor = '#95a5a6';
    }
    
    document.body.appendChild(notification);
    
    // 3초 후 자동 제거
    setTimeout(() => {
        if (notification.parentNode) {
            notification.parentNode.removeChild(notification);
        }
    }, 3000);
    
    // 클릭시 즉시 제거
    notification.addEventListener('click', () => {
        if (notification.parentNode) {
            notification.parentNode.removeChild(notification);
        }
    });
}