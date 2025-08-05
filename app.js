// 분석 버튼에 이벤트 리스너 추가
const analyzeBtn = document.getElementById('analyze-btn');
analyzeBtn.addEventListener('click', loadAndAnalyzeData);

function loadAndAnalyzeData() {
    showProgress(0, '데이터 파일을 불러오는 중...');
    
    // 버튼 비활성화
    const analyzeBtn = document.getElementById('analyze-btn');
    analyzeBtn.disabled = true;
    analyzeBtn.textContent = '처리 중...';
    
    // data_erp.xlsx 파일을 fetch로 읽기
    fetch('data_erp.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('파일을 찾을 수 없습니다.');
            }
            showProgress(20, '파일을 읽는 중...');
            return response.arrayBuffer();
        })
        .then(data => {
            showProgress(40, '엑셀 데이터를 파싱하는 중...');
            
            // 비동기 처리를 위해 setTimeout 사용
            setTimeout(() => {
                try {
                    // SheetJS를 사용해 엑셀 파일 읽기
                    const workbook = XLSX.read(data, { type: 'array' });

                    // 첫 번째 시트 이름 가져오기
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    // 시트 데이터를 JSON 형태로 변환
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);

                    showProgress(60, `${jsonData.length}개 행을 처리하는 중...`);
                    console.log('원본 데이터:', jsonData);

                    // 비동기로 데이터 처리
                    setTimeout(() => {
                        processDataAsync(jsonData);
                    }, 100);
                    
                } catch (error) {
                    hideProgress();
                    resetButton();
                    console.error('데이터 처리 오류:', error);
                    alert('데이터 처리 중 오류가 발생했습니다: ' + error.message);
                }
            }, 100);
        })
        .catch(error => {
            hideProgress();
            resetButton();
            console.error('파일 로드 오류:', error);
            alert('data_erp.xlsx 파일을 찾을 수 없습니다. 파일이 같은 폴더에 있는지 확인해주세요.');
        });
}

async function processDataAsync(jsonData) {
    try {
        showProgress(70, '빈 열을 제거하는 중...');
        await delay(50);
        
        // 빈 헤더 열 제거
        const cleanedData = removeEmptyColumns(jsonData);
        console.log('빈 열 제거된 데이터:', cleanedData);

        showProgress(80, '데이터를 필터링하는 중...');
        await delay(50);
        
        // "규격" 헤더가 없거나 빈 행 필터링
        const filteredData = filterDataBySpec(cleanedData);
        console.log('필터링된 데이터:', filteredData);

        showProgress(90, '날짜와 시간을 포맷팅하는 중...');
        await delay(50);
        
        // 날짜 및 시간 포맷팅
        const formattedData = formatDateAndTime(filteredData);
        console.log('포맷팅된 데이터:', formattedData);

        showProgress(95, '테이블을 생성하는 중...');
        await delay(50);
        
        // 전역 변수에 데이터 저장 (향후 사용을 위해)
        window.excelData = formattedData;

        // 데이터를 테이블로 표시
        displayDataAsTable(formattedData);
        
        showProgress(100, `완료! ${formattedData.length}개의 데이터를 성공적으로 로드했습니다.`);
        
        // 2초 후 진행바 숨기기
        setTimeout(() => {
            hideProgress();
            resetButton();
        }, 2000);
        
    } catch (error) {
        hideProgress();
        resetButton();
        console.error('데이터 처리 오류:', error);
        alert('데이터 처리 중 오류가 발생했습니다: ' + error.message);
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
    analyzeBtn.textContent = '데이터 보기';
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