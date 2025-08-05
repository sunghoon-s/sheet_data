// 분석 버튼에 이벤트 리스너 추가
const analyzeBtn = document.getElementById('analyze-btn');
analyzeBtn.addEventListener('click', loadAndAnalyzeData);

function loadAndAnalyzeData() {
    // data_erp.xlsx 파일을 fetch로 읽기
    fetch('data_erp.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('파일을 찾을 수 없습니다.');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            // SheetJS를 사용해 엑셀 파일 읽기
            const workbook = XLSX.read(data, { type: 'array' });

            // 첫 번째 시트 이름 가져오기
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // 시트 데이터를 JSON 형태로 변환
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            console.log('원본 데이터:', jsonData); // 개발자 도구에서 데이터 확인

            // "규격" 헤더가 없거나 빈 행 필터링
            const filteredData = filterDataBySpec(jsonData);
            console.log('필터링된 데이터:', filteredData);

            // 전역 변수에 데이터 저장 (향후 사용을 위해)
            window.excelData = filteredData;

            // 데이터를 테이블로 표시
            displayDataAsTable(filteredData);
        })
        .catch(error => {
            console.error('파일 로드 오류:', error);
            alert('data_erp.xlsx 파일을 찾을 수 없습니다. 파일이 같은 폴더에 있는지 확인해주세요.');
        });
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