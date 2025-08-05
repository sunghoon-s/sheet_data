// 분석 버튼에 이벤트 리스너 추가
const analyzeBtn = document.getElementById('analyze-btn');
analyzeBtn.addEventListener('click', loadAndAnalyzeData);

// 차트를 그릴 canvas 요소 가져오기
const chartCanvas = document.getElementById('myChart');
let myChart = null; // 차트 객체를 저장할 변수

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

            console.log(jsonData); // 개발자 도구에서 데이터 확인

            // 데이터 분석 및 차트 생성 함수 호출
            analyzeAndDrawChart(jsonData);
        })
        .catch(error => {
            console.error('파일 로드 오류:', error);
            alert('data_erp.xlsx 파일을 찾을 수 없습니다. 파일이 같은 폴더에 있는지 확인해주세요.');
        });
}

function analyzeAndDrawChart(data) {
    // --- 데이터 분석 로직 (예시) ---
    // '항목'별 '값'을 합산한다고 가정해봅시다.
    // 예시 엑셀 데이터: [{항목: '과일', 값: 10}, {항목: '채소', 값: 20}, {항목: '과일', 값: 15}]

    const analysisResult = {};

    data.forEach(row => {
        const category = row['항목']; // 엑셀의 '항목' 열
        const value = row['값'];      // 엑셀의 '값' 열

        if (analysisResult[category]) {
            analysisResult[category] += value;
        } else {
            analysisResult[category] = value;
        }
    });

    const labels = Object.keys(analysisResult); // 차트의 X축 이름들 (예: ['과일', '채소'])
    const values = Object.values(analysisResult); // 차트의 Y축 값들 (예: [25, 20])

    // --- 차트 그리기 로직 ---
    if (myChart) {
        myChart.destroy(); // 이전에 그려진 차트가 있다면 파괴
    }

    myChart = new Chart(chartCanvas, {
        type: 'bar', // 막대그래프
        data: {
            labels: labels,
            datasets: [{
                label: '항목별 합계',
                data: values,
                backgroundColor: [
                    'rgba(255, 99, 132, 0.5)',
                    'rgba(54, 162, 235, 0.5)',
                    'rgba(255, 206, 86, 0.5)',
                    'rgba(75, 192, 192, 0.5)',
                ]
            }]
        },
        options: {
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}