window.exportToExcel = (jsonData, fileName) => {
    const ws = XLSX.utils.json_to_sheet(jsonData);

    // 헤더 스타일 적용
    const headerCells = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1']; // 헤더 셀들의 위치
    headerCells.forEach(cell => {
        if (ws[cell]) {
            ws[cell].s = {
                fill: {
                    fgColor: { rgb: "FB6B90" } // 배경색 (fb6b90 색상)
                },
                font: {
                    bold: true, // 글씨 굵게
                    color: { rgb: "FFFFFF" } // 글자색 흰색
                },
                alignment: {
                    horizontal: "center", // 텍스트 중앙 정렬
                    vertical: "center" // 세로 중앙 정렬
                }
            };
        }
    });

    // 숫자 포맷 적용 (계약금, 잔금, 최종금액 열)
    const numericColumns = ['D', 'E', 'F']; // 계약금, 잔금, 최종금액 열
    numericColumns.forEach(col => {
        for (let row = 2; row <= ws['!ref'].split(':')[1].substring(1); row++) { // 첫 번째 행부터 마지막 행까지
            const cellAddress = `${col}${row}`;
            if (ws[cellAddress]) {
                ws[cellAddress].s = {
                    numFmt: '#,##0' // 숫자 포맷: 1,000,000 형식
                };
                ws[cellAddress].v = ws[cellAddress].v.toString();
            }
        }
    });

    // 열 너비 자동 조정
    const columnWidths = {
        'A': 150,
        'B': 150,
        'C': 150,
        'D': 150,
        'E': 150,
        'F': 150,
        'G': 150,
        'H': 150,
        'I': 150
    };
    for (let col in columnWidths) {
        ws['!cols'] = ws['!cols'] || [];
        ws['!cols'][col.charCodeAt(0) - 65] = { wpx: columnWidths[col] }; // 열 크기 지정
    }

    // 워크북 생성 및 파일 저장
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Checklist");
    XLSX.writeFile(wb, fileName + ".xlsx");
};
