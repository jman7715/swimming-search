const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs'); // 파일 읽기를 위한 fs 모듈 추가
const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname)));

// 엑셀의 시간 숫자 값을 HH:mm 형식으로 변환하는 함수
function excelTimeToStr(excelTime) {
    if (typeof excelTime !== 'number') return excelTime || '';
    const totalMinutes = Math.round(excelTime * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

// 메인 입력 화면 (GET /)
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html lang="ko">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>2026 아쿠아 마라톤 대진</title>
            <link rel="stylesheet" href="/style.css">
        </head>
        <body>
            <div class="container">
                <h2>2026 아쿠아 마라톤 대진</h2>
                <form action="/search" method="POST">
                    <div class="form-group">
                        <label for="name">선수 이름</label>
                        <input type="text" id="name" name="name" required placeholder="이름을 입력하세요">
                    </div>
                    <div class="form-group">
                        <label for="phone4">전화번호 뒤 4자리</label>
                        <input type="text" id="phone4" name="phone4" required maxlength="4" placeholder="뒤 4자리 입력" inputmode="numeric">
                    </div>
                    <button type="submit">조회하기</button>
                </form>
            </div>
        </body>
        </html>
    `);
});

// 검색 결과 처리 화면 (POST /search)
app.post('/search', (req, res) => {
    const inputName = req.body.name.trim();
    const inputPhone4 = req.body.phone4.trim().padStart(4, '0');

    try {
        const workbook = xlsx.readFile('swimmers.xlsx'); 
        const sheetName = '선수명단';
        const sheet = workbook.Sheets[sheetName];
        
        // 항목명(헤더)을 키값으로 하는 객체 배열 형태로 변환
        const dataRows = xlsx.utils.sheet_to_json(sheet);

        // 이름과 전화번호 뒤 4자리가 일치하는 행 찾기
        const foundRow = dataRows.find(row => {
            const rowName = String(row['이름'] || '').trim(); 
            const rawPhone = String(row['연락처4'] || row['연락처'] || '').trim(); 
            const rowPhone4 = rawPhone.slice(-4).padStart(4, '0');

            return rowName === inputName && rowPhone4 === inputPhone4;
        });

        let resultHtml = '';
        if (foundRow) {
            const name = foundRow['이름'] || '';           // 이름
            const gender = foundRow['성별'] || '';         // 성별
            
            // 종목: 숫자인 경우 천단위 콤마 및 'm' 추가 (예: 5000 -> 5,000m)
            let game = foundRow['종목'] || '';
            if (!isNaN(game) && game !== '') {
                game = Number(game).toLocaleString() + 'm';
            }

            const matchNo = foundRow['경기번호'] || '';    // 경기번호
            const bNo = foundRow['등번호'] || '';          // 등번호
            const standbyTime = excelTimeToStr(foundRow['등록시간']); // 등록시간
            const startTime = excelTimeToStr(foundRow['경기시간']);   // 경기시간

            resultHtml = `
                <div class="result-box">
                    <h3>[ 선수 정보 ]</h3>
                    <p><strong>이름:</strong> ${name}</p>
                    <p><strong>성별:</strong> ${gender}</p>
                    <p><strong>종목:</strong> ${game}</p>
                    <p><strong>경기 번호:</strong> 제 ${matchNo} 경기</p>
                    <p><strong>등번호:</strong> ${bNo}</p>
                    <p><strong>등록 시간:</strong> ${standbyTime}</p>
                    <p><strong>경기 시간:</strong> ${startTime}</p>
                </div>
            `;
        } else {
            resultHtml = `
                <div class="error-box">
                    <p>일치하는 선수 정보를 찾을 수 없습니다.<br>이름과 전화번호 뒤 4자리를 다시 확인해주세요.</p>
                </div>
            `;
        }

        // notice.txt 파일 읽어오기 (없으면 기본 안내 문구 출력, 줄바꿈 및 HTML 태그 지원)
        let noticeText = "등록된 공지사항이 없습니다.";
        try {
            if (fs.existsSync('notice.txt')) {
                noticeText = fs.readFileSync('notice.txt', 'utf8').replace(/\n/g, '<br>');
            }
        } catch (e) {
            console.error("공지사항 파일을 읽는 중 오류 발생:", e);
        }

        res.send(`
            <!DOCTYPE html>
            <html lang="ko">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>수영대회 대진 확인 결과</title>
                <link rel="stylesheet" href="/style.css">
            </head>
            <body>
                <div class="container">
                    <h2>2026 아쿠아 마라톤 대진</h2>
                    ${resultHtml}
                    
                    <div class="notice">
                        <h3>[ 공지사항 ]</h3>
                        <p>${noticeText}</p>
                    </div>

                    <div class="images">
                        <h3>[ 안내 이미지 ]</h3>
                        <img src="/info_01.jpg" alt="안내 이미지 1" onerror="this.style.display='none'">
                        <img src="/info_02.jpg" alt="안내 이미지 2" onerror="this.style.display='none'">
                    </div>

                    <a href="/" class="back-btn">돌아가기</a>
                </div>
            </body>
            </html>
        `);

    } catch (error) {
        console.error(error);
        res.send(`
            <div style="margin: 40px; font-family: Arial; text-align: center;">
                <h3>오류가 발생했습니다.</h3>
                <p>엑셀 파일 이름(swimmers.xlsx)이 올바른지 확인해주세요.</p>
                <a href="/">돌아가기</a>
            </div>
        `);
    }
});

app.listen(PORT, () => {
    console.log(`서버가 실행되었습니다: http://localhost:${PORT}`);
});