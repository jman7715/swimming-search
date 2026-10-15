const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs'); // 파일 읽기를 위한 fs 모듈 추가
const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname)));

// 메인 입력 화면 (GET /)
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html lang="ko">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>신청자 명단 조회</title>
            <link rel="stylesheet" href="/style.css">
        </head>
        <body>
            <div class="container">
                <h2>신청자 명단 조회 entry</h2>
                <form action="/search" method="POST">
                    <div class="form-group">
                        <label for="name">이름</label>
                        <input type="text" id="name" name="name" required placeholder="이름을 입력하세요">
                    </div>
                    <div class="form-group">
                        <label for="phone4">연락처 뒤 4자리</label>
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
        const workbook = xlsx.readFile('entry_list.xlsx'); 
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        const dataRows = rows.slice(1); // 1행(헤더) 제외하고 2행부터 데이터로 취급

        // 엑셀 열 인덱스 기준: 
        // A열(0): 접수일시, B열(1): 이름, D열(3): 연락처뒤4자리
        const foundRow = dataRows.find(row => {
            const rowName = String(row[1] || '').trim(); 
            const rawPhone = String(row[3] || '').trim(); 
            const rowPhone4 = rawPhone.slice(-4).padStart(4, '0'); // 뒤 4자리 추출 보장

            return rowName === inputName && rowPhone4 === inputPhone4;
        });

        let resultHtml = '';
        if (foundRow) {
            const regDate = foundRow[0] || '';   // 접수일시 (A열 / 인덱스 0)
            const name = foundRow[1] || '';      // 이름 (B열 / 인덱스 1)
            const gender = foundRow[4] || '';    // 성별 (E열 / 인덱스 4)
            const team = foundRow[5] || '';      // 소속 (F열 / 인덱스 5)
            const sportGame = foundRow[6] || ''; // 희망 종목 (G열 / 인덱스 6)
            const expectedRecord = foundRow[7] || ''; // 예상기록 (H열 / 인덱스 7)
            const shirt = foundRow[8] || '';     // 셔츠 (I열 / 인덱스 8)
            const entryFee = foundRow[9] || '';     // 셔츠 (J열 / 인덱스 9)

            resultHtml = `
                <div class="result-box">
                    <h3>[ 신청 정보 ]</h3>
                    <p><strong>접수일시:</strong> ${regDate}</p>
                    <p><strong>이름:</strong> ${name}</p>
                    <p><strong>성별:</strong> ${gender}</p>
                    <p><strong>소속:</strong> ${team}</p>
                    <p><strong>희망 종목:</strong> ${sportGame}</p>
                    <p><strong>예상 기록:</strong> ${expectedRecord}</p>
                    <p><strong>셔츠:</strong> ${shirt}</p>
                    <p><strong>참가비:</strong> ${entryFee}</p>
                </div>
            `;
        } else {
            resultHtml = `
                <div class="error-box">
                    <p>일치하는 신청자 정보를 찾을 수 없습니다.<br>이름과 연락처 뒤 4자리를 다시 확인해주세요.</p>
                </div>
            `;
        }

        // notice_entry.txt 파일 읽어오기 (없으면 기본 안내 문구 출력)
        let noticeText = "등록된 공지사항이 없습니다.";
        try {
            if (fs.existsSync('notice_entry.txt')) {
                noticeText = fs.readFileSync('notice_entry.txt', 'utf8').replace(/\n/g, '<br>');
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
                <title>신청자 조회 결과 </title>
                <link rel="stylesheet" href="/style.css">
            </head>
            <body>
                <div class="container">
                    <h2>신청자 명단 조회 결과 entry</h2>
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
                <p>엑셀 파일 이름(entry_list.xlsx)이 올바른지 확인해주세요.</p>
                <a href="/">돌아가기</a>
            </div>
        `);
    }
});

app.listen(PORT, () => {
    console.log(`서버가 실행되었습니다: http://localhost:${PORT}`);
});