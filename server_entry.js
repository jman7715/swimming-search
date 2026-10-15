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
                <h2>신청자 명단 조회 entry </h2>
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
        const workbook = xlsx.readFile('swimmers.xlsx'); 
        const sheetName = '선수명단';
        const sheet = workbook.Sheets[sheetName];
        
        // 항목명(헤더)을 키값으로 하는 객체 배열 형태로 변환
        const dataRows = xlsx.utils.sheet_to_json(sheet);

        // 항목명 기준으로 데이터 찾기
        const foundRow = dataRows.find(row => {
            const rowName = String(row['이름'] || '').trim(); 
            const rawPhone = String(row['연락처'] || row['연락처4'] || '').trim(); 
            const rowPhone4 = rawPhone.slice(-4).padStart(4, '0'); // 뒤 4자리 추출 보장

            return rowName === inputName && rowPhone4 === inputPhone4;
        });

        let resultHtml = '';
        if (foundRow) {
            const regDate = foundRow['접수날짜'] || ''; 
            const name = foundRow['이름'] || ''; 
            const gender = foundRow['성별'] || ''; 
            const club = foundRow['소속클럽'] || ''; 
            const desiredEvent = foundRow['희망종목'] || ''; 
            const expectedTime = foundRow['예상기록'] || ''; 
            const souvenir1 = foundRow['기념품1'] || ''; 
            const entryFee = foundRow['참가비현황'] || ''; 

            resultHtml = `
                <div class="result-box">
                    <h3>[ 신청 정보 ]</h3>
                    <p><strong>접수일시:</strong> ${regDate}</p>
                    <p><strong>이름:</strong> ${name}</p>
                    <p><strong>성별:</strong> ${gender}</p>
                    <p><strong>소속:</strong> ${club}</p>
                    <p><strong>희망 종목:</strong> ${desiredEvent}</p>
                    <p><strong>예상 기록:</strong> ${expectedTime}</p>
                    <p><strong>셔츠:</strong> ${souvenir1}</p>
                    <p><strong>참가비현황:</strong> ${entryFee}</p>
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
                <p>엑셀 파일 이름(swimmers.xlsx)이 올바른지 확인해주세요.</p>
                <a href="/">돌아가기</a>
            </div>
        `);
    }
});

app.listen(PORT, () => {
    console.log(`서버가 실행되었습니다: http://localhost:${PORT}`);
});