const bot = BotManager.getCurrentBot();

const webappurl = "https://script.google.com/macros/s/AKfycbw2nBaXmUhPHiSHmanmOtR41ROdzB0sEk9-bA6ygwN27Y6nDecGoiTxQ38ZIOm_0tmjew/exec";

const sheetName = "2025 2학기 시간표";

// const webappurl = "https://script.google.com/macros/s/AKfycbz3s2jUWCqfJEu5ZMH7i56wgHbXHTyBpokRS6EMzVDPThbxCOlQWf2OQKMxT3HNb_3LiA/exec";

function divideMsg(str) {

    // 1. 숫자가 처음 나오는 위치 찾기
    const firstDigitIndex = str.search(/\d/);
    if (firstDigitIndex === -1) return null; // 숫자가 없으면 잘못된 입력
    else str = str.slice(firstDigitIndex);

    // 1. 공백 및 쓸데없는거 제거
    const noSpace = str.replace(/(\s+)|(시)|(연습실)/g, "");

    // 2. "장소" 이후 문자열 제거
    const trimmed = noSpace.replace(/(다목적|마루|방음|매트).*$/, "$1");
    
    // 3. 정규식 매칭 (날짜(요일) / 시간 / 장소)
    const regex = /^(\d{1,2}\/\d{1,2}\(([월화수목금토일])\))(\d{1,2})-(\d{1,2})(다목적|마루|방음|매트)$/;
    const match = trimmed.match(regex);
    if (!match) return null; // 예외 1: 양식 불일치

    // 여기까지 오면 형식은 올바르다고 가정
    const fullDateStr = match[1];   // ex) "9/6(토)"
    const weekdayStr = match[2];    // ex) "토"
    const startHour = parseInt(match[3], 10);
    const endHour = parseInt(match[4], 10);
    const room = match[5];

    // 4. 시간 유효성 검사 (예외 4. 12-15시)
    if (
        isNaN(startHour) || isNaN(endHour)
        || startHour < 9 || endHour > 24 || startHour >= endHour ) return null; // return "시간형식불일치";

    // 5. 날짜 처리
    const today = new Date();
    today.setHours(0, 0, 0, 0); // 자정 기준(날짜만 비교하려고 나머지를 0으로)
    const [month, day] = fullDateStr.split("(")[0].split("/").map(Number);

    // 가까운 연도 계산 (올해 또는 내년)
    let year = today.getFullYear();
    let candidateDate = new Date(year, month - 1, day);

    // 오늘보다 과거라면 내년으로 계산
    if (candidateDate < today) {
        candidateDate = new Date(year + 1, month - 1, day);
        year = year + 1;
    }

    // 6. 날짜 범위 검사 (예외 2: 4주 이내)
    const fourWeeksLater = new Date(today);
    fourWeeksLater.setDate(today.getDate() + 7 * 4);
    if (candidateDate < today || candidateDate > fourWeeksLater) return null;// ;"과거or4주뒤"

    // 7. 요일 검사 (예외 3)
    const weekdays = ["일", "월", "화", "수", "목", "금", "토"];
    const realWeekday = weekdays[candidateDate.getDay()];
    if (realWeekday !== weekdayStr) return null;// "날짜와 요일 불일치";

    // 모든 조건 통과 → 배열 반환
    return [`${month}/${day}(${weekdayStr})`, `${startHour}-${endHour}`, room];
}


function getCellRange(input, todayStr = "9/3") {
  const [dateStr, timeStr, roomStr] = input;

  // ------------------------
  // 1. 주차 index 계산
  // ------------------------
  function getWeekStart(date) {
    const d = new Date(date);
    const day = d.getDay(); // 0=일 ~ 6=토
    const diff = d.getDate() - day + (day === 0 ? -6 : 1); // 월요일 시작
    return new Date(d.setDate(diff));
  }

  const today = new Date("2025/" + todayStr); // 예시: "9/3" → 2025/9/3
  const target = new Date("2025/" + dateStr.split("(")[0]); // "9/10(수)" → 2025/9/10

  const todayWeekStart = getWeekStart(today);
  const targetWeekStart = getWeekStart(target);

  const diffWeeks = Math.round((targetWeekStart - todayWeekStart) / (7 * 24 * 60 * 60 * 1000));
  const weekIndex = diffWeeks; // 오늘 주차=0, 다음주=1, 다다음주=2 …

  // ------------------------
  // 2. 주차 블록 시작 행
  // ------------------------
  const blockStartRow = 1 + weekIndex * 40;

  // ------------------------
  // 3. 연습실 오프셋
  // ------------------------
  const roomOffsets = {
    "마루": { col: "A", row: 3 },
    "방음": { col: "I", row: 3 },
    "다목적": { col: "A", row: 21 },
    "매트": { col: "I", row: 21 }
  };
  const { col, row } = roomOffsets[roomStr];

  // ------------------------
  // 4. 열/행 계산
  // ------------------------
  const weekDays = ["월", "화", "수", "목", "금", "토", "일"];
  const dayStr = dateStr.match(/\((.)\)/)[1]; // "(수)" → "수"
  const dayOffset = weekDays.indexOf(dayStr) + 1; // 월=1 ~ 일=7

  const [startHour, endHour] = timeStr.split("-").map(Number);
  const startOffset = startHour - 9;
  const endOffset = endHour - 10;

  // 열 계산 (col은 블록 시작 열, dayOffset 더함)
  const startCol = String.fromCharCode(col.charCodeAt(0) + dayOffset);
  // 행 계산 (row는 블록 시작 행, 시간 offset 더함)
  const startRow = blockStartRow + row + startOffset;
  const endRow = blockStartRow + row + endOffset;

  return `${startCol}${startRow}:${startCol}${endRow}`;
}

function read(range) {
    // Jsoup execute()를 쓰면 body를 문자열로 뽑기 쉬움
    var url = webappurl + "?sheet=" + encodeURIComponent(sheetName) + "&range=" + range;
    var res = org.jsoup.Jsoup.connect(url) .ignoreContentType(true) .method(org.jsoup.Connection.Method.GET).execute();
            
    var body = res.body(); // JSON 문자열
    var data = JSON.parse(body);
    if (data.ok) {
        // 2차원 배열 -> 보기 좋게 문자열로
        // var lines = data.values.map(row => row.join(" | ")).join("\n");
        return ("[조회 결과]\n" + data.values);
    } else {
        return ("에러: " + data.error);
    }

}

function onMessage(msg) { try {

    const placesList = ["다목적", "마루", "방음", "매트"];
    if (msg.room == "김건우" && placesList.some(placesList => msg.content.includes(placesList))) {
        // msg.reply("test");

        var input = msg.content.split("\n");
        var output = [];

        for (i=0; i<input.length; i++) {
            var temp_msg = divideMsg(input[i]);
            if (temp_msg != null) {
                var temp_cell = getCellRange(temp_msg);
                output.push(temp_msg + " -> " + temp_cell);

                msg.reply(read(temp_cell));
            }
        }
        msg.reply(output.join("\n"));
        
    }
    
    if (msg.room == "김건우" && msg.content == "hi") {
        
        msg.reply("hi");
        
        /*
        if (true) {
              const payload = {
                    sheet: sheetName,
                    mode: "append",
                    values: ["항목1","항목2","123"]
                };

                var res = org.jsoup.Jsoup
                    .connect(webappurl)
                    .header("Content-Type", "application/json")
                    .requestBody(JSON.stringify(payload))
                    .ignoreContentType(true)
                    .method(org.jsoup.Connection.Method.POST)
                    .execute();

                var data = JSON.parse(res.body());
                msg.reply(data.ok ? "추가 성공!" : "실패: " + data.error);
        }


        */
 
        


        
        


            msg.reply("bye");
        }
    } catch(e) {
        msg.reply(e.lineNumber + "\n" + e);
    }


}

bot.addListener(Event.MESSAGE, onMessage);


