const bot = BotManager.getCurrentBot();

const webappurl = "https://script.google.com/macros/s/AKfycbzLQy7INWBz1ZQU1rKN8EosGZAWoA6Xu-bq3QKga1LheLCdIlQlKCsSMLI2BGh-qH4/exec";

const sheetName = "2025 2학기 시간표";

// input: "9/11 (목) 13-15시 다목적 신청 부탁드립니다!"
// output: ["9/11(목)","13-15","다목적"]
function divideMsg(str) {

    // 1. 숫자가 처음 나오는 위치 찾기
    const firstDigitIndex = str.search(/\d/);
    if (firstDigitIndex === -1) return null; // 숫자가 없으면 잘못된 입력
    else str = str.slice(firstDigitIndex);

    // 1. 공백 및 쓸데없는거 제거
    const noSpace = str.replace(/(\s+)|(시)|(요일)|(연습실)/g, "");
    const changeBar = noSpace.replace(/(~)/g, "-");

    // 2. "장소" 이후 문자열 제거
    const trimmed = changeBar.replace(/(다목적|마루|방음|매트).*$/, "$1");
    
    // 3. 정규식 매칭 (날짜(요일) / 시간 / 장소)
    const regex = /^(\d{1,2}\/\d{1,2}\(([월화수목금토일])\))(\d{1,2})-(\d{1,2})(다목적|마루|방음|매트)$/;
    const match = trimmed.match(regex);
    if (!match) return "001"; // 예외 1: 양식 불일치

    // 여기까지 오면 형식은 올바르다고 가정
    const fullDateStr = `${match[1]}(${match[2]})`;     // ex) "9/6(토)"
    const weekdayStr = match[2];                        // ex) "토"
    const startHour = parseInt(match[3], 10);           // ex) 10
    const endHour = parseInt(match[4], 10);             // ex) 15
    const room = match[5];                              // ex) "마루"

    // 4. 시간 유효성 검사 (예외 4. 12-15시)
    if (
        isNaN(startHour) || isNaN(endHour)
        || startHour < 9 || endHour > 24 || startHour >= endHour ) return "002"; // 시간형식불일치

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
    if (candidateDate < today || candidateDate > fourWeeksLater) return "003"; // 과거or4주뒤

    // 7. 요일 검사 (예외 3)
    const weekdays = ["일", "월", "화", "수", "목", "금", "토"];
    const realWeekday = weekdays[candidateDate.getDay()];
    if (realWeekday !== weekdayStr) return "004";// "날짜와 요일 불일치";

    // 모든 조건 통과 → 배열 반환
    return [`${month}/${day}(${weekdayStr})`, `${startHour}-${endHour}`, room];
}

// input: ["9/11(목)","13-15","다목적"]
// output: "E26:E27"
function getCellRange(input, todayStr) {
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

// read sheet -> 배열 형태로 반환 (병합된것도 읽기)
function read(range) {
    // Jsoup execute()를 쓰면 body를 문자열로 뽑기 쉬움
    var url = webappurl + "?sheet=" + encodeURIComponent(sheetName) + "&range=" + range;
    var res = org.jsoup.Jsoup.connect(url) .ignoreContentType(true) .method(org.jsoup.Connection.Method.GET).execute();
            
    var body = res.body(); // JSON 문자열
    var data = JSON.parse(body);
    if (data.ok) {
        // 2차원 배열 -> 보기 좋게 문자열로
        // var lines = data.values.map(row => row.join(" | ")).join("\n");
        if (data.values.join("") == "") return "신청 가능";
        return (data.values);
    } else {
        return ("에러: " + data.error);
    }

}

// 마지막 행 아래에 추가하는 기능이라 잘 안쓸듯?
function append() {
    var payload = {
        sheet: "Sheet1",
        mode: "append",
        values: ["이름", "점수"],
        backgrounds: ["#d9ead3", "#cfe2f3"], // A열-연두, B열-하늘
        fontColors: ["#000000", "#ff0000"], // A열-검정, B열-빨강
        bold: true
    };
}

// 시트에 반영하는 부분
function write(cellRange, cellValue, cellColor) {
    var payload = {
        sheet: sheetName,
        mode: "update",
        range: cellRange,
        values: makeArray(cellValue, cellRange),
        backgrounds: makeArray(cellColor, cellRange),
        fontColors: makeArray("#000000", cellRange),
        bold: false
    };

    var res = org.jsoup.Jsoup
        .connect(webappurl)
        .header("Content-Type", "application/json")
        .requestBody(JSON.stringify(payload))
        .ignoreContentType(true)
        .method(org.jsoup.Connection.Method.POST)
        .execute();

    var data = JSON.parse(res.body());
    return (data.ok ? "추가 성공!" : "실패: " + data.error);
}

// 시트에 반영 2 (병합 모드)
function mergeWrite(cellRange, cellValue, cellColor) {
    var payload = {
        sheet: sheetName,
        mode: "mergeWrite",
        range: cellRange,
        value: cellValue,
        backgrounds: makeArray(cellColor, cellRange)
    };

    var res = org.jsoup.Jsoup
        .connect(webappurl)
        .header("Content-Type", "application/json")
        .requestBody(JSON.stringify(payload))
        .ignoreContentType(true)
        .method(org.jsoup.Connection.Method.POST)
        .execute();

    var data = JSON.parse(res.body());
    return (data.ok ? "추가 성공!" : "실패: " + data.error);
}

// input: ("aaa", "B4:C6")
// output: [["aaa","aaa"],["aaa","aaa"],["aaa","aaa"]]
function makeArray(str, range) {
    // 열 문자(A~Z...)를 숫자로 변환하는 헬퍼 함수
    function colToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
        num = num * 26 + (col.charCodeAt(i) - 64); // 'A' → 1
        }
        return num;
    }

    // 셀 주소("B4") → {col: number, row: number}
    function parseCell(cell) {
        const match = cell.match(/([A-Z]+)([0-9]+)/);
        return {
        col: colToNumber(match[1]),
        row: parseInt(match[2], 10)
        };
    }

    // 범위 분리 ("B4:C6")
    const [startCell, endCell] = range.split(":");
    const start = parseCell(startCell);
    const end = parseCell(endCell);

    const numRows = end.row - start.row + 1;
    const numCols = end.col - start.col + 1;

    // 2차원 배열 생성
    const result = Array.from({ length: numRows }, () =>
        Array.from({ length: numCols }, () => str)
    );

    return result;
}

function makeKakaoMsg(array, requestor) {
    if (array.length == 0) return null;

    var result = [];
    for (j=0; j<array.length; j++) {

        var info = array[j].split(" ->")[0].split(",");
        mergeWrite(getCellRange(info, getDate()), requestor + " 신청", "#ffff00");      // 시트에 반영하기
        result.push(info[0] + " " + info[1] + " " + info[2] + "연습실")
        
    }
    // return [result, no];
    
    if (result.length == 0) return null;
    else {
        result.push("라 신청합니다!");
        return result.join("\n");
    }
    
}

// 현재 날짜 반환 
// output: "11/9"
function getDate() {
    const today = new Date();
    const month = today.getMonth() + 1; // 월은 0부터 시작하므로 +1
    const day = today.getDate();

    return `${month}/${day}`;
}

function onMessage(msg) {
    try {

        const placesList = ["다목적", "마루", "방음", "매트"];
        if (msg.room == "김건우" && !msg.content.includes("취소")
            && placesList.some(placesList => msg.content.includes(placesList))) {
            // msg.reply("test");

            var input = msg.content.split("\n");
            var output = {
                "yes": {
                    "다목적": [],
                    "마루": [],
                    "방음": [],
                    "매트": []
                },
                "no": {
                    "다목적": [],
                    "마루": [],
                    "방음": [],
                    "매트": []
                },
                "미확인": []
            };

            const errors = ["001", "002", "003", "004"];
            for (i=0; i<input.length; i++) {
                var temp_msg = divideMsg(input[i]);
                if (temp_msg != null) {
                    if (errors.includes(temp_msg)) {
                        output["미확인"].push(input[i] + " -> " + temp_msg);

                        //msg.reply(read(temp_cell));
                    } else {
                        var temp_cell = getCellRange(temp_msg, getDate());
                        var temp_value = read(temp_cell);
                        if (temp_value == "신청 가능") {
                            output["yes"][temp_msg[2]].push(temp_msg + " -> " + temp_cell + " (" + temp_value + ")");
                        } else {
                            output["no"][temp_msg[2]].push(temp_msg + " -> " + temp_cell + " (" + temp_value + ")");
                        }
                    }
                }
            }
            msg.reply(JSON.stringify(output, null, 4));
/*
            for (i=0; i<placesList.length; i++) {
                var kakao = makeKakaoMsg(output[placesList[i]],  msg.author.name);
                if (kakao == null) continue;
                else {
                    msg.reply(kakao);
                }
            }
*/

            // 불발 건이 하나도 없으면 -> 자동으로 신청하고 신청 완료 메시지 출력
            var noCount = 0;
            for (i=0; i<placesList.length; i++) {
                noCount += output["no"][placesList[i]].length;
            }
            noCount += output["미확인"].length;

            if (noCount == 0) {
                for (i=0; i<placesList.length; i++) {
                    var kakao = makeKakaoMsg(output["yes"][placesList[i]],  msg.author.name);
                    if (kakao == null) continue;
                    else {
                        msg.reply(kakao);
                    }
                }
                msg.reply("[자동] 신청해드렸습니다. 스프레드시트 확인 부탁드립니다 :>\n신청인: " + msg.author.name);
            }

            
            
        }
        
        if (msg.room == "김건우" && msg.content == "hi") {
            
            msg.reply("hi");

            var a = {};

            // msg.reply(makeArray("김건우 신청", "B3:C6"));


            // msg.reply(write("B4:B6", "김건우 신청", "#ffff00"));
            // msg.reply(mergeWrite("B7:B9", "김건우 신청2", "#ffff00"));
    

            msg.reply("bye");
        }


    } catch(e) {
        msg.reply(e.lineNumber + "\n" + e);
    }


}

bot.addListener(Event.MESSAGE, onMessage);


[["aaa", "aaa"], ["aaa", "aaa"], ["aaa", "aaa"]]
