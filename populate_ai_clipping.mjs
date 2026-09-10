import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath = "outputs/news-clipping-template/뉴스_클리핑_양식.xlsx";
const outputDir = "outputs/news-clipping-2026-09-09";
await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const sheet = workbook.worksheets.getItem("뉴스 클리핑");

const articles = [
  [new Date(2026, 8, 9), "노인 우울증 AI로 조기 대응…삼성서울병원, BCI 융합 연구 착수", "삼성서울병원이 AI와 뇌-컴퓨터 인터페이스를 결합해 노년기 우울증을 조기에 대응하는 연구를 시작했다.", new Date(2026, 8, 8), "데일리안", "https://m.dailian.co.kr/news/view/1687741/%EB%85%B8%EC%9D%B8-%EC%9A%B0%EC%9A%B8%EC%A6%9D-AI%EB%A1%9C-%EC%A1%B0%EA%B8%B0-%EB%8C%80%EC%9D%91%EC%82%BC%EC%84%B1%EC%84%9C%EC%9A%B8-2026"],
  [new Date(2026, 8, 9), "서울 강서구, AI 챗봇 서비스 가동…24시간 행정정보 제공", "강서구가 10개 누리집의 행정 데이터를 통합한 AI 챗봇을 운영해 24시간 민원 정보를 제공한다.", new Date(2026, 8, 8), "데일리안", "https://m.dailian.co.kr/news/view/1687558/%EC%84%9C%EC%9A%B8-%EA%B0%95%EC%84%9C%EA%B5%AC-AI-%EC%B1%97%EB%B4%87-%EC%84%9C%EB%B9%84%EC%8A%A4-%EA%B0%80%EB%8F%992-2026"],
  [new Date(2026, 8, 9), "국가AI전략위 출범 1주년…국민·기업 체감할 성과 만든다", "국가AI전략위가 대형 AI 프로젝트 연계를 강화해 산업·공공 등 분야의 성과 창출에 집중하겠다고 밝혔다.", new Date(2026, 8, 8), "지디넷코리아", "https://zdnet.co.kr/view/?no=20260908135131"],
  [new Date(2026, 8, 9), "행안부, 조지아와 AI·디지털정부 협력 본격화…코카서스 진출 교두보 마련", "행정안전부가 조지아와 AI·디지털정부 협력을 강화해 국내 ICT 기업의 코카서스 진출 기반을 마련한다.", new Date(2026, 8, 8), "지디넷코리아", "https://zdnet.co.kr/news/?from=pc&lstcode=0200"],
  [new Date(2026, 8, 9), "중기 AI 활용 걸음마 수준 … AI 잘 쓰게해야 일자리 생겨", "전문가는 높은 도입 비용으로 중소기업의 AI 활용이 낮다며 창업·투자 생태계 강화를 제안했다.", new Date(2026, 8, 8), "매일경제", "https://www.mk.co.kr/news/business/12147393"],
  [new Date(2026, 8, 9), "국가AI委, 정책서 실행으로 메가 프로젝트 성과 챙긴다", "국가인공지능전략위원회가 3대 메가프로젝트와 7대 SEED 프로젝트의 실행·성과 관리 중심으로 운영을 전환한다.", new Date(2026, 8, 8), "매일경제", "https://www.mk.co.kr/news/it/12147353"],
  [new Date(2026, 8, 9), "AI와의 성공적인 공존 찾는다...27회 세계지식포럼 개막", "세계지식포럼이 AI의 활용과 위험, 인간과의 공존을 주제로 서울에서 개막했다.", new Date(2026, 8, 8), "매일경제", "https://www.mk.co.kr/news/business/12147513"],
];

for (let index = 0; index < articles.length; index += 1) {
  const row = 6 + index;
  const [clipDate, title, summary, publishedDate, publisher, url] = articles[index];
  sheet.getRange(`A${row}:E${row}`).values = [[clipDate, null, summary, publishedDate, publisher]];
  const safeTitle = title.replaceAll('"', '""');
  const safeUrl = url.replaceAll('"', '""');
  sheet.getRange(`B${row}`).formulas = [[`=HYPERLINK("${safeUrl}","${safeTitle}")`]];
  sheet.getRange(`B${row}`).format.font = { name: "Arial", size: 10, color: "#0563C1", underline: true };
}

workbook.recalculate();
const check = await workbook.inspect({ kind: "table", range: "뉴스 클리핑!A5:E12", include: "values,formulas", tableMaxRows: 8, tableMaxCols: 5 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 50 }, summary: "final formula error scan" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "뉴스 클리핑", range: "A1:E14", scale: 1.5, format: "png" });
await fs.writeFile(`${outputDir}/preview.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/AI_뉴스_클리핑_2026-09-09.xlsx`);
