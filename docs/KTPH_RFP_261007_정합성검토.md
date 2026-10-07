# KTPH RFP 발표자료(261007) 정합성 검토

- 대상: `KTPH_RFP_MONIT_Presentation_261007.pdf` (26장)
- 사용처: 2026-10-08(목) 15:30 ALPS 회의 — KTPH·TTSH·WH RFP, 안건은 **유지보수비 필요성 + 설치 과정 설명** (볼트 `NHG.md`, 10/6 주간회의 안건 13)
- 대조한 볼트 노트(Google Drive 동기화본): `NHG.md`, `MECS PRO.md`, `모닛.md`, `제품 개발·품질 현황.md`, `KTPH Smart Diaper Sensor - Trial in B76 20251203.md`(+ 원본 PDF), `Eng_MECS PRO Gateway_Introduction_redesign260519.md`, `Intega Healthcare.md`, 저장소의 `build_deck.js`(RFP·Master Agreement 근거)

심각도: 🔴 발표 전 반드시 수정 · 🟠 수정 권장 · 🟡 다듬기

---

## A. 사실 정합성 (볼트·원자료와 충돌)

| # | 장 | 덱 내용 | 볼트·원자료 | 심각도 |
|---|---|---|---|---|
| A1 | 7 | 트라이얼 보고서를 **"Monit Diapers Sensor Trial in Singhealth"**, "IAD in Singapore & Singhealth", "In Singhealth: 33 patients" 로 표기 | 원본 보고서(2025-12-03): **"Monit Diapers Sensor Trial in KTPH"**, "In KTPH: 33 patients with IAD in 2024". KTPH는 NHG 소속이고 SingHealth는 별개 그룹(`NHG.md` 접점 지도) | 🔴 |
| A2 | 7 | Phase 1 기간 **24 Nov 2025 – 30 Nov 2025** | 원본: **9 July 2025 – 24 August 2025 (Week 1–7)**, 15명, Ward B76 | 🔴 |
| A3 | 6 | 감지 정확도 **92%**, 경쟁사 **~30%** | KTPH 실측: **정확도 76.2%, 민감도 72.7%, 특이도 82.6%** (7주 평균, 드롭아웃 제외). 92%·30%의 출처는 볼트에 없음. 2장에서 "Clinically validated … by KTPH"라고 하므로 KTPH 앞에서 수치가 직접 충돌 | 🔴 |
| A4 | 2, 3 | ISO 27001/27017/27018 **"11월 인증/등록 예정"** | 9/10 홈페이지에 공개한 「MONIT Corp Data Protection Notice v1.0」: **"2027년 상반기 취득 목표"**. 공개 약속과 덱이 다름 | 🔴 |
| A5 | 21–22 | 게이트웨이 접속 비밀번호 **00000000**, 사무실 SSID(SeochoAICT) 화면 노출 | 싱가포르 보안심사가 **기기별 고유 비밀번호를 필수 요구** → 9/14 "엄격 기관용 펌웨어 SKU" 결정. 다인시스 작업 미착수(9/28 대표 직접 조율). 덱이 요구사항 미충족 상태를 그대로 보여줌 | 🔴 |
| A6 | 24 | 자격증명 암호화 저장, **A/B 펌웨어 슬롯·실패 시 롤백**, 병원 IT의 **원격 자격증명 변경** | 볼트 근거 없음. 게이트웨이 칩셋·펌웨어는 외주(다인시스) 소관이고 확인 기록이 없음. 확인 안 되면 표현 완화 필요 | 🟠 |
| A7 | 3 vs 20 | 3장: 게이트웨이 **CE·KC 취득 완료**(번호 기재) / 20장 스펙표: 게이트웨이 **"Certification in progress — KC and CE (RED)"** | 9/28 보고: "게이트웨이 해외 인증 마무리"(개별 완료일·번호 미기록). 덱 안에서 서로 모순 | 🟠 |
| A8 | 3 | 게이트웨이 CE를 **EMC 2014/30/EU**로 표기 | BLE+Wi-Fi 무선기기는 통상 **RED 2014/53/EU** 대상. 인증서 원본 확인 필요. KC 번호 "R-R-mNT-SSG"는 잘린 것으로 보임 | 🟠 |
| A9 | 12 | 헤드라인 **S$96.1k "Average operating cost/Year"** | 96.1k = 480,825/5 (일회성 셋업비 포함). 운영비 평균은 466,539/5 = **S$93.3k** | 🟠 |
| A10 | 12 | **"5% annual escalation built in"** | 실제 표: 원격지원 ~7.2%/년, 예방정비 ~3%/년, 침투테스트 Y1 12,000 → Y2 7,000. 일괄 5%가 아님 (합계·소계는 모두 검산 일치, Y4·Y5 1 차이는 반올림) | 🟠 |
| A11 | 18 | Gantt에 **"Contract, deposit & insurance"** | 직전 커밋 e2396be에서 보증금(security deposit) 언급을 삭제하기로 함 | 🟠 |
| A12 | 18 vs 13 | Gantt "Mobilisation **(≥ 28 days)**" / 13장 "**Within** 28 days of award" | 같은 조건을 반대로 표기 | 🟠 |
| A13 | 13 vs 16 | 13장 **6단계** 방법론, 16장 **4단계**(Plan/Pilot/Roll out/Optimise) | 한 덱에 방법론이 두 개. 4단계를 6단계의 요약으로 명시하거나 하나로 통일 | 🟠 |
| A14 | 19 | Vanguard "Scale: Two residents", 날짜 5/20–7/3 | 볼트: PoC 성공 → **300명 구독 도입 결정(9/23), 수량 재조정 회신 대기(10/6)**. 확정 전이면 "구독 계약 진행 중" 정도로만 | 🟡 |
| A15 | 20 | HeartCare **"23 Osaka based facilities"**, "2,000 residents" | 볼트: **14개 요양시설, 센서 약 2,000개**, 기저귀 100만 장(출처 1건) | 🟠 |
| A16 | 20 | "Samsung Insurance Noble Life … Scale: 180 residents" | 볼트: **삼성 노블카운티 POC 25명 확정(9/28)**, Wi-Fi 공사 10월 말 후 착수. 180명 근거 없음 | 🟠 |
| A17 | 2, 20 | "NHIS (Korea FDA)", "NHIS (K-FDA)" | NHIS = 국민건강보험공단. 식약처(MFDS)와 다른 기관. "Selected as an innovative welfare device by NHIS (National Health Insurance Service, Korea)"로 | 🟠 |
| A18 | 20 | NHIS "2024 March", "for 1.23 Million" | 볼트에는 "2023 건보공단 예비급여 시범사업"만 있음. 연도·수치 확인 필요 | 🟡 |
| A19 | 2 | 글로벌 진출국에 **Poland** 포함 | 폴란드는 NEUCA **LOI 단계**(GIV JV 경유, 9/29). 진출국 목록에서는 빼거나 "LOI" 표기. Netherland → **Netherlands** | 🟡 |
| A20 | 2 | Address (Singapore) "9 Straits View, Marina One West #05-07" | 볼트·기존 스크립트에 MONIT 싱가포르 법인 기록 없음("Singapore entity / UEN still to add"). Intega 주소라면 "Singapore distributor (Intega)"로 | 🟡 |
| A21 | 5 | "100% Prevention of Pressure Ulcer", "25% diaper cost", "30% workload" | 볼트에 근거 없음(IR 근거는 "수작업 점검 84% 절감"). 병원 조달 문서에서 "100% 예방"은 의료적 효능 주장으로 읽힘 | 🟠 |
| A22 | 전체 | 푸터 RFP 번호 **KTPH-RFP-26-165-MJ** | RFP 원문 기반 스크립트는 165, 볼트는 **26-006-MJ**. 원문 기준이면 볼트를 정정 | 🟡 |

## B. 미완성·오류 표기

| # | 장 | 내용 | 심각도 |
|---|---|---|---|
| B1 | 21 | 큰 글씨 **"OK" / "NA"** (내부 체크 표시가 그대로 노출) | 🔴 |
| B2 | 19 | **"[INPUT REQUIRED]"** 플레이스홀더, "30 participans" 오타, SingHealth 기간 미기재 | 🔴 |
| B3 | 20 | Korean hospitals "Scope & timeline" 빈칸 | 🟠 |
| B4 | 21, 22 | 같은 장(게이트웨이 설치 1~6단계) **중복** | 🟠 |
| B5 | 11 | 푸터 **"Attachment 1 to the MOU · MONIT Corp. · MECS PRO Mobile App"** (다른 문서에서 옮겨온 흔적) | 🟠 |
| B6 | 16 | WH·TTSH 화살표 안 잔여 글자("s", "5715S" 흐릿한 텍스트) | 🟠 |
| B7 | 17 | KTPH **Tower D가 두 줄**(7-series·8-series 구분 누락) | 🟡 |
| B8 | 20 | YouTube "English-**TBD**" | 🟡 |
| B9 | 8–11 | "Fuctionalities" → Functionalities | 🟡 |
| B10 | 6, 7 | 제목 문법: "make different level of sensing accuracy", "Monit finalized … care staffs" | 🟡 |
| B11 | 23 | "2.4 Ghz" → GHz, "Only required" → "only a 2.4 GHz Wi-Fi network is required" | 🟡 |

## C. 디자인 톤앤매너 불일치

| 항목 | 현재 상태 |
|---|---|
| 서체 | 표지: 세리프(Cambria 계열) / 본문 다수: 맑은고딕 계열 / 표: Arial / 23장: 세리프 본문 / 캡처 이미지 안: 각종 서체 → 최소 4종 혼재 |
| 섹션 라벨(주황 띠) | "• Introduction…", "· Network…", "2. Industry…", "NETWORK CONNECTIVITY…"(대문자) — 접두기호·대소문자·번호 체계 제각각 |
| 제목 영역 | ① 네이비 띠 + 가운데 정렬 2줄 제목(3·13~19장) ② 진한 네이비 + 왼쪽 정렬 + 부제 + 청록 밑줄(8~12장) ③ 띠 없음(4·20장) — 세 가지 양식 |
| 색상 | 네이비 #002060 계열, 진네이비 #0B2A4A 계열, 파랑 #1F5FD1 계열 박스(2장), 청록 #00A3AD, 주황 띠, 노랑 하이라이트(17장 표 합계) |
| 표 | 네이비 헤더 표(12·13장) / 연파랑 셀 표(3장 하단) / 엑셀 캡처 표(17장) / 회색 테두리 표(20장) |
| 푸터·쪽번호 | 푸터 있는 장/없는 장 혼재, 쪽번호 없음 |
| 이미지 | 다른 자료 캡처(7·17·19·20·21~23장)가 원 서체·색으로 들어와 톤이 깨짐 |

통일안(제안): 기존 `build_deck.js` 테마를 기준으로 — 제목 Cambria, 본문 Calibri, 네이비 #0B2A4A · 청록 #00A3AD · 회색 #6B7A8F · 연회색 #F1F4F8, 주황 띠 대신 상단 청록 섹션 라벨(대문자, 번호 체계 `1 · COMPANY INTRODUCTION`), 왼쪽 정렬 액션 타이틀, 전 장 공통 푸터 + 쪽번호, 표는 네이비 헤더·줄무늬 1종.

## D. 10/8 ALPS 회의 관점에서 빠진 것 (보강 후보)

1. **유지보수비가 왜 필요한가** — ALPS가 직접 물은 안건인데 12장은 금액표뿐. "없으면 무엇이 멈추는가"(서버·보안패치·인증서 갱신 200일 주기·침투테스트·24×7 모니터링·현장 예방정비) 1장
2. **KTPH B76 트라이얼 자체를 레퍼런스로** — 가장 강한 근거인데 레퍼런스(19·20장)에서 빠져 있음
3. **설치 과정 1장 요약** — 병원이 준비할 것(2.4 GHz SSID·비밀번호·영문 SSID, 콘센트, VLAN/MAC 화이트리스트) + 설치 순서 + 소요시간. 현재는 스마트폰 화면 캡처 3장
4. **데이터 보호 요약** — Azure 싱가포르 리전, 케어 이벤트 14일·감사로그 12개월 보존, 침해 시 1일 내 통지, DPO, 기기 등록 시 개인정보 미수집(기기 별명만)
