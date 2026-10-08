# ALPS 회의 예상 질문 20선 — 공격적·고난이도 (2026-10-08 15:30)

- 상대: ALPS(KTPH·TTSH·WH 조달 모회사) + 병원 IT·간호·조달 담당 가능
- 회의 안건: **유지보수비 필요성 + 설치 과정** (NHG.md, 10/6 주간회의 안건 13)
- 목표: **계약 확보.** 유지보수비를 지키는 것보다 "MONIT를 탈락시킬 이유"를 없애는 것이 먼저다
- 근거: 볼트(NHG·MECS PRO·모닛·제품 개발·품질 현황), KTPH B76 트라이얼 보고서, 비용 모델(build_cost_slide.js), 261008 덱

---

## 0. 회의 전에 머리에 넣을 것

**한 문장 포지션** — *"The system price builds it; the annual charge keeps it running, secure and certified for five years — at a price you can lock today."*

**숫자** (SGD, 5년, ex-GST)

| 항목 | 금액 |
|---|---|
| 5년 운영·유지보수 | 489,695 (연평균 97.9k) |
| 그중 클라우드 호스팅(Azure·WAF·SaaS) | 225.4k |
| 보안·인증(침투테스트·ISO) | 65.3k |
| 원격지원·24×7 모니터링 | 96.9k |
| 현장 예방정비(연 4회, 3개 병원) | 102.0k |
| 일회성 셋업 | 14,286 |

**협상 카드** (회의장에서 먼저 꺼내지 말 것. 상대가 금액을 압박하면 하나씩, 대가를 받고 내준다)

| 카드 | 5년 운영비 | 절감 | 받아낼 것 |
|---|---|---|---|
| 기준안 | 489.7k | — | — |
| ① 1년 차 침투테스트를 시스템 가격에 흡수 (사전 침투테스트가 원래 시스템 가격에 포함된 구조라 논리적) | 477.7k | −12.0k | 유지보수 항목 자체는 인정 |
| ② 인상률 5% → 3% | 약 472k | 약 −17.5k | 5년 고정가 확인 |
| ①+② | 약 460k | 약 −29.5k | 3개 병원 공통 과금 기준으로 채택 |
| 셋업비 14,286 면제 | — | −14.3k (일회성) | 파일럿 병동 일정 확정 |
| **하한선** — 인상률 0% + ① | 약 435k | 약 −54.5k | 이 아래로는 서비스 범위를 줄여야 함 (예방정비 연 4회 → 2회) |

※ 대표 확인 필요: 실제 하한선과 결정 권한. 위 금액은 비용 모델 기준 추정치다.

**절대 하지 말 것**
- 확인 안 된 수치 말하기: 92% 정확도, 욕창 100% 예방, ISO 이미 취득 등
- 볼트에 기록된 내부 사정 꺼내기: 개발자 1인, 외주사 지연, 매출 구조, "2억→1억", 2036 타임스탬프
- "Intega가 잘못해서" 같은 파트너 탓

---

## 1. 유지보수비 (회의 핵심 안건)

### Q1. "시스템 가격을 이미 내는데 왜 매년 유지보수비를 또 내야 하나? 이중 청구 아닌가?"
*Why do we pay an annual charge on top of the system price? Isn't this double charging?*

- **의도:** 유지보수 항목 자체를 삭제하려는 1차 공격
- **답변:** 두 비용은 대상이 다르다. 시스템 가격은 **한 번 만드는 것**(병원별 격리 환경 구축, SSO 연동, 하드웨어, 설치)이다. 연간 비용은 **5년 동안 켜 두는 것**(Azure 서버가 매 시간 과금되는 클라우드 비용, 보안 패치, 인증서 갱신, 침투테스트, 24×7 모니터링, 현장 점검)이다. 연간 비용의 46%가 클라우드 호스팅 실비라서, 이 비용이 없으면 1년 차 이후 알림이 멈춘다.
- **스크립트:** "The system price pays for building your environment once. The annual charge pays for running it — about half of it is Azure hosting in Singapore that is billed to us every hour. If we stop paying it, the nurses' alerts stop. There's no overlap: every line on slide 13 is a cost that recurs every year."
- **하지 말 것:** "다른 고객도 다 낸다" 같은 비교 논리

### Q2. "작년 10월 견적은 5년 21만 달러였다. 지금은 두 배가 넘는다. 설명하라."
*Your October 2025 quotation was S$210k over five years. Why has it more than doubled?*

- **의도:** 가격 신뢰도 공격. 가장 위험한 질문
- **답변:** 범위가 바뀌었다. 2025년 견적은 일반 클라우드 공용 환경 기준이었다. 이번 RFP는 공공병원 클라우드 보안 요구사항을 반영해 다음이 추가됐다: **병원별 격리 테넌시와 n+1 이중화, WAF, 연 1회 침투테스트, ISO 27001/27017/27018, 24×7 보안 모니터링, 병원 계정 SSO 연동**. 증가분은 거의 전부 이 보안 요구에서 나온다.
- **스크립트:** "The 2025 figure was for a standard shared cloud service. This RFP asks for a site-isolated tenancy per hospital, a WAF, annual penetration testing, ISO 27001/27017/27018 and 24×7 security monitoring. Those requirements are where the increase comes from — and if ALPS decides some of them aren't needed, we'll take them out line by line."
- **카드:** 마지막 문장으로 "범위를 빼면 가격도 빠진다"는 원칙을 보여준다. 먼저 금액을 깎지 않는다.

### Q3. "가격을 5년 고정한다면서 매년 5%씩 오른다. 모순 아닌가? 싱가포르 물가는 그만큼 오르지 않는다."
*You say prices are fixed, yet they rise 5% a year. Inflation here is nowhere near 5%.*

- **의도:** 인상률 공격
- **답변:** 계약상 5년간 가격을 바꿀 수 없으니, 미래의 인건비·클라우드 단가 상승을 **오늘 확정된 스케줄**로 넣은 것이다. 이렇게 하면 병원은 5년치 예산을 지금 확정할 수 있다. 다만 인상률은 조정할 수 있다.
- **스크립트:** "Because the charges can't change during the term, we priced the next five years today so your budget is certain. If ALPS prefers a lower escalation, we can discuss it — what matters to us is one fixed schedule for all three hospitals."
- **카드:** 압박이 계속되면 ② 인상률 3%(약 −17.5k)

### Q4. "유지보수를 빼고 시스템만 사면 어떻게 되나? 우리 IT가 직접 운영하겠다."
*What if we buy the system only and our IT runs it?*

- **의도:** 유지보수 계약 없이 가는 시나리오 탐색
- **답변:** 클라우드 SaaS라서 병원 IT가 MONIT 플랫폼을 직접 운영할 수 없다(소스·인프라는 MONIT 소유). 유지보수를 빼면 클라우드 호스팅과 보안 의무를 이행할 주체가 없어진다. 대안으로 **범위 축소형**을 제안한다: 현장 예방정비를 연 4회에서 2회로 줄이고, 1차 현장 대응을 병원 간호 챔피언·IT가 맡는다.
- **스크립트:** "It's a managed cloud service, so there's no version your IT team could run in-house. What we can do is shape the service: for example, two on-site rounds a year instead of four, with your ward champions handling first-line checks."

### Q5. "예방정비 연 4회가 왜 필요한가? 센서는 고장 나면 교체하면 되지 않나?"
*Why four preventive-maintenance rounds a year?*

- **답변:** 배터리 교체 주기(2~3개월)와 맞춘 것이다. 점검 때 하는 일: 릴레이 연결 상태, 배터리와 센서 상태, **병동 구조나 침대 배치가 바뀐 뒤의 커버리지 재측정**. 고장 난 뒤 대응하면 그사이 알림이 빠진다.
- **스크립트:** "Sensor batteries run two to three months, so the rounds follow that cycle — and each round re-tests coverage after any ward or bed changes, before a gap turns into a missed alert."
- **확인 필요:** 배터리 비용이 유지보수비에 포함되는지 회의 전에 대표가 확정할 것

---

## 2. 성능·임상 안전

### Q6. "KTPH 트라이얼 정확도가 76%, 민감도 72.7%다. 젖은 기저귀 4개 중 1개를 놓친다는 뜻 아닌가? 이걸 왜 사야 하나?"
*Your own trial at KTPH shows 72.7% sensitivity — you miss one wet diaper in four.*

- **의도:** 가장 강한 기술 공격. 상대가 보고서를 가지고 있다
- **답변 3단계:**
  1. **인정한다.** 이 수치는 KTPH 간호사들이 직접 측정한 실측값이고, 그래서 덱에 그대로 실었다.
  2. **그 뒤에 바뀐 것을 말한다.** Phase 1은 감도 3단계·태블릿 직결 방식이었다. 이후 감도를 5단계 + 개인 맞춤 15단계로 다시 설계했고 전 구간을 더 민감하게 조정했다. 연결은 병동 단위 릴레이 네트워크로 바꿨다. 놓친 사례의 상당수는 감도 설정과 연결 문제였다.
  3. **비교 기준을 바로잡는다.** 지금 기준은 "정해진 시간 순회"이고, 순회 사이에는 젖은 기저귀를 하나도 못 잡는다. 센서는 간호를 대체하는 게 아니라 **순회 사이를 메우는 것**이다.
- **승부수:** 파일럿 병동 인수 기준을 KTPH와 함께 정하자고 제안한다. 같은 방식으로 다시 재서 통과해야 확대 설치한다.
- **스크립트:** "Those are KTPH's own numbers, which is why they're on our slide. Since Phase 1 we've rebuilt sensitivity — five presets plus fifteen personal levels, tuned more sensitive across the range — and replaced tablet pairing with a ward relay network, which addresses most of the misses. And the comparison isn't 100%: between fixed rounds, today nobody catches a wet diaper. We propose the pilot ward is accepted against a detection threshold agreed with KTPH nursing before rollout continues."
- **하지 말 것:** 92% 언급, "트라이얼 방식이 잘못됐다"

### Q7. "알람 음소거·연결 문제로 드롭아웃이 있었다. 간호사 알람 피로는 어떻게 하나?"
*Your trial had dropouts from muted alarms and connection issues. Our nurses already have alarm fatigue.*

- **답변:** 오경보율(FP)은 9%였다. 이제 환자별로 감도를 조정할 수 있어서 과민한 환자는 둔감하게, 놓치는 환자는 민감하게 맞춘다. 알림 지속 시간(1~10초)도 조정된다. 알림은 병동 대시보드와 담당 기기로만 가고, 병원 전체 호출 체계와는 분리된다. 연결 문제는 릴레이 네트워크와 Wi-Fi 끊김 시 데이터 버퍼링(복구 후 재전송)으로 해결했다.
- **스크립트:** "False alarms were 9%. Sensitivity is now set per patient, alert duration is adjustable, and alerts go only to the ward dashboard and assigned devices. The connection drop-outs came from tablet pairing; the relay network buffers data and re-sends it after any Wi-Fi interruption."

### Q8. "센서를 환자 간에 재사용하나? 감염관리는?"
*Is the sensor reused between patients? What about infection control?*

- **답변:** 센서는 기저귀 **겉면**에 부착되고 피부에 닿지 않는다. 재사용하며 배터리만 교체한다. 환자 간 재사용 시 소독은 병원 감염관리팀(IPC) 프로토콜을 따르고, 파일럿 단계에서 IPC와 승인된 소독 절차를 문서로 확정하겠다.
- **스크립트:** "The sensor sits on the outside of the diaper, not on skin. Between patients it's cleaned to your IPC protocol — we'll agree and document the approved wipe-down with your IPC team during the pilot."
- **하지 말 것:** 검증 안 된 소독제 호환성 단정. 방수·방습 수준은 묻기 전에는 말하지 않는다(현재 버전은 방습 구조가 아님, IP67/68은 다음 버전). 물으면 "기저귀 겉면 사용 기준으로 설계됐고, 겉기저귀용 보호 필름을 제공한다"

### Q9. "기존 기저귀 공급사가 있다. 기저귀를 바꿔야 하나?"
*We have an existing diaper contract. Do we have to switch diapers?*

- **의도:** 기존 기저귀 계약과 충돌하는지 확인
- **답변:** **아니다.** Safeguard 스티커로 지금 쓰는 기저귀 그대로 사용한다. 기존 공급 계약에 영향이 없다.
- **스크립트:** "No. The Safeguard sticker fits the sensor to the diapers you already buy, so your current diaper contract is untouched."
- **하지 말 것:** 전용 기저귀 전략은 절대 언급하지 않는다

---

## 3. 보안·데이터

### Q10. "ISO 27001/27017/27018이 아직 없다. 인증 없는 업체에 환자 데이터를 맡기라는 건가?"
*You're not ISO-certified yet. Why should we trust you with patient data?*

- **답변:**
  1. 인증 심사가 진행 중이고 **11월 취득** 예정이다. 인증서가 나오는 대로 제출한다.
  2. 호스팅 인프라인 Microsoft Azure 싱가포르 리전은 이미 같은 ISO 인증을 보유하고 있다.
  3. 시스템이 처리하는 개인정보 자체가 최소화돼 있다. 기기 등록 시 이름·생년·성별을 받지 않고 기기 별명만 쓴다.
- **스크립트:** "Our certification audit is in progress for November and we'll submit the certificates on issue. The platform runs on Azure Singapore, which already holds ISO 27001, 27017 and 27018. And we hold very little to protect: no names, birth dates or gender — only a device nickname and soiling events."
- **조건 제시:** 필요하면 "인증서 제출을 최종 인수 조건으로 걸어도 좋다"고 말한다 → 신뢰 확보. 단, 11월 일정에 확신이 있을 때만

### Q11. "릴레이 기기 비밀번호가 기기마다 다른가? 초기 비밀번호가 다 00000000 아닌가?"
*Do all your relays share the same default password?*

- **의도:** 보안 심사에서 이미 지적된 항목. 기본 비밀번호 화면 캡처를 본 적 있을 수 있다
- **답변:** 일반 시장용은 공통 설정 비밀번호지만, **싱가포르 공공병원 납품분은 기기별 고유 비밀번호 펌웨어**로 공급하고 비밀번호 목록은 병원 IT에 넘긴다. 또 릴레이에는 개인정보가 저장되지 않고, 바깥으로만 연결된다(외부에서 릴레이로 접속하는 통로 없음).
- **스크립트:** "Units for your hospitals ship with a unique set-up password per device, and we hand the list to hospital IT. The relay stores no personal data and only makes outbound connections."
- **내부 리스크:** 펌웨어 이원화를 외주사가 아직 시작하지 않았다. 설치가 3개월 차부터라 일정은 맞출 수 있지만, 대표가 납기를 확정해 둘 것

### Q12. "릴레이가 중국산이다. 공공병원 네트워크에 중국산 IoT를 붙이라는 건가?"
*The relay is made in China. Why should we put it on a public-hospital network?*

- **답변:** 제조만 중국이고 **설계·펌웨어·서버는 MONIT(한국) 소유**다. IMDA·CE·KC에 등록돼 있다. 연결 방식이 위험을 줄인다: 병원 망에서는 바깥으로만 통신하는 일반 Wi-Fi 기기이고, 전용 IoT SSID/VLAN에 분리하고 MAC 화이트리스트로 관리할 수 있다. 매년 침투테스트에 릴레이를 포함하겠다.
- **스크립트:** "It's manufactured in China but designed, owned and updated by MONIT in Korea, and registered with IMDA. It's an outbound-only Wi-Fi client you can isolate on an IoT VLAN with MAC whitelisting, and we'll include it in the annual penetration test scope."

### Q13. "RFP는 병원 계정 로그인(SSO) 연동을 요구한다. 지금 되나?"
*Is single sign-on with our hospital accounts working today?*

- **답변:** 프로젝트 범위에 들어 있는 개발 항목이고, 시스템 가격에 포함된 병원 환경 구축의 일부다. **3개 병원 공통으로 한 번 구축**하는 방식을 제안한다. 병원마다 따로 만들지 않으니 일정과 비용이 줄고, 이후 ALPS 산하 다른 기관에도 그대로 쓸 수 있다. 착수는 설계 단계에서 병원 IT와 함께 한다.
- **스크립트:** "SSO is part of the build in the system price. We propose one integration shared by all three hospitals rather than three separate ones — faster now, and reusable for any other ALPS institution later."
- **하지 말 것:** "이미 완료" (미착수 상태)

### Q14. "MONIT가 망하거나 서비스를 중단하면 우리 데이터와 시스템은?"
*What happens to our data and service if MONIT fails or exits?*

- **답변:** 데이터는 병원별 격리 테넌시에 있고 언제든 CSV로 내보낼 수 있다. 계약 종료 시 데이터 반환·삭제 절차를 계약서에 넣겠다. 현장 지원은 싱가포르 파트너 Intega가 하고, 예비 부품은 싱가포르에 보관한다. 필요하면 소스코드 에스크로도 검토할 수 있다.
- **스크립트:** "Each hospital's data sits in its own tenancy and can be exported at any time; we'll write data return and deletion on exit into the contract. Spares are held in Singapore with Intega, and we're open to discussing source-code escrow."

---

## 4. 회사 역량·계약 구조

### Q15. "MONIT는 작은 회사다. 3개 병원 동시 구축과 5년 지원을 감당할 수 있나?"
*You're a small company. Can you deliver three hospitals and support them for five years?*

- **의도:** 재무·인력 리스크. 매출 구조는 절대 노출하지 않는다
- **답변:** 일본 Heart Care에서 14개 시설에 센서 약 2,000개를 운영 중이고, 한국에서는 국민건강보험공단 혁신 복지용구로 지정돼 있다. 같은 시스템이 3개 병원 규모를 이미 넘어서 돌아간다. 설치는 병동 단위로 반복 가능한 절차이고, 현장은 Intega와 함께 한다. 삼성벤처투자 등 기관투자자가 있고, 재무제표를 제출할 수 있다.
- **스크립트:** "The same system runs about 2,000 sensors across 14 facilities in Japan, more than these three hospitals combined. Installation is a repeatable ward routine delivered with Intega on the ground, and we're backed by institutional investors including Samsung Venture Investment — we're happy to provide financial statements."

### Q16. "계약 당사자는 유통사(DCH Auriga)인데, 문제가 생기면 MONIT에 직접 책임을 물을 수 있나?"
*Our contract is with the distributor. Who is accountable if something goes wrong?*

- **답변:** 제조사로서 MONIT가 기술·서비스 의무를 **백투백**으로 보증한다. 프로젝트 매니저는 MONIT 단일 창구다. 필요하면 ALPS 앞으로 제조사 보증서(letter of undertaking)를 별도로 제출하겠다.
- **스크립트:** "MONIT backs every technical and service obligation back-to-back as the manufacturer, and our project manager is your single point of contact. We'll provide a manufacturer's letter of undertaking to ALPS if that helps."
- **효과:** 계약 신뢰를 높이고, 동시에 MONIT가 조달 측과 직접 관계를 만드는 계기가 된다

### Q17. "레퍼런스가 2명, 6명, 30명 수준이다. 병원 규모 운영 실적이 있나?"
*Your Singapore references are two, six and thirty people. Where's a hospital-scale deployment?*

- **답변:** 싱가포르 병원 레퍼런스는 바로 **KTPH 자신**이다. 7주 동안 15명, 818건의 실측 데이터가 있다. 규모 실적은 일본(약 2,000개)과 한국 NHIS 프로그램이다. 이번 입찰은 KTPH 파일럿에서 나온 피드백을 반영한 버전이다.
- **스크립트:** "Our hospital reference in Singapore is KTPH itself — seven weeks, 818 measured readings, and this release incorporates your nurses' feedback. Scale comes from Japan, with about 2,000 sensors in daily use."

---

## 5. 일정·설치 (회의 두 번째 안건)

### Q18. "5개월 안에 3개 병원 가동이 정말 가능한가? 생산 리드타임은?"
*Is five months realistic, including manufacturing lead time?*

- **답변:** 일정의 핵심은 생산을 기다리지 않는 것이다. 사이트 조사가 끝나면(1~2개월 차) 릴레이 수량을 확정하고, 그동안 **파일럿 병동은 기존 재고로 시작**한다. 확대 설치는 양산 물량으로 3개월 차 이후 진행한다. 병동당 1~2개를 하루에 설치할 수 있다.
- **스크립트:** "We don't wait on production for the pilot — pilot wards start from existing stock while the survey fixes the final relay count, and rollout from month three uses the production batch at one to two wards a day."
- **확인 필요:** 파일럿 3개 병동분(릴레이 약 30대 + 센서) 재고 보유 여부를 대표가 확인할 것

### Q19. "우리 병원 Wi-Fi는 5GHz이고 보안 정책상 IoT 기기를 붙이기 어렵다."
*Our network is 5 GHz and our security policy restricts IoT devices.*

- **답변:** 릴레이는 2.4GHz만 지원한다. 대부분의 병원 무선 AP는 2.4/5GHz를 동시 지원하므로 **IoT 전용 2.4GHz SSID를 VLAN으로 분리**해 주면 된다. MAC 주소는 납품 전에 전달한다. 병원이 준비할 것은 19장 체크리스트 한 장이다.
- **스크립트:** "The relay needs a 2.4 GHz SSID; most hospital access points broadcast both bands, so a dedicated IoT SSID on its own VLAN is usually enough. We send the MAC addresses before delivery, and slide 19 lists everything IT needs to prepare."

### Q20. "결론적으로, MONIT를 선택해야 할 이유를 30초 안에 말해 보라."
*In thirty seconds: why MONIT?*

- **답변 (마무리 발언 겸용):**
  1. KTPH에서 이미 검증됐고, 그 피드백으로 개선한 시스템이다.
  2. 기저귀 계약을 바꿀 필요가 없다.
  3. 5년 비용이 오늘 확정된다.
  4. 싱가포르 데이터 보관, ISO 인증, 병원별 격리 환경을 갖췄다.
  5. 3개 병원 공통 구조라 ALPS 산하 다른 기관으로 그대로 확장할 수 있다.
- **스크립트:** "You've already tested it on your own ward, and this version carries your nurses' feedback. It works with the diapers you already buy, the five-year cost is fixed today, data stays in Singapore under ISO-certified controls, and it's built once for all three hospitals so any other ALPS institution can join on the same terms."

---

## 회의 직전 체크 (대표)

1. 유지보수비 하한선과 협상 권한 확정 (카드 ①·② 사용 여부)
2. 배터리 교체 비용이 유지보수비에 포함되는지
3. 기기별 비밀번호 펌웨어 납기 (외주사)
4. ISO 11월 취득 일정의 확실성 (인수 조건으로 걸어도 되는지)
5. 파일럿 3개 병동분 재고
6. 센서 소독 가이드 유무
