/*
 * 제품 종류와 판정 규격 설정 파일
 * - 이 파일의 규격값은 시연용 예시값입니다. 실제 사내 규격값은 GitHub에 올리지 말고
 *   사내 PC에 있는 사본에서만 수정해서 사용하세요.
 * - 값을 바꾸면 화면, 엑셀 출력, 자동 테스트에 모두 반영됩니다.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.APP_CONFIG = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  return {
    // 고객사명 노출을 막기 위해 일반 명칭 사용. 사내 사본에서만 실제 고객사명으로 변경
    orderNoLabel: 'Customer Order No.',

    // 표지 오른쪽 "제품명 참고" 표
    cableTypes: [
      { id: 'Type 1', name: '2.9㎜ Riser Cable',     fiber: 'BIF, UBIF', color: 'WH', spec: '2.9mm' },
      { id: 'Type 2', name: '2.9㎜ Riser I/O Cable', fiber: 'BIF, UBIF', color: 'WH', spec: '2.9mm' },
      { id: 'Type 3', name: '2.9㎜ Riser I/O Cable', fiber: 'BIF, UBIF', color: 'BK', spec: '2.9mm' },
      { id: 'Type 4', name: '2.9㎜ Plenum Cable',    fiber: 'BIF, UBIF', color: 'WH', spec: '2.9mm' },
      { id: 'Type 5', name: '4.8㎜ Riser Cable',     fiber: 'UBIF',      color: 'WH', spec: '4.8mm' },
      { id: 'Type 6', name: '4.8㎜ Riser I/O Cable', fiber: 'UBIF',      color: 'WH', spec: '4.8mm' },
      { id: 'Type 7', name: '4.8㎜ Riser I/O Cable', fiber: 'UBIF',      color: 'BK', spec: '4.8mm' },
      { id: 'Type 8', name: '4.8㎜ Plenum Cable',    fiber: 'UBIF',      color: '-',  spec: '4.8mm' }
    ],

    // 색상 코드 → 보고서 표기
    colorNames: { WH: 'White', BK: 'Black', '-': '-' },

    // 판정 규격 (예시값). min/max 중 필요한 것만 지정
    specs: {
      '2.9mm': {
        cableOD:     { min: 2.80, max: 3.00, unit: 'mm', digits: 2 },
        tightBuffer: { min: 850,  max: 950,  unit: 'um' },
        att1310:     { max: 0.50, unit: 'dB/km', digits: 2 },
        att1550:     { max: 0.40, unit: 'dB/km', digits: 2 },
        marking:     { min: 990,  max: 1010, unit: 'mm' }
      },
      '4.8mm': {
        cableOD:     { min: 4.60, max: 5.00, unit: 'mm', digits: 2 },
        tightBuffer: { min: 850,  max: 950,  unit: 'um' },
        att1310:     { max: 0.50, unit: 'dB/km', digits: 2 },
        att1550:     { max: 0.40, unit: 'dB/km', digits: 2 },
        marking:     { min: 990,  max: 1010, unit: 'mm' }
      }
    },

    // marking 칸이 숫자(인쇄간격 mm)면 위 marking 범위로, 글자면 아래 값으로 판정
    markingPassValues: ['OK', 'GOOD', 'PASS', '양호', '합격']
  };
});
