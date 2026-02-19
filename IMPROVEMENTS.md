# 엑셀 파일 내보내기 개선사항 - 랜덤 서명 삽입

## 개요
엑셀 파일 내보내기 시 원본 양식을 유지하면서 서명을 무작위로 삽입하는 기능에 대한 검증 및 개선을 완료했습니다.

## 발견된 문제점

### 1. 동일 행에서 서명 중복 사용 문제 ⚠️
**문제**: 
- 한 사람의 행에 여러 개의 서명 placeholder(1, ○ 등)가 있을 때, 동일한 서명 variant가 여러 번 사용될 수 있었습니다.
- 예: 홍길동의 서명이 3개의 variant(v1, v2, v3)가 있는데, 3개의 placeholder 모두에 v1이 배치될 수 있었습니다.
- 이는 자연스럽지 않은 결과를 초래했습니다 (실제로는 같은 사람이 매번 조금씩 다르게 서명함).

**해결방법**:
```typescript
// 각 행마다 사용된 서명 variant 추적
const usedVariantsInRow = new Set<string>();

// 우선 사용하지 않은 variant 선택
const unusedSigs = availableSigs.filter(sig => !usedVariantsInRow.has(sig.variant));

if (unusedSigs.length > 0) {
  // 사용하지 않은 variant 우선 사용
  const randomSigIndex = Math.floor(Math.random() * unusedSigs.length);
  selectedSig = unusedSigs[randomSigIndex];
} else {
  // 모든 variant가 사용된 경우, 다시 처음부터
  usedVariantsInRow.clear();
  const randomSigIndex = Math.floor(Math.random() * availableSigs.length);
  selectedSig = availableSigs[randomSigIndex];
}
```

### 2. 랜덤 숫자 생성 코드 가독성 문제
**문제**: 
- `Math.floor(Math.random() * 11) - 5` 같은 코드가 반복되어 가독성이 떨어졌습니다.
- 주석이 있어도 실제 범위를 이해하기 어려웠습니다.

**해결방법**:
```typescript
// 새로운 헬퍼 함수 추가 (excelUtils.ts)
export const randomInt = (min: number, max: number): number => {
  return Math.floor(Math.random() * (max - min + 1)) + min;
};

// 사용 예:
const rotation = randomInt(-5, 5);    // -5부터 5까지
const offsetX = randomInt(-4, 4);     // -4부터 4까지
const offsetY = randomInt(-2, 2);     // -2부터 2까지
```

### 3. 엣지 케이스 처리 부족
**문제**: 
- 서명 variant가 유효하지 않은 경우에 대한 안전장치가 부족했습니다.
- 이론적으로는 발생하지 않지만, 방어 코드가 필요했습니다.

**해결방법**:
```typescript
// Safety check: ensure we have a valid signature
if (!selectedSig || !selectedSig.variant) {
  console.warn(`  [autoMatch] 경고: (${cell.row},${cell.col}) 유효하지 않은 서명 - 스킵`);
  continue;
}
```

## 개선 효과

### 1. 더 자연스러운 서명 배치
- **이전**: 동일한 서명이 여러 위치에 반복될 수 있었음
- **이후**: 각 행에서 가능한 한 다른 variant를 사용하여 자연스러운 변화를 제공

### 2. 코드 가독성 향상
- **이전**: `Math.floor(Math.random() * 11) - 5` (무엇을 하는지 파악하기 어려움)
- **이후**: `randomInt(-5, 5)` (명확하고 직관적)

### 3. 안정성 향상
- 예외 상황에 대한 방어 코드 추가
- 더 많은 로깅으로 디버깅 용이성 향상

## 기술 상세

### 수정된 파일
1. **services/excelUtils.ts**
   - `randomInt` 헬퍼 함수 추가
   
2. **services/excelService.ts**
   - `autoMatchSignatures` 함수 개선
   - 행 단위 서명 variant 추적 로직 추가
   - 안전성 검사 추가

### 랜덤 값 범위
현재 사용하는 랜덤 파라미터:
- **rotation**: -5도 ~ 5도 (자연스러운 기울기)
- **scale**: 0.95 ~ 1.1 (크기 변화)
- **offsetX**: -4px ~ 4px (수평 위치 변화)
- **offsetY**: -2px ~ 2px (수직 위치 변화)

이러한 범위는 실제 사람이 서명할 때의 자연스러운 변화를 반영합니다.

## 테스트 방법

### 테스트 케이스 1: 다중 placeholder
1. 엑셀 파일 생성:
   - 성명 열에 "홍길동"
   - 같은 행에 서명 placeholder 3개 (1, 1, 1)
   
2. 홍길동의 서명 파일 3개 업로드 (v1, v2, v3)

3. 처리 후 확인사항:
   - ✅ 3개 위치에 서로 다른 variant가 배치됨
   - ✅ 같은 variant가 반복 사용되지 않음 (충분한 variant가 있는 경우)

### 테스트 케이스 2: Variant 부족
1. 엑셀 파일 생성:
   - 성명 열에 "홍길동"
   - 같은 행에 서명 placeholder 5개
   
2. 홍길동의 서명 파일 2개만 업로드

3. 처리 후 확인사항:
   - ✅ 처음 2개는 다른 variant 사용
   - ✅ 3번째부터는 다시 처음부터 순환하여 사용
   - ✅ 오류 없이 정상 처리됨

### 테스트 케이스 3: 랜덤성 검증
1. 동일한 파일을 여러 번 처리

2. 확인사항:
   - ✅ 매번 다른 rotation, scale, offset 값이 적용됨
   - ✅ 서명이 조금씩 다른 위치와 각도로 배치됨

## 호환성

### 이전 버전과의 호환성
- ✅ 기존 엑셀 파일 형식 그대로 사용 가능
- ✅ 기존 서명 파일 형식 그대로 사용 가능
- ✅ 병합된 셀, 인쇄영역 등 모든 기존 기능 유지

### 알려진 제한사항
- Math.random() 사용: 암호학적으로 안전한 난수가 아님 (UI 용도로는 충분)
- 서명 variant가 1개뿐인 경우: 모든 위치에 동일한 서명 사용 (당연한 동작)

## 추가 개선 가능 사항

### 향후 고려사항
1. **더 복잡한 서명 선택 알고리즘**
   - 전체 문서 레벨에서 서명 분포 최적화
   - 연속된 행에서 동일 variant 회피

2. **커스터마이징 가능한 랜덤 범위**
   - 사용자가 rotation, scale, offset 범위 조정 가능
   - UI 설정 추가

3. **서명 품질 점수**
   - 각 서명에 품질 점수 부여
   - 중요한 위치에 더 좋은 품질의 서명 우선 배치

## 결론

이번 개선으로 엑셀 파일 내보내기 시 서명이 더욱 자연스럽고 다양하게 배치됩니다. 
코드의 가독성과 안정성도 향상되었으며, 기존 기능은 모두 유지됩니다.

## 참고 자료
- [TESTING_GUIDE.md](./TESTING_GUIDE.md) - 전체 테스트 가이드
- [CHANGELOG.md](./CHANGELOG.md) - 변경 이력
