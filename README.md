<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://github.com/user-attachments/assets/0aa67016-6eaf-458a-adb2-6e31a0763ed6" />
</div>

# Run and deploy your AI Studio app

This contains everything you need to run your app locally.

View your app in AI Studio: https://ai.studio/apps/drive/1xNmgL6kRc5Ii6yD6gh-c4i0xboY3pGhT

## Run Locally

**Prerequisites:**  Node.js


1. Install dependencies:
   `npm install`
2. Set the `GEMINI_API_KEY` in [.env.local](.env.local) to your Gemini API key
3. Run the app:
   `npm run dev`

## CI / Deployment

자동 빌드와 Vercel 배포 워크플로우가 추가되어 있습니다. 원격 리포지토리(`origin`)에 푸시하면 GitHub Actions가 자동으로 빌드를 실행합니다. Vercel로 자동 배포하려면 아래 추가 설정이 필요합니다.

1. GitHub 저장소 시크릿 설정
   - `VERCEL_TOKEN`: Vercel에서 생성한 배포 토큰
   - (선택) `VERCEL_ORG_ID`, `VERCEL_PROJECT_ID`

2. Vercel 연동 방법
   - Vercel 대시보드에서 프로젝트를 생성하거나 기존 프로젝트의 토큰을 발급받습니다.
   - 위 시크릿을 GitHub 리포지토리에 추가합니다.
   - `main` 브랜치에 푸시하면 `.github/workflows/deploy.yml`이 자동으로 실행되어 배포됩니다.

더 도와드릴까요? 제가 자동으로 커밋한 워크플로를 원격에 푸시해 드리려면 지금 푸시를 시도하겠습니다.

## 진단 스크립트

로컬 빌드 로그와 환경 정보를 자동으로 수집하는 스크립트가 추가되었습니다:

- `scripts/collect_build_info.ps1` — PowerShell 스크립트로 `build-diagnostics/` 폴더에 결과를 저장합니다.

실행 방법:
```powershell
cd 'C:\Users\user\OneDrive\Desktop\개발실\SMART-SIGN-PRO-main'
.\scripts\collect_build_info.ps1
```
스크립트 실행 후 `build-diagnostics\build-output.txt`와 `build-diagnostics\npm-install.txt` 파일을 여기 붙여주시면 제가 로그를 분석해 자동으로 수정하겠습니다.
