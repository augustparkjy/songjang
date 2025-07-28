#!/bin/bash

echo "🔨 엑셀 통합 프로그램 빌드 시작..."

# 기존 빌드 파일 정리
echo "🧹 기존 빌드 파일 정리 중..."
rm -rf dist build *.spec

# 단일 실행 파일 빌드
echo "📦 단일 실행 파일 빌드 중..."
uv run pyinstaller --onefile --name "엑셀통합프로그램" excel_merger.py

# 앱 번들 빌드
echo "📱 macOS 앱 번들 빌드 중..."
uv run pyinstaller --onedir --windowed --name "엑셀통합프로그램_앱" excel_merger.py

# 코드 서명 제거 (파일 손상 방지)
echo "🔓 코드 서명 제거 중..."
if [ -f "dist/엑셀통합프로그램" ]; then
    xattr -cr dist/엑셀통합프로그램
    echo "✅ 단일 실행 파일 코드 서명 제거 완료"
fi

if [ -f "dist/엑셀통합프로그램_앱.app/Contents/MacOS/엑셀통합프로그램_앱" ]; then
    xattr -cr dist/엑셀통합프로그램_앱.app/Contents/MacOS/엑셀통합프로그램_앱
    echo "✅ 앱 번들 코드 서명 제거 완료"
fi

# 압축 파일 생성
echo "📦 배포용 압축 파일 생성 중..."
cd dist

# 단일 실행 파일 압축
if [ -f "엑셀통합프로그램" ]; then
    zip -r ../엑셀통합프로그램_단일파일.zip 엑셀통합프로그램
    echo "✅ 단일 실행 파일 압축 완료"
fi

# 앱 번들 압축
if [ -d "엑셀통합프로그램_앱.app" ]; then
    zip -r ../엑셀통합프로그램_앱.zip 엑셀통합프로그램_앱.app
    echo "✅ 앱 번들 압축 완료"
fi

cd ..

# 파일 크기 확인
echo "📊 빌드 결과:"
ls -lh dist/
echo ""
echo "📦 압축 파일:"
ls -lh *.zip

echo ""
echo "🎉 빌드 완료!"
echo ""
echo "📋 사용 방법:"
echo "1. 단일 실행 파일: ./dist/엑셀통합프로그램"
echo "2. macOS 앱: ./dist/엑셀통합프로그램_앱.app"
echo "3. 압축 파일: 엑셀통합프로그램_단일파일.zip, 엑셀통합프로그램_앱.zip"
echo ""
echo "⚠️  압축 해제 후 실행 시 '손상된 파일' 오류가 발생하면:"
echo "   xattr -cr /path/to/엑셀통합프로그램" 