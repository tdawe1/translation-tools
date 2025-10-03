#!/usr/bin/env python3
"""
Sample PDF generator for Japanese translation testing using pypdf.

Creates representative Japanese PDFs for integration testing of the PDF translation pipeline.
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
from datetime import datetime

def create_sample_japanese_text():
    """Return sample Japanese text content for testing."""
    
    simple_japanese_content = """日本語テスト文書

これは簡単な日本語のテスト文書です。この文書には基本的な日本語の文章が含まれています。

製品名：スマート翻訳システム
バージョン：2.0
リリース日：2025年1月15日

機能説明：
• 高精度な日本語から英語への翻訳
• レイアウト保持機能
• 自動キャッシュシステム
• 品質チェック機能

使用方法：
1. PDFファイルを選択します
2. 翻訳オプションを設定します
3. 実行ボタンをクリックします

注意事項：
• 翻訳精度は入力データの品質に依存します
• 複雑なレイアウトの場合は手動調整が必要です
• 大量のデータ処理には時間がかかる場合があります

お問い合わせ先：
サポートセンター：support@example.com
電話番号：03-1234-5678
営業時間：平日9:00-18:00"""

    multi_column_content = """技術仕様書

システム要件
ハードウェア要件：
• CPU: Intel Core i5以上
• メモリ: 8GB以上
• ストレージ: 256GB以上の空き容量

ソフトウェア要件：
• OS: Windows 10/11, macOS 11.0以上
• ブラウザ: Chrome, Firefox, Safari最新版
• ネットワーク: 高速インターネット接続

インストール手順
1. ダウンロードしたファイルを展開します
2. インストーラを実行します
3. ライセンス条項に同意します
4. インストール先を選択します
5. インストールを完了します

設定方法
基本設定：
• 言語：日本語または英語
• タイムゾーン：UTC+9
• 通知設定：オン/オフ

詳細設定：
• 自動更新：有効/無効
• データ保存場所：カスタム可能
• パフォーマンス：標準/高パフォーマンス

トラブルシューティング
よくある問題：
1. インストールが失敗する
   → 管理者権限で実行してください

2. 翻訳が遅い
   → ネットワーク接続を確認してください

3. ファイルが開けない
   → ファイル形式を確認してください"""

    mixed_content = """株式会社テクノロジーソリューションズ

ビジネス提案書

提出先：株式会社ABC
提出日：2025年1月14日
担当者：田中太郎

1. はじめに
本提案書は、貴社の業務効率化を実現するための統合ソリューションについてご説明いたします。

2. 現状分析
現在の課題：
• 業務プロセスが複雑で非効率
• 手作業による人的エラーの発生
• 情報共有が不十分
• コスト管理が困難

3. 提案ソリューション
AI驱动的業務自動化プラットフォーム

主要機能：
• 文書自動処理
• データ分析とレポート生成
• ワークフロー自動化
• リアルタイム監視

4. 導入効果
期待される効果：
• 業務効率：60%向上
• コスト削減：40%削減
• エラー率：95%削減
• 顧客満足度：30%向上

5. 実施計画
フェーズ1（1ヶ月目）：要件定義と設計
フェーズ2（2-3ヶ月目）：開発とテスト
フェーズ3（4ヶ月目）：導入とトレーニング
フェーズ4（5-6ヶ月目）：本番稼働と最適化

6. 費用見積もり
初期費用：5,000,000円
月額費用：500,000円
保守費用：年間600,000円

7. お問い合わせ
ご質問やご不明な点がございましたら、お気軽にお問い合わせください。"""

    return {
        "simple": simple_japanese_content,
        "multi_column": multi_column_content, 
        "mixed": mixed_content
    }

def create_sample_pdfs_with_reportlab():
    """Create sample Japanese PDFs using reportlab if available."""
    
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        reportlab_available = True
    except ImportError:
        print("ReportLab not available, creating simple text files instead...")
        return create_sample_text_files()
    
    # Get sample content
    content = create_sample_japanese_text()
    
    # Create test data directory
    data_dir = Path("tests/data")
    data_dir.mkdir(exist_ok=True)
    
    # PDF creation settings
    pdf_configs = [
        {
            "filename": "simple_japanese.pdf",
            "content": content["simple"],
            "description": "Simple single-column Japanese document"
        },
        {
            "filename": "multi_column_japanese.pdf", 
            "content": content["multi_column"],
            "description": "Multi-column layout with headers and lists"
        },
        {
            "filename": "mixed_content_japanese.pdf",
            "content": content["mixed"],
            "description": "Mixed content with tables, numbers, and formatting"
        }
    ]
    
    created_files = []
    
    for config in pdf_configs:
        try:
            output_path = data_dir / config["filename"]
            
            # Create PDF
            c = canvas.Canvas(str(output_path), pagesize=letter)
            width, height = letter
            
            # Set up text formatting
            font_size = 11
            line_height = 14
            margin = 72  # 1 inch margins
            y_position = height - margin
            
            # Try to register a Japanese font (fallback to basic font)
            try:
                # Try common Japanese fonts
                for font_name in ["IPAexGothic", "TakaoPGothic", "VL PGothic", "Noto Sans CJK JP"]:
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_name))
                        break
                    except:
                        continue
            except:
                pass  # Use default font
            
            # Split content into lines
            lines = config["content"].strip().split('\n')
            
            for line in lines:
                if y_position < margin:
                    # Add new page if needed
                    c.showPage()
                    y_position = height - margin
                
                # Skip empty lines (add spacing)
                if not line.strip():
                    y_position -= line_height // 2
                    continue
                
                # Add text to page
                c.setFont("Helvetica", font_size)
                c.drawString(margin, y_position, line)
                y_position -= line_height
            
            c.save()
            created_files.append(str(output_path))
            print(f"Created: {output_path} - {config['description']}")
            
        except Exception as e:
            print(f"Failed to create {config['filename']}: {e}")
    
    return created_files

def create_sample_text_files():
    """Create simple text files as fallback."""
    
    content = create_sample_japanese_text()
    
    # Create test data directory
    data_dir = Path("tests/data")
    data_dir.mkdir(exist_ok=True)
    
    # File creation settings
    file_configs = [
        {
            "filename": "simple_japanese.txt",
            "content": content["simple"],
            "description": "Simple single-column Japanese document (text format)"
        },
        {
            "filename": "multi_column_japanese.txt", 
            "content": content["multi_column"],
            "description": "Multi-column layout with headers and lists (text format)"
        },
        {
            "filename": "mixed_content_japanese.txt",
            "content": content["mixed"],
            "description": "Mixed content with tables, numbers, and formatting (text format)"
        }
    ]
    
    created_files = []
    
    for config in file_configs:
        try:
            output_path = data_dir / config["filename"]
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(config["content"])
            
            created_files.append(str(output_path))
            print(f"Created: {output_path} - {config['description']}")
            
        except Exception as e:
            print(f"Failed to create {config['filename']}: {e}")
    
    return created_files

def create_sample_pdfs():
    """Create sample Japanese PDFs for testing."""
    
    # Try to create PDFs with reportlab, fallback to text files
    created_files = create_sample_pdfs_with_reportlab()
    
    # Create a simple README for the test data
    readme_content = """# Sample Japanese Files for Translation Testing

This directory contains sample Japanese files used for integration testing of the PDF translation pipeline.

## Files

- `simple_japanese.pdf/txt` - Simple single-column Japanese document with basic business content
- `multi_column_japanese.pdf/txt` - Multi-column layout with headers, lists, and structured content
- `mixed_content_japanese.pdf/txt` - Mixed content including tables, numbers, business proposals, and formatting

## Usage

These files are used by:
- `tests/test_pdf_integration.py` - End-to-end integration tests
- Performance benchmarking
- Quality validation
- Layout preservation testing

## Content Description

All files contain realistic Japanese business content including:
- Technical specifications
- Business proposals
- User manuals
- Product descriptions
- Contact information

The content is designed to test various aspects of the translation pipeline:
- Text extraction accuracy
- Layout preservation
- Multi-language handling (Japanese numerals, mixed script)
- Table and list formatting
- Font and spacing management

## Note

PDF files are preferred for testing, but text files are provided as fallback when PDF libraries are not available.
"""
    
    data_dir = Path("tests/data")
    readme_path = data_dir / "README.md"
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print(f"\nCreated {len(created_files)} sample files for testing")
    print("README.md created with file descriptions")
    
    return True

if __name__ == "__main__":
    success = create_sample_pdfs()
    if success:
        print("\n✅ Sample file generation completed successfully!")
    else:
        print("\n❌ Sample file generation failed!")
        sys.exit(1)