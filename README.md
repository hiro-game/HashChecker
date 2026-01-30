# HashChecker

PowerShell + WPF で動作する **ハッシュチェック用フロントエンドアプリ**です。  
ファイルの CRC32 / MD5 / SHA1 / SHA256 を高速に計算し、一覧表示・比較・CSV/TSV 出力が可能です。

---

## このアプリについて

本アプリは **Microsoft Copilot によって作成された Copilot 製アプリ**です。  
Windows 11 + PowerShell 7.5.4 で動作確認済みですが、PowerShell 5.1 でも動作します。

![HashChecker](https://github.com/user-attachments/assets/ "アプリウィンドウ")

---

## 必要要件

本アプリは **PowerShell と WPF のみで動作**します。  
追加のランタイムや外部 DLL のインストールは不要です。

- PowerShell 5.1 または PowerShell 7.x  
- Windows 標準の .NET Framework / .NET Runtime（追加インストール不要）  
- GUI（WPF）が利用可能な Windows 環境  

ハッシュ計算はすべてアプリ内部で完結しており、  
外部ツールや追加ライブラリは一切必要ありません。

---

## 使用方法と機能

### ● 基本操作
- ファイルまたはフォルダをウィンドウへドラッグ＆ドロップ  
- ハッシュ値を計算し、一覧に追加  
- 必要な列（サイズ / CRC32 / MD5 / SHA1 / SHA256）をチェックボックスで切り替え  
- 選択した行、または全行を CSV / TSV 形式で保存可能  
- 文字コードは以下から選択可能  
  - UTF-8（BOMあり / なし）  
  - Shift-JIS  

### ● 主な機能
- ハッシュ値の高速計算（複数ファイル対応）
- ファイルサイズの表示
- CSV / TSV 出力（文字コード選択可）
- 最前面表示の ON/OFF（ピンアイコン）
- UI はダークテーマベースで視認性良好
- ドラッグ＆ドロップ対応
- 選択行のコピー（右クリックメニュー）

---

## 特徴

- **PowerShell スクリプトとは思えない操作性**  
  WPF を使用した GUI により、一般的な Windows アプリと同じ感覚で操作できます。

- **外部依存がゼロ**  
  ハッシュ計算はすべて内部処理で完結。追加ツール不要。

- **高速で安定したハッシュ計算**  
  大量ファイルでもスムーズに処理。

- **柔軟な出力形式**  
  CSV / TSV、UTF-8 / Shift-JIS に対応。

- **Copilot と人間の共同開発による高品質コード**  
  読みやすさ・保守性を重視した構成。

---

## 動作要件

- Windows 10 / 11  
- PowerShell 5.1 または PowerShell 7.x  
- .NET Framework / .NET Runtime（Windows 標準でOK）  

---

## インストール方法

### 1. ZIP をダウンロードして展開
GitHub の「Code → Download ZIP」から取得し、任意のフォルダに展開します。

### 2. スクリプトを実行
- PowerShell 5.1 の場合：
```
powershell ./HashChecker.ps1
```
※PowerShell 7を使用する場合は「powershell」を「pwsh」に変更

### 3. ショートカットから起動したい場合

デスクトップにショートカットを作成し、リンク先を以下のようにします：
```
powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File .\HashChecker.ps1
```
※PowerShell 7を使用する場合は「powershell」を「pwsh」に変更


---

## ライセンス

本アプリは **MIT License** の下で公開されています。  
商用利用・改変・再配布すべて自由です。
---
MIT License

Copyright (c) ...

Permission is hereby granted, free of charge, to any person obtaining a copy


（全文は LICENSE ファイルを参照してください）

---
