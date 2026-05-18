# 技術経営専門職学位課程 修了要件チェッカー macOS用

このフォルダは、macOS上で修了要件チェッカーを実行・ビルドするための配布セットです。

## そのまま起動する

ターミナルでこのフォルダに移動し、以下を実行してください。

```bash
chmod +x run_mac.sh
./run_mac.sh
```

ブラウザでローカルURLが開きます。通常は `http://127.0.0.1:8765/` です。

## .appを作成する

Python 3 が入っているMacで、以下を実行してください。

```bash
chmod +x build_mac.sh
./build_mac.sh
```

成功すると、以下にmacOSアプリが作成されます。

```text
dist/tim-requirements-checker.app
```

## 対応入力

- `.xlsx`
- `.xlsm`
- 教務Webシステムの「成績閲覧：成績一覧（大学院）」PDF

## 注意

未署名アプリのため、初回起動時にmacOSのGatekeeper警告が出る場合があります。
その場合は Finder でアプリを右クリックし、「開く」を選択してください。
