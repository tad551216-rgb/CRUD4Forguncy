# CAS CRUD一覧ジェネレーター（PWA）

Forguncyの `.fgcp` を渡すと、**ページ別CRUD一覧Excel**（4シート・色分け）を生成してダウンロードできるブラウザアプリです。
`generate_crud.py` のロジックをそのままJSへ移植しています。

## 最大の特長：端末内で完結
- `.fgcp`（CASプロジェクト一式）もForguncyドキュメントも、**どこにも送信・保存しません**。
- 解凍（JSZip）・CRUD抽出・Excel生成（ExcelJS）は、すべてブラウザの中だけで実行されます。
- サーバー不要。オフラインでも動作します（PWA）。

## 使い方
1. このページをブラウザで開く。
2. **①** に `.fgcp` をドロップ（必須）。コピーを推奨（元ファイルは変更しません）。
3. **②** に Forguncyの「ドキュメント生成」Excel をドロップ（任意）。付けるとページにNo.が入り、その順で並びます。省略すると Pages フォルダ順。
4. 「CRUD一覧を生成」を押す。
5. 完了したら「Excelをダウンロード」。`CAS_ページ別CRUD一覧_YYYYMMDD.xlsx` が保存されます。

## 出力シート
| シート | 内容 |
|---|---|
| 統計サマリ | 全体統計＋業務領域別ページ数 |
| CRUDマトリクス | 全ページのC/R/U/D・業務領域・注記（C緑・U黄・D橙で色分け、オートフィルタ・行固定付き） |
| テーブル別逆引き | テーブル→操作ページの逆引き |
| 削除ページ(要注意) | D操作を行うページのみ抽出 |

## GitHub Pages への配置
1. リポジトリ（例：`cas-crud`）を作り、**この `cas-crud-pwa/` の中身を直下に**置く。
2. push 後、Settings → Pages → Branch を `main` / `(root)` に設定。
3. `https://<ユーザー名>.github.io/cas-crud/` で公開。HTTPSなのでPWA（インストール・オフライン）が有効になります。

> パスは相対参照（`./`）なので、project page（サブパス配信）でもそのまま動きます。

## ローカルで確認する
Service Worker は `localhost` か HTTPS でのみ動きます。

```bash
cd cas-crud-pwa
python3 -m http.server 8000
# ブラウザで http://localhost:8000/
```

## オフライン／ホーム画面に追加
初回読み込み後、ライブラリ（JSZip/ExcelJS）も含めてキャッシュされ、オフラインで起動できます。
ブラウザの「ホーム画面に追加」「アプリをインストール」でアプリとして使えます。

## ファイル構成
```
cas-crud-pwa/
├── index.html               UI
├── app.js                   UIグルー（ファイル選択・進捗・DL）
├── crud-core.js             ★中核ロジック（generate_crud.py の移植）
├── manifest.webmanifest     PWA設定
├── sw.js                    Service Worker（オフライン）
├── vendor/
│   ├── jszip.min.js         3.10.1（.fgcp/.xlsx 解凍）
│   └── exceljs.min.js       4.4.0（Excel生成）
├── icons/                   アプリアイコン（192/512/maskable）
├── make_icons.py            アイコン再生成（任意）
└── test.js                  コア検証テスト（Node、任意）
```

## 更新したいとき
`crud-core.js` などを変更したら、`sw.js` の `CACHE = 'cas-crud-v1'` の番号を上げてください（`v2` など）。
古いキャッシュが破棄され、新しい版が反映されます。

## ロジックの限界（`generate_crud.py` と同じ）
保守で「漏れ」になりうる箇所。値が想定と違うときはここを疑ってください。
1. **サーバーサイドコマンド（SSC）のDB書込みは自動検出されない。** `crud-core.js` の `SSC_NOTES` にハードコードした注記のみ。SSCを追加したページは手で追記が必要。
2. **`UpdateType` 未指定は「更新(U)」扱い。** 実データでadd相当がType省略されていると C が U に化けます。最初の出力でCが少なければここを確認。
3. **コマンドを日本語化→正規表現で抽出**する方式のため、テーブル名に `]` が含まれると壊れます（CASでは無さそうですが念のため）。
4. **ドロップダウンの DataSource を R に加算**。テーブル名でなくクエリ名/数式が入ると、Rに軽くノイズが乗ることがあります。

## テスト（任意）
```bash
npm install jszip exceljs   # 開発時のみ
node test.js                # 合成.fgcpでコアを検証
```
