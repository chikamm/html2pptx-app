# html2pptx-app

HTMLスライドを、スクリーンショットではなく**PPTXネイティブの要素**（テキストボックス・図形・画像）に変換するWEBアプリ。

3段構えのハイブリッド構成で、どんなHTMLが来ても最善のルートで変換します。

```
アップロードされたHTML
        │
        ▼
① 既知テンプレート判定 (registry.js の fingerprint)
   一致 → 専用ジェネレータ（cheerioでHTMLを解析し、意味を理解した上で
           手作り品質のレイアウトを再現。最高精度・低コスト）
        │ 不一致
        ▼
② 汎用ルールベース変換 (extractDom.js + genericGenerator.js)
   Puppeteer(Chromium)で実レンダリングし、要素ごとの座標・スタイルを
   取得 → テキストボックス/図形/画像に機械的に変換。任意のHTMLに対応。
        │
        ▼
③ AIフォールバック (aiFallback.js)
   ②の結果で「重なり」「はみ出し」を自動検知したスライドだけ、
   Claude APIに座標の補正案を作らせて再描画。
   ANTHROPIC_API_KEY 未設定なら自動的にスキップ（②の結果をそのまま採用）。
```

## セットアップ

```bash
npm install
npm start        # http://localhost:8787
```

環境変数（すべて任意）:

| 変数 | 説明 |
|---|---|
| `PORT` | サーバーのポート（既定 8787） |
| `PUPPETEER_EXECUTABLE_PATH` | システムのChromiumパス。Dockerイメージでは自動設定済み |
| `ANTHROPIC_API_KEY` | 設定するとAIフォールバック層が有効になる |
| `ANTHROPIC_MODEL` | 既定 `claude-sonnet-4-5` |
| `ALLOW_GENERIC_FALLBACK` | `false` にすると、既知テンプレート内で未知のスライド型に遭遇した際に汎用エンジンへフォールバックしない（Chromium未使用の環境向け） |

## デプロイ（Docker）

```bash
docker build -t html2pptx-app .
docker run -p 8787:8787 -e ANTHROPIC_API_KEY=sk-... html2pptx-app
```

DockerfileはPuppeteer自身にChromiumをダウンロードさせず、`apt-get install chromium` で
システムのChromiumを使う構成にしています。オフライン/社内ネットワーク環境では
Puppeteerの既定ダウンロード先（`storage.googleapis.com`）がブロックされていることが多く、
この構成の方が確実にビルドできます（本プロジェクトの開発環境自体がそうでした）。

## テスト

Chromiumが使えない開発環境でも、パイプラインの各層を個別に検証できます。

```bash
npm run test:mock    # 汎用エンジン単体（手書きの抽出済みDOMモデルを使用、Chromium不要）
npm run test:sample  # 専用テンプレート単体（samples/fruits_frappe.html を実際にcheerioで解析）
```

実際のHTMLに対する ①→②（Puppeteer） の抽出まで含めたE2Eは、Chromiumが利用可能な
環境（Dockerコンテナ内など）で `POST /api/convert` を叩いて確認してください。

```bash
curl -X POST http://localhost:8787/api/convert \
  -F "file=@samples/fruits_frappe.html" \
  -o converted.pptx
```

## Microsoft Teamsに組み込む

`teams-app/` にTeamsアプリのパッケージ一式（`manifest.json` / `color.png` / `outline.png`）が入っています。
これをアップロードすると、Teamsの左側アプリ一覧やチャネルのタブとして、このWEBアプリをそのまま開けるようになります
（会話でHTMLを送ってやり取りするBotではなく、今のアップロード画面がTeamsの中に埋め込まれる形です）。

手順:

1. まずこのアプリをどこかにデプロイし、`https://xxxx.onrender.com` のような公開HTTPSのURLを用意する
   （Teamsのタブは`localhost`を読み込めないため必須）。
2. `teams-app/manifest.json` 内の `CHANGE-ME.example.com` をすべて実際のドメインに置き換える
   （`developer.websiteUrl` / `privacyUrl` / `termsOfUseUrl` / `staticTabs[0].contentUrl` /
   `staticTabs[0].websiteUrl` / `configurableTabs[0].configurationUrl` / `validDomains[0]` の7箇所）。
3. `teams-app/` フォルダの中身（`manifest.json`, `color.png`, `outline.png` の3ファイル、サブフォルダなし）
   をzipにまとめ直す。
4. Teamsで **アプリ → アプリを管理 → カスタムアプリをアップロード** から、作成したzipをアップロード
   （組織でカスタムアプリのアップロードが許可されていない場合は、情報システム部門にこのzipを渡して
   Teams管理センターから組織向けに公開してもらう）。
5. 個人アプリとして自分のTeamsに追加するか、共有したいチームのチャネルで「＋」→アプリ名で検索→
   タブとして追加。

`public/teams/config.html` はチャネル/グループチャット用タブの設定画面、`privacy.html` / `terms.html` は
manifestが要求する規約ページの最低限のひな形です。内容は実態に合わせて書き換えてください。

## 新しい「既知テンプレート」を追加する

自社で繰り返し使うHTMLテンプレートがあれば、`server/generator/templates/` に
新しいファイルを作り、`{ name, fingerprint(html), generate(html, options) }` を
`registerTemplate()` で登録するだけで①のルートに乗ります。
`famimaFrappeTemplate.js` が実装の参考になります（cheerioでの要素抽出、
共通ヘッダー/フッターのヘルパー、スライド種別ごとの生成関数、未知の
スライド型は汎用エンジンにフォールバックする仕組み、を一通り含んでいます）。

## 既知の制約 (v1)

- **グラデーション背景/塗り**: pptxgenjsはグラデーション塗りをサポートしないため、
  代表色1色に近似しています（`colorUtils.gradientToSolidHex`）。
- **インラインSVG**: 汎用エンジンはSVGを既定でスキップします（`options.rasterizeSvg`
  にラスタライズ関数を渡せば埋め込み可能）。
- **Webフォント**: PowerPoint側にインストールされていないフォントは意図しないものに
  置換されるため、`genericGenerator.mapFontFamily()` で安全なフォント群に正規化しています。
- **リッチテキストの粒度**: 汎用エンジンは `<b>/<strong>/<span>/<br>` 程度のインライン
  装飾は個別ランとして再現しますが、それ以上に複雑な入れ子装飾は簡略化されます。
- **アニメーション/インタラクション**: 静的なレイアウトのみが対象です。

## ディレクトリ構成

```
server/
  index.js                        Express本体・/api/convert
  extractor/extractDom.js         Puppeteerによる汎用DOM抽出（②の入力）
  generator/
    genericGenerator.js           抽出モデル→pptxgenjs（②本体、自動QAスコアも算出）
    aiFallback.js                 ③ Claude APIによる低信頼スライドの補正
    templates/
      registry.js                 既知テンプレートの登録・指紋照合
      famimaFrappeTemplate.js     ①の実装例（MaterialPR社内テンプレート）
  utils/
    colorUtils.js                 rgb()/gradient のパース・変換
    imageUtils.js                 <img src> → base64 data URI 解決
public/                           アップロードUI
samples/fruits_frappe.html        動作確認用サンプル（このプロジェクトの過去の変換対象）
test/
  run_mock_model.js                ②を手書きモデルで検証（Chromium不要）
  run_sample.js                    ①を実HTMLで検証（Chromium不要）
```
