# Caddie｜経営者のためのWeb集客・AI対策診断

10問・100点満点のWeb集客診断ページです。点数の計算と結果表示はブラウザ内で行い、回答はGoogle Apps Script経由でスプレッドシートに保存します。

- `index.html` — 診断ページ本体（GitHub Pagesで公開）
- `gas/Code.gs` — 回答受け取り用のApps Script（使い方はファイル冒頭に記載）

回答の送信先は `index.html` の `ENDPOINT` に、GASのウェブアプリURLを設定します。
