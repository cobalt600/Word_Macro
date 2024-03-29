

# 論文執筆のときに変更箇所（追記箇所）を赤字にするマクロ
## 参考にしたもの
https://xp-song.github.io/posts/trackchange-word/

https://qiita.com/yuifu/items/821def235673068eb6ea

## 大きな流れ
Wordに備え付けられている「比較」という機能を使って、2つの文書の差分を「変更履歴」として記録したファイルを作成。この新しいファイルに対して「変更履歴を選択→選択箇所を赤字にする」というマクロを走らせる。最後に変更履歴を全部承認して完成。

## 使い方
1. 比較→2つの文書を比較
1. 比較の設定で書式設定と文字種の変換のチェックを外す。<strong>超重要。</strong>これをやらないと、インデントや余白の修正が全ページに適応されていたときにエラーになる
1. OKをおすと、比較箇所が変更履歴に放り込まれた新しい文書が作成される
1. 新しい文書でマクロを実行すると、変更箇所が赤字になる
1. それを保存して投稿（Endnoteのリンク切り忘れないようにね））

## 今後の方針
1. 比較するところも自動化したい
1. 書式の設定を見逃すようにすればもっと使い勝手が良いかもしれない。
1. 理想は2つのファイルを放り込むと、日付の新旧でどっちが校正版かを見分けて、自動で比較→赤字までやってほしい
1. VBAはWinかMacかに依存することがあるので、できれば全部Pythonで処理したいけど、ちょっとむずかしいかもしれない。
