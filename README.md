Excelの関数としてChatGPTなどのLLMを呼び出すための、VBAスクリプトジェネレータです。

## 使い方

[https://te-chan.github.io/excelgpt/](https://te-chan.github.io/excelgpt/) でアプリを使用できます。
OpenAIのAPIキー、関数名、システムプロンプトを入力し、生成された `{関数名}.bas`ファイルと [VBA JSON](https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1) を
モジュールとして追加することで、Excel内で関数として使用できます。

## 機能

 - [x] OpenAIのAPIキー、関数名、システムプロンプトを入力し、GPT関数を作成できます。
 - [ ] OpenAIのモデル名変更
 - [ ] 他のモデル追加
 - [ ] わかりやすいマニュアルの作成

## 開発と貢献

テストサーバーをコマンドから起動してください。:

```bash
npm run dev
# or
yarn dev
# or
pnpm dev
# or
bun dev
```

 [http://localhost:3000](http://localhost:3000) をブラウザから確認してください。

バグ修正や追加機能の要望などがあれば、IssueやPull-Requestを是非送ってください。

## ライセンス

MIT 