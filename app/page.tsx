'use client'

import { useState } from 'react'
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Textarea } from "@/components/ui/textarea"
import { Card, CardContent, CardFooter, CardHeader, CardTitle } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { ClipboardCopy, Download } from 'lucide-react'

function sanitizeVBA(input: string): string {
  // ダブルクォートを2つのダブルクォートに置き換える
  input = input.replace(/"/g, '""');
  // 改行文字を '" & vbCrLf & "' に置き換える
  input = input.replace(/\r\n|\n|\r/g, '" & vbCrLf & "');
  return input;
}

export default function Component() {
  const [apiKey, setApiKey] = useState('')
  const [functionName, setFunctionName] = useState('')
  const [systemPrompt, setSystemPrompt] = useState('')
  const [generatedScript, setGeneratedScript] = useState('')
  const [functionNameError, setFunctionNameError] = useState('')

  const generateVBAScript = () => {
    const sanitizedFunctionName = sanitizeVBA(functionName);
    const sanitizedSystemPrompt = sanitizeVBA(systemPrompt);
    const sanitizedApiKey = sanitizeVBA(apiKey);

    const script = `

Function ${sanitizedFunctionName}(prompt As String) As String
    On Error GoTo ErrorHandler ' エラーハンドリングの設定
    Dim xmlHttp As Object
    Dim url As String
    Dim apiKey As String
    Dim data As String
    Dim response As String
    Dim json As Object
    
    ' OpenAI APIのエンドポイントURL
    url = "https://api.openai.com/v1/chat/completions"
    
    ' OpenAI APIキー（ここにあなたのAPIキーを設定）
    apiKey = "${sanitizedApiKey}"

    ' プロンプト内の特殊文字をエスケープ
    prompt = Replace(prompt, "\", "\\")
    prompt = Replace(prompt, """", "\""")
    
    ' APIに送信するJSON形式のデータ
    data = "{""model"":""gpt-4"",""messages"":[{""role"":""system"",""content"":""${sanitizedSystemPrompt}""},{""role"":""user"",""content"":""" & prompt & """}]}"
    
    ' XMLHTTPオブジェクトの作成
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' HTTP POSTリクエストを初期化
    xmlHttp.Open "POST", url, False
    
    ' リクエストヘッダーの設定（APIキーとContent-Type）
    xmlHttp.setRequestHeader "Authorization", "Bearer " & apiKey
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    
    ' リクエストを送信（データも送信）
    xmlHttp.Send data
    
    ' レスポンスを変数に保存
    response = xmlHttp.ResponseText
    
    ' 即時ウィンドウにレスポンスを出力（デバッグ用）
    Debug.Print "API Response: " & response
    
    ' JSONレスポンスを解析するために、JSONオブジェクトを取得
    Set json = JsonConverter.ParseJson(response)
    
    ' 即時ウィンドウに解析後のJSONを出力（デバッグ用）
    Debug.Print "Parsed JSON: " & json("choices")(1)("message")("content")
    
    ' レスポンスから生成されたテキストを取得
    ChatCompletion = json("choices")(1)("message")("content")
    
    ' オブジェクトを解放
    Set xmlHttp = Nothing
    Set json = Nothing
    Exit Function

ErrorHandler:
    ChatCompletion = "Error: " & Err.Description
End Function


    `.trim()

    setGeneratedScript(script)
  }

  const copyToClipboard = () => {
    navigator.clipboard.writeText(generatedScript)
      .then(() => alert('スクリプトがクリップボードにコピーされました！'))
      .catch(err => console.error('コピーに失敗しました:', err))
  }

  const downloadVBAFile = async () => {
    // Download the BAS file from the URL
    const basFileUrl = 'https://raw.githubusercontent.com/VBA-tools/VBA-JSON/refs/heads/master/JsonConverter.bas';
    const basFileResponse = await fetch(basFileUrl);
    const basFileBlob = await basFileResponse.blob();
  
    // Create a Blob for the generated script
    const scriptBlob = new Blob([generatedScript], { type: 'text/plain' });
  
    // Function to trigger download
    const triggerDownload = (blob: Blob, filename: string) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    };
  
    // Trigger downloads for both files
    triggerDownload(scriptBlob, `${functionName || 'OpenAIModule'}.bas`);
    triggerDownload(basFileBlob, 'JsonConverter.bas');
  };

  const validateFunctionName = (name: string) => {
    const regex = /^[a-zA-Z0-9_]+$/
    if (!regex.test(name)) {
      setFunctionNameError('関数名には英数字、アンダースコア、スペースのみ使用できます。')
      return false
    }
    setFunctionNameError('')
    return true
  }

  const handleFunctionNameChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newName = e.target.value
    setFunctionName(newName)
    validateFunctionName(newName)
  }



  return (
    <Card className="w-full max-w-2xl mx-auto">
      <CardHeader>
        <CardTitle>VBAスクリプトジェネレーター</CardTitle>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="space-y-2">
          <Label htmlFor="apiKey">APIキー</Label>
          <Input
            id="apiKey"
            value={apiKey}
            onChange={(e) => setApiKey(e.target.value)}
            placeholder="sk-..."
          />
        </div>
        <div className="space-y-2">
          <Label htmlFor="functionName">関数名</Label>
          <Input
            id="functionName"
            value={functionName}
            onChange={handleFunctionNameChange}
            placeholder="CallOpenAI"
            aria-invalid={functionNameError ? 'true' : 'false'}
            aria-describedby={functionNameError ? 'functionName-error' : undefined}
          />
          {functionNameError && (
            <Alert variant="destructive">
              <AlertDescription id="functionName-error">{functionNameError}</AlertDescription>
            </Alert>
          )}
        </div>
        <div className="space-y-2">
          <Label htmlFor="systemPrompt">システムプロンプト</Label>
          <Textarea
            id="systemPrompt"
            value={systemPrompt}
            onChange={(e) => setSystemPrompt(e.target.value)}
            placeholder="あなたは優秀なアシスタントです。"
            rows={3}
          />
        </div>
        <Button onClick={generateVBAScript} className="w-full" disabled={!!functionNameError}>
          スクリプト生成
        </Button>
        {generatedScript && (
          <div className="space-y-2">
            <Label htmlFor="generatedScript">生成されたVBAスクリプト</Label>
            <Textarea
              id="generatedScript"
              value={generatedScript}
              readOnly
              rows={10}
              className="font-mono text-sm"
            />
          </div>
        )}
      </CardContent>
      {generatedScript && (
        <CardFooter className="flex flex-col sm:flex-row gap-2">
          <Button onClick={copyToClipboard} className="w-full sm:w-1/2">
            <ClipboardCopy className="mr-2 h-4 w-4" />
            クリップボードにコピー
          </Button>
          <Button onClick={downloadVBAFile} className="w-full sm:w-1/2">
            <Download className="mr-2 h-4 w-4" />
            VBAファイルをダウンロード
          </Button>
        </CardFooter>
      )}
    </Card>
  )
}