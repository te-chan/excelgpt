'use client'

import { useState } from 'react'
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Textarea } from "@/components/ui/textarea"
import { Card, CardContent, CardFooter, CardHeader, CardTitle, CardDescription } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import { ClipboardCopy, Download, Plus, X, Key, Code, MessageSquare, Zap, FileCode } from 'lucide-react'
import { Alert, AlertDescription } from "@/components/ui/alert"
import { Separator } from "@/components/ui/separator"

interface FewShot {
  id: number;
  human: string;
  ai: string;
}

export default function Component() {
  const [apiKey, setApiKey] = useState('')
  const [functionName, setFunctionName] = useState('')
  const [systemPrompt, setSystemPrompt] = useState('')
  const [generatedScript, setGeneratedScript] = useState('')
  const [functionNameError, setFunctionNameError] = useState('')
  const [fewShots, setFewShots] = useState<FewShot[]>([])

  const validateFunctionName = (name: string) => {
    const regex = /^[a-zA-Z0-9_\s]+$/
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

  const addFewShot = () => {
    setFewShots([...fewShots, { id: Date.now(), human: '', ai: '' }])
  }

  const removeFewShot = (id: number) => {
    setFewShots(fewShots.filter(fs => fs.id !== id))
  }

  const updateFewShot = (id: number, field: 'human' | 'ai', value: string) => {
    setFewShots(fewShots.map(fs => fs.id === id ? { ...fs, [field]: value } : fs))
  }

  const generateVBAScript = () => {
    if (!validateFunctionName(functionName)) {
      return
    }

    const fewShotsString = fewShots.map(fs => 
      `    ' Human: ${fs.human.replace(/'/g, "''")}
    ' AI: ${fs.ai.replace(/'/g, "''")}`
    ).join('\n')

    const script = `Attribute VB_Name = "OpenAIModule"

Sub ${functionName.trim()}()
    ' VBAスクリプトの詳細は省略されています
    ' 以下の変数が使用されます：
    '   apiKey: ${apiKey}
    '   functionName: ${functionName}
    '   systemPrompt: ${systemPrompt}
    ' Few-shot examples:
${fewShotsString}
    ' OpenAI APIを使用してGPT-3.5-turboモデルにリクエストを送信します
End Sub
    `.trim()

    setGeneratedScript(script)
  }

  const copyToClipboard = () => {
    navigator.clipboard.writeText(generatedScript)
      .then(() => alert('スクリプトがクリップボードにコピーされました！'))
      .catch(err => console.error('コピーに失敗しました:', err))
  }

  const downloadVBAFile = () => {
    const blob = new Blob([generatedScript], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${functionName.trim() || 'OpenAIModule'}.bas`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  return (
    <Card className="mt-5 w-full max-w-2xl mx-auto bg-gradient-to-br from-gray-50 to-gray-100 dark:from-gray-900 dark:to-gray-800">
      <CardHeader className="space-y-1">
        <CardTitle className="text-2xl font-bold text-center">VBAスクリプトジェネレーター</CardTitle>
        <CardDescription className="text-center">OpenAI APIを使用したVBAスクリプトを簡単に生成</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="apiKey" className="flex items-center space-x-2 text-sm font-medium">
            <Key className="w-4 h-4" />
            <span>APIキー</span>
          </Label>
          <Input
            id="apiKey"
            value={apiKey}
            onChange={(e) => setApiKey(e.target.value)}
            placeholder="sk-..."
            className="bg-white dark:bg-gray-800"
          />
        </div>
        <div className="space-y-2">
          <Label htmlFor="functionName" className="flex items-center space-x-2 text-sm font-medium">
            <Code className="w-4 h-4" />
            <span>関数名</span>
          </Label>
          <Input
            id="functionName"
            value={functionName}
            onChange={handleFunctionNameChange}
            placeholder="CallOpenAI"
            aria-invalid={functionNameError ? 'true' : 'false'}
            aria-describedby={functionNameError ? 'functionName-error' : undefined}
            className="bg-white dark:bg-gray-800"
          />
          {functionNameError && (
            <Alert variant="destructive">
              <AlertDescription id="functionName-error">{functionNameError}</AlertDescription>
            </Alert>
          )}
        </div>
        <div className="space-y-2">
          <Label htmlFor="systemPrompt" className="flex items-center space-x-2 text-sm font-medium">
            <MessageSquare className="w-4 h-4" />
            <span>システムプロンプト</span>
          </Label>
          <Textarea
            id="systemPrompt"
            value={systemPrompt}
            onChange={(e) => setSystemPrompt(e.target.value)}
            placeholder="あなたは優秀なアシスタントです。"
            rows={3}
            className="bg-white dark:bg-gray-800"
          />
        </div>
        <Separator />
        <div className="space-y-4">
          <div className="flex justify-between items-center">
            <Label className="flex items-center space-x-2 text-sm font-medium">
              <Zap className="w-4 h-4" />
              <span>Few-shot examples</span>
            </Label>
            <Button onClick={addFewShot} variant="outline" size="sm" className="bg-white dark:bg-gray-800">
              <Plus className="mr-2 h-4 w-4" />
              FewShot追加
            </Button>
          </div>
          {fewShots.map((fs, index) => (
            <Card key={fs.id} className="p-4 bg-white dark:bg-gray-800 relative">
              <Button
                onClick={() => removeFewShot(fs.id)}
                variant="ghost"
                size="sm"
                className="absolute top-2 right-2"
                aria-label={`FewShot ${index + 1}を削除`}
              >
                <X className="h-4 w-4" />
              </Button>
              <div className="space-y-2">
                <Label htmlFor={`human-${fs.id}`} className="text-sm font-medium">Human</Label>
                <Textarea
                  id={`human-${fs.id}`}
                  value={fs.human}
                  onChange={(e) => updateFewShot(fs.id, 'human', e.target.value)}
                  placeholder="ユーザーの入力..."
                  rows={2}
                  className="bg-gray-50 dark:bg-gray-700"
                />
              </div>
              <div className="space-y-2 mt-2">
                <Label htmlFor={`ai-${fs.id}`} className="text-sm font-medium">AI</Label>
                <Textarea
                  id={`ai-${fs.id}`}
                  value={fs.ai}
                  onChange={(e) => updateFewShot(fs.id, 'ai', e.target.value)}
                  placeholder="AIの応答..."
                  rows={2}
                  className="bg-gray-50 dark:bg-gray-700"
                />
              </div>
            </Card>
          ))}
        </div>
        <Button onClick={generateVBAScript} className="w-full bg-blue-600 hover:bg-blue-700 text-white" disabled={!!functionNameError}>
          <FileCode className="mr-2 h-4 w-4" />
          スクリプト生成
        </Button>
        {generatedScript && (
          <div className="space-y-2">
            <Label htmlFor="generatedScript" className="flex items-center space-x-2 text-sm font-medium">
              <FileCode className="w-4 h-4" />
              <span>生成されたVBAスクリプト</span>
            </Label>
            <Textarea
              id="generatedScript"
              value={generatedScript}
              readOnly
              rows={10}
              className="font-mono text-sm bg-gray-50 dark:bg-gray-700"
            />
          </div>
        )}
      </CardContent>
      {generatedScript && (
        <CardFooter className="flex flex-col sm:flex-row gap-2">
          <Button onClick={copyToClipboard} className="w-full sm:w-1/2 bg-green-600 hover:bg-green-700 text-white">
            <ClipboardCopy className="mr-2 h-4 w-4" />
            クリップボードにコピー
          </Button>
          <Button onClick={downloadVBAFile} className="w-full sm:w-1/2 bg-purple-600 hover:bg-purple-700 text-white">
            <Download className="mr-2 h-4 w-4" />
            VBAファイルをダウンロード
          </Button>
        </CardFooter>
      )}
    </Card>
  )
}