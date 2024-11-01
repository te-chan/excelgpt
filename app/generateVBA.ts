import {GenerationOptions} from './page';

interface APIRequest { 
    model?: string;
    messages: [{ role: string; content: string; }];
    temperature?: number;
    top_p?: number;
    max_tokens?: number;
}

function sanitizeAsVBAString(payload: string){
    return payload.replace(/"/g, '""');
}

export function generateOpenAI(options: GenerationOptions){
    const apiKey = sanitizeAsVBAString(options.apiKey);

    // construct ai messages
    const apiRequest: APIRequest = {
        model: "gpt-4o",
        messages: [{
            "role": "system",
            "content": options.systemPrompt
        }]
    };
    options.fewShots.forEach((shot) => {
        apiRequest.messages.push({
            role: "user",
            content: shot.human
        })
        apiRequest.messages.push({
            role: "assistant",
            content: shot.ai
        });
    })
    apiRequest.messages.push({
        role: "user",
        content: "<!-PROMPT_PLACEHOLDER-!>"
    });

    let apiPayload = sanitizeAsVBAString(JSON.stringify(apiRequest))
    apiPayload = apiPayload.replace("<!-PROMPT_PLACEHOLDER-!>", '" & prompt & "')

    return `
    
Function ${options.functionName}(prompt As String) As String
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
    apiKey = "${apiKey}"
    
    ' APIに送信するJSON形式のデータ
    data = "${apiPayload}"
    
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
    ${options.functionName} = json("choices")(1)("message")("content")
    
    ' オブジェクトを解放
    Set xmlHttp = Nothing
    Set json = Nothing
    Exit Function

ErrorHandler:
    ${options.functionName} = "Error: " & Err.Description
End Function


`
}

export function generateAzureOpenAI(options: GenerationOptions){
    const apiKey = sanitizeAsVBAString(options.apiKey);

    // construct ai messages
    const apiRequest: APIRequest = {
        temperature: 0,
        top_p: 0.95,
        max_tokens: 800,
        messages: [{
            "role": "system",
            "content": options.systemPrompt
        }]
    };
    options.fewShots.forEach((shot) => {
        apiRequest.messages.push({
            role: "user",
            content: shot.human
        })
        apiRequest.messages.push({
            role: "assistant",
            content: shot.ai
        });
    })
    apiRequest.messages.push({
        role: "user",
        content: "<!-PROMPT_PLACEHOLDER-!>"
    });

    let apiPayload = sanitizeAsVBAString(JSON.stringify(apiRequest))
    apiPayload = apiPayload.replace("<!-PROMPT_PLACEHOLDER-!>", '" & prompt & "')

    return `
    
Function ${options.functionName}(prompt As String) As String
    On Error GoTo ErrorHandler ' エラーハンドリングの設定
    Dim xmlHttp As Object
    Dim url As String
    Dim apiKey As String
    Dim data As String
    Dim response As String
    Dim json As Object
    
    ' OpenAI APIのエンドポイントURL
    url = "${options.azureEndpoint}/openai/deployments/${options.azureDeployment}/chat/completions?api-version=2024-02-15-preview"
    
    ' OpenAI APIキー（ここにあなたのAPIキーを設定）
    apiKey = "${apiKey}"
    
    ' APIに送信するJSON形式のデータ
    data = "${apiPayload}"
    
    ' XMLHTTPオブジェクトの作成
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' HTTP POSTリクエストを初期化
    xmlHttp.Open "POST", url, False
    
    ' リクエストヘッダーの設定（APIキーとContent-Type）
    xmlHttp.setRequestHeader "api-key", apiKey
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
    ${options.functionName} = json("choices")(1)("message")("content")
    
    ' オブジェクトを解放
    Set xmlHttp = Nothing
    Set json = Nothing
    Exit Function

ErrorHandler:
    ${options.functionName} = "Error: " & Err.Description
End Function


`
}