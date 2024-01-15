function Invoke-AIGeneratedResponse {
    param (
        $ApiKey,
        $Question,
        $Model = "text-davinci-003",
        $Temperature = 1,
        $MaxTokens = 200
    )

    $RequestBody = @{
        prompt      = $Question
        model       = $Model
        temperature = $Temperature
        max_tokens  = $MaxTokens
    }

    $Header = @{ Authorization = "Bearer $ApiKey" }
    $RequestBody = $RequestBody | ConvertTo-Json

    $RestMethodParameter = @{
        Method      = 'Post'
        Uri         = 'https://api.openai.com/v1/completions'
        body        = $RequestBody
        Headers     = $Header
        ContentType = 'application/json'
    }

    try {
        $Response = (Invoke-RestMethod @RestMethodParameter).choices[0].text
    }
    catch {
        Write-Error "Failed to get AI-generated response: $_"
        return
    }

    return $Response
}

# Main script
$apiKey = "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"

do {
    $question = Read-Host "Enter a question (type 'exit' to quit)"

    if ($question -eq 'exit') {
        break
    }

    $generatedResponse = Invoke-AIGeneratedResponse -ApiKey $apiKey -Question $question

    if ($generatedResponse) {
        Write-Host "AI-generated response: $generatedResponse"
    } else {
        Write-Host "Failed to get AI-generated response."
    }
} while ($true)

cls
