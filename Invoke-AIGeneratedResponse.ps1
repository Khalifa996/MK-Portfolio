$apiKey = "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"
# Function to get the AI-generated response
function Get-AIGeneratedResponse {
    param (
        [string]$ApiKey,
        [string]$Question,
        [string]$Model = "text-davinci-003",
        [double]$Temperature = 0.7,
        [int]$MaxTokens = 70
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

while ($true) {
    $question = Read-Host "Enter a question (type 'exit' to quit)"
    
    if ($question -eq "exit") {
        break
    }

    $generatedResponse = Get-AIGeneratedResponse -ApiKey $apiKey -Question $question

    if ($generatedResponse) {
        Write-Host "AI-generated response: $generatedResponse"
    } else {
        Write-Host "Failed to get AI-generated response."
    }
}
