param(
    [string]$ResourceGroup = "rg-Enterprise-IQ",
    [string]$Location = "eastus",
    [string]$NameSuffix = "d0b4af",
    [string]$SearchSku = "basic",
    [string]$SearchIndex = "workiq-docs",
    [string]$ChatDeployment = "gpt-4.1-mini",
    [string]$ChatModel = "gpt-4.1-mini",
    [string]$ChatModelVersion = "2025-04-14",
    [string]$EmbeddingDeployment = "text-embedding-ada-002",
    [string]$EmbeddingModel = "text-embedding-ada-002",
    [string]$EmbeddingModelVersion = "2",
    [switch]$SkipCurrentUserRoleAssignments
)

$ErrorActionPreference = "Stop"

$searchName = "srch-workiq-$NameSuffix"
$openAiName = "oai-workiq-$NameSuffix"

az group show --name $ResourceGroup --query "{name:name,location:location}" --output json | Out-Host

$existingSearch = az search service show `
    --resource-group $ResourceGroup `
    --name $searchName `
    --query name `
    --output tsv 2>$null

if (-not $existingSearch) {
    az search service create `
        --resource-group $ResourceGroup `
        --name $searchName `
        --location $Location `
        --sku $SearchSku `
        --partition-count 1 `
        --replica-count 1 `
        --semantic-search standard `
        --output json | Out-Host
}
else {
    Write-Host "Search service already exists: $searchName"
}

$existingOpenAi = az cognitiveservices account show `
    --resource-group $ResourceGroup `
    --name $openAiName `
    --query name `
    --output tsv 2>$null

if (-not $existingOpenAi) {
    az cognitiveservices account create `
        --resource-group $ResourceGroup `
        --name $openAiName `
        --location $Location `
        --kind OpenAI `
        --sku S0 `
        --custom-domain $openAiName `
        --yes `
        --output json | Out-Host
}
else {
    Write-Host "OpenAI account already exists: $openAiName"
}

$existingChatDeployment = az cognitiveservices account deployment show `
    --resource-group $ResourceGroup `
    --name $openAiName `
    --deployment-name $ChatDeployment `
    --query name `
    --output tsv 2>$null

if (-not $existingChatDeployment) {
    az cognitiveservices account deployment create `
        --resource-group $ResourceGroup `
        --name $openAiName `
        --deployment-name $ChatDeployment `
        --model-name $ChatModel `
        --model-version $ChatModelVersion `
        --model-format OpenAI `
        --sku-capacity 10 `
        --sku-name Standard `
        --output json | Out-Host
}
else {
    Write-Host "Chat deployment already exists: $ChatDeployment"
}

$existingEmbeddingDeployment = az cognitiveservices account deployment show `
    --resource-group $ResourceGroup `
    --name $openAiName `
    --deployment-name $EmbeddingDeployment `
    --query name `
    --output tsv 2>$null

if (-not $existingEmbeddingDeployment) {
    az cognitiveservices account deployment create `
        --resource-group $ResourceGroup `
        --name $openAiName `
        --deployment-name $EmbeddingDeployment `
        --model-name $EmbeddingModel `
        --model-version $EmbeddingModelVersion `
        --model-format OpenAI `
        --sku-capacity 10 `
        --sku-name Standard `
        --output json | Out-Host
}
else {
    Write-Host "Embedding deployment already exists: $EmbeddingDeployment"
}

if (-not $SkipCurrentUserRoleAssignments) {
    $assignee = az ad signed-in-user show --query id --output tsv
    $openAiScope = az cognitiveservices account show `
        --resource-group $ResourceGroup `
        --name $openAiName `
        --query id `
        --output tsv

    $openAiRoles = az role assignment list `
        --assignee $assignee `
        --scope $openAiScope `
        --query "[].roleDefinitionName" `
        --output tsv

    if ($openAiRoles -notcontains "Cognitive Services OpenAI User") {
        az role assignment create `
            --assignee $assignee `
            --role "Cognitive Services OpenAI User" `
            --scope $openAiScope `
            --output none
    }

    if ($openAiRoles -notcontains "Cognitive Services User") {
        az role assignment create `
            --assignee $assignee `
            --role "Cognitive Services User" `
            --scope $openAiScope `
            --output none
    }
}

$searchEndpoint = "https://$searchName.search.windows.net"
$openAiEndpoint = "https://$openAiName.openai.azure.com/openai/v1"

Write-Host ""
Write-Host "WorkIQ environment values:"
Write-Host "AZURE_SEARCH_ENDPOINT=$searchEndpoint"
Write-Host "AZURE_SEARCH_INDEX=$SearchIndex"
Write-Host "AZURE_OPENAI_ENDPOINT=$openAiEndpoint"
Write-Host "AZURE_AI_FOUNDRY_ENDPOINT=$openAiEndpoint"
Write-Host "AZURE_AI_FOUNDRY_DEPLOYMENT=$ChatDeployment"
Write-Host "AZURE_EMBEDDING_DEPLOYMENT=$EmbeddingDeployment"