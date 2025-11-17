# Azure DevOps Webhook Creation Script - Enhanced Version
# This script creates Jellyfish webhooks with a custom authorization header
# with flexible project selection options (select specific, exclude by name, or all).

# --- UPDATE THESE VARIABLES ---
$Organization = "your-organization" # UPDATE with your Azure DevOps organization name
$PersonalAccessToken = "PAT TOKEN"         # UPDATE with your ADO Personal Access Token
$jellyfishUrl = "https://app.jellyfish.co/ingest-webhooks/ado/" # UPDATE with your Jellyfish endpoint
$jellyfishBearerToken = "abc123"           # <-- UPDATE with your Jellyfish Bearer Token
# ------------------------------

# Define the event types you want to create webhooks for.
$eventTypesToCreate = @(
    "workitem.created",
    "workitem.deleted",
    "workitem.restored",
    "workitem.updated"
)

Write-Host "=== Azure DevOps Webhook Creation Script - Enhanced ===" -ForegroundColor Cyan
Write-Host "Organization: $Organization"
Write-Host "Target URL: $jellyfishUrl"
Write-Host "Will create webhooks for: $($eventTypesToCreate -join ', ')"
Write-Host ""

# Function to get all projects
function Get-AllProjects {
    param(
        [string]$OrgName,
        [string]$PAT
    )
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"))
    $headers = @{
        'Authorization' = "Basic $base64AuthInfo"
        'Content-Type' = 'application/json'
    }
    $uri = "https://dev.azure.com/$OrgName/_apis/projects?api-version=7.1-preview.4"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
        return $response.value
    }
    catch {
        Write-Error "Failed to get projects: $($_.Exception.Message)"
        return $null
    }
}

# Function to list existing webhooks (to check for duplicates)
function Get-WebhookSubscriptions {
    param(
        [string]$OrgName,
        [string]$PAT
    )
    
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"))
    $headers = @{
        'Authorization' = "Basic $base64AuthInfo"
        'Content-Type' = 'application/json'
    }
    
    $uri = "https://dev.azure.com/$OrgName/_apis/hooks/subscriptions?api-version=7.1-preview.1"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
        return $response.value
    }
    catch {
        Write-Error "Failed to get webhooks: $($_.Exception.Message)"
        return $null
    }
}

# Function to display projects and allow user selection (inclusion or exclusion)
function Select-Projects {
    param(
        [array]$AllProjects
    )
    
    Write-Host "=== Project Selection ===" -ForegroundColor Cyan
    Write-Host "Choose how you want to select projects:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Select specific projects (by number)" -ForegroundColor White
    Write-Host "  [2] Exclude specific projects (by name)" -ForegroundColor White
    Write-Host "  [3] Select all projects" -ForegroundColor White
    Write-Host "  [Enter] Cancel" -ForegroundColor Gray
    Write-Host ""
    
    do {
        $modeInput = Read-Host "Choose selection mode (1-3)"
        
        if ([string]::IsNullOrWhiteSpace($modeInput)) {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return @()
        }
        
        switch ($modeInput) {
            "1" {
                return Select-ProjectsByNumber -AllProjects $AllProjects
            }
            "2" {
                return Select-ProjectsByExclusion -AllProjects $AllProjects
            }
            "3" {
                Write-Host "✓ Selected all $($AllProjects.Count) projects" -ForegroundColor Green
                return $AllProjects
            }
            default {
                Write-Host "Invalid selection. Please enter 1, 2, or 3." -ForegroundColor Red
            }
        }
    } while ($true)
}

# Function to select projects by number (original functionality)
function Select-ProjectsByNumber {
    param(
        [array]$AllProjects
    )
    
    Write-Host ""
    Write-Host "=== Select Projects by Number ===" -ForegroundColor Cyan
    Write-Host "Available projects in organization:" -ForegroundColor Yellow
    Write-Host ""
    
    # Display all projects with numbers
    for ($i = 0; $i -lt $AllProjects.Count; $i++) {
        Write-Host "  [$($i + 1)] $($AllProjects[$i].name)" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "Selection options:" -ForegroundColor Yellow
    Write-Host "  • Enter project numbers (e.g., 1,3,5 or 1-3 or 1 3 5)" -ForegroundColor Gray
    Write-Host "  • Press Enter to go back" -ForegroundColor Gray
    Write-Host ""
    
    do {
        $userInput = Read-Host "Select projects"
        
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            return @()
        }
        
        # Parse user input (supports comma-separated, space-separated, and ranges)
        $selectedProjects = @()
        $validSelection = $true
        
        # Replace common separators and split
        $numbers = $userInput -replace '[,\s]+', ' ' -split ' ' | Where-Object { $_ -ne '' }
        
        foreach ($number in $numbers) {
            if ($number -match '(\d+)-(\d+)') {
                # Handle ranges like "1-3"
                $start = [int]$matches[1]
                $end = [int]$matches[2]
                for ($i = $start; $i -le $end; $i++) {
                    if ($i -ge 1 -and $i -le $AllProjects.Count) {
                        $selectedProjects += $AllProjects[$i - 1]
                    } else {
                        Write-Host "Invalid project number: $i (valid range: 1-$($AllProjects.Count))" -ForegroundColor Red
                        $validSelection = $false
                        break
                    }
                }
            } elseif ($number -match '^\d+$') {
                # Handle individual numbers
                $projectIndex = [int]$number
                if ($projectIndex -ge 1 -and $projectIndex -le $AllProjects.Count) {
                    $selectedProjects += $AllProjects[$projectIndex - 1]
                } else {
                    Write-Host "Invalid project number: $projectIndex (valid range: 1-$($AllProjects.Count))" -ForegroundColor Red
                    $validSelection = $false
                    break
                }
            } else {
                Write-Host "Invalid input format: $number" -ForegroundColor Red
                $validSelection = $false
                break
            }
        }
        
        if ($validSelection -and $selectedProjects.Count -gt 0) {
            # Remove duplicates
            $selectedProjects = $selectedProjects | Sort-Object name -Unique
            
            Write-Host ""
            Write-Host "✓ Selected $($selectedProjects.Count) project(s):" -ForegroundColor Green
            foreach ($project in $selectedProjects) {
                Write-Host "  • $($project.name)" -ForegroundColor White
            }
            
            Write-Host ""
            $confirm = Read-Host "Proceed with these projects? (y/N)"
            if ($confirm.ToLower() -eq 'y' -or $confirm.ToLower() -eq 'yes') {
                return $selectedProjects
            }
            Write-Host "Please make a new selection." -ForegroundColor Yellow
        } elseif ($validSelection) {
            Write-Host "No projects selected. Please try again." -ForegroundColor Red
        }
        
        Write-Host ""
    } while ($true)
}

# Function to select projects by excluding specific ones by name
function Select-ProjectsByExclusion {
    param(
        [array]$AllProjects
    )
    
    Write-Host ""
    Write-Host "=== Exclude Projects by Name ===" -ForegroundColor Cyan
    Write-Host "Available projects in organization ($($AllProjects.Count) total):" -ForegroundColor Yellow
    Write-Host ""
    
    # Display all projects (sorted for easier reading)
    $sortedProjects = $AllProjects | Sort-Object name
    foreach ($project in $sortedProjects) {
        Write-Host "  • $($project.name)" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "Exclusion options:" -ForegroundColor Yellow
    Write-Host "  • Enter project names to exclude (supports partial names)" -ForegroundColor Gray
    Write-Host "  • Separate multiple names with commas (e.g., 'test,demo,staging')" -ForegroundColor Gray
    Write-Host "  • Matching is case-insensitive" -ForegroundColor Gray
    Write-Host "  • Press Enter to go back" -ForegroundColor Gray
    Write-Host ""
    
    do {
        $userInput = Read-Host "Enter project names to exclude"
        
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            return @()
        }
        
        # Parse exclusion input
        $exclusionTerms = $userInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        $excludedProjects = @()
        $selectedProjects = @()
        
        # Find projects to exclude based on name matching
        foreach ($term in $exclusionTerms) {
            $matchingProjects = $AllProjects | Where-Object { $_.name -like "*$term*" }
            if ($matchingProjects) {
                $excludedProjects += $matchingProjects
                Write-Host "  Found matches for '$term':" -ForegroundColor Yellow
                foreach ($match in $matchingProjects) {
                    Write-Host "    - $($match.name)" -ForegroundColor Gray
                }
            } else {
                Write-Host "  No matches found for '$term'" -ForegroundColor Red
            }
        }
        
        if ($excludedProjects.Count -eq 0) {
            Write-Host "No projects matched the exclusion terms. Please try again." -ForegroundColor Red
            continue
        }
        
        # Remove duplicates from excluded projects
        $excludedProjects = $excludedProjects | Sort-Object name -Unique
        
        # Get the remaining projects (all projects minus excluded ones)
        $selectedProjects = $AllProjects | Where-Object { 
            $currentProject = $_
            -not ($excludedProjects | Where-Object { $_.id -eq $currentProject.id })
        }
        
        Write-Host ""
        Write-Host "✓ Excluding $($excludedProjects.Count) project(s):" -ForegroundColor Red
        foreach ($project in $excludedProjects | Sort-Object name) {
            Write-Host "  ✗ $($project.name)" -ForegroundColor Red
        }
        
        Write-Host ""
        Write-Host "✓ Will process $($selectedProjects.Count) remaining project(s):" -ForegroundColor Green
        foreach ($project in $selectedProjects | Sort-Object name) {
            Write-Host "  • $($project.name)" -ForegroundColor White
        }
        
        if ($selectedProjects.Count -eq 0) {
            Write-Host ""
            Write-Host "No projects would be processed after exclusions. Please adjust your exclusion criteria." -ForegroundColor Red
            continue
        }
        
        Write-Host ""
        $confirm = Read-Host "Proceed with these $($selectedProjects.Count) projects? (y/N)"
        if ($confirm.ToLower() -eq 'y' -or $confirm.ToLower() -eq 'yes') {
            return $selectedProjects
        }
        Write-Host "Please adjust your exclusion criteria." -ForegroundColor Yellow
        Write-Host ""
    } while ($true)
}

# Function to create a new webhook subscription
function Create-WebhookSubscription {
    param(
        [string]$OrgName,
        [string]$PAT,
        [string]$ProjectID,
        [string]$EventType,
        [string]$JellyfishUrl,
        [string]$BearerToken
    )
    
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"))
    $headers = @{
        'Authorization' = "Basic $base64AuthInfo"
        'Content-Type' = 'application/json'
    }
    
    # Define the payload for the new webhook, now including the httpHeaders
    $body = @{
        publisherId      = "tfs"
        eventType        = $EventType
        consumerId       = "webHooks"
        consumerActionId = "httpRequest"
        publisherInputs  = @{
            projectId = $ProjectID
        }
        consumerInputs   = @{
            url                   = $JellyfishUrl
            httpHeaders           = "Authorization: Bearer $BearerToken" # Custom header
            resourceDetailsToSend = "all"
            messagesToSend        = "none"
            detailedMessagesToSend= "none"
        }
    } | ConvertTo-Json -Depth 5 # Use -Depth to ensure nested objects are converted
    
    $uri = "https://dev.azure.com/$OrgName/_apis/hooks/subscriptions?api-version=7.1-preview.1"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body
        Write-Host "   ✓ Successfully created webhook for '$EventType' (ID: $($response.id))" -ForegroundColor Green
        return $response
    }
    catch {
        Write-Error "   ❌ Failed to create webhook for '$EventType': $($_.Exception.Message)"
        if ($_.Exception.Response) {
            Write-Host "   HTTP Status: $($_.Exception.Response.StatusCode)" -ForegroundColor Red
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $errorBody = $reader.ReadToEnd();
            Write-Host "   Error Details: $errorBody" -ForegroundColor Red
        }
        return $null
    }
}

# --- Main Script Execution ---

# 1. Get all projects
Write-Host "Getting all projects in organization..." -ForegroundColor Yellow
$allProjects = Get-AllProjects -OrgName $Organization -PAT $PersonalAccessToken
if (-not $allProjects) {
    Write-Error "Failed to get projects. Exiting."
    exit 1
}
Write-Host "✓ Found $($allProjects.Count) projects" -ForegroundColor Green
Write-Host ""

# 2. Let user select which projects to process
$selectedProjects = Select-Projects -AllProjects $allProjects
if ($selectedProjects.Count -eq 0) {
    Write-Host "No projects selected. Exiting." -ForegroundColor Yellow
    exit 0
}

# 3. Get all existing webhooks (for duplicate checking)
Write-Host ""
Write-Host "Retrieving all existing webhooks..." -ForegroundColor Yellow
$allWebhooks = Get-WebhookSubscriptions -OrgName $Organization -PAT $PersonalAccessToken
if (-not $allWebhooks) {
    Write-Host "Could not retrieve existing webhooks, but will proceed with creation." -ForegroundColor Yellow
} else {
    Write-Host "✓ Found $($allWebhooks.Count) total existing webhook(s)" -ForegroundColor Green
}
Write-Host ""

# 4. Loop through each selected project and create missing webhooks
Write-Host "=== Processing Selected Projects ===" -ForegroundColor Cyan
foreach ($project in $selectedProjects) {
    $projectId = $project.id
    $projectName = $project.name
    
    Write-Host "--- Processing Project: $projectName ---" -ForegroundColor Cyan
    
    # Find existing Jellyfish webhooks for this specific project
    $existingProjectJellyfishHooks = $allWebhooks | Where-Object {
        ($_.publisherInputs.projectId -eq $projectId) -and
        ($_.consumerInputs.url -eq $jellyfishUrl)
    }

    # Loop through the event types we want to create
    foreach ($eventType in $eventTypesToCreate) {
        
        # Check if a hook for this event type *already exists*
        $hookExists = $existingProjectJellyfishHooks | Where-Object { $_.eventType -eq $eventType }
        
        if ($hookExists) {
            # If it exists, just report it and do nothing
            Write-Host "   • Webhook for '$eventType' already exists. (Status: $($hookExists.status))" -ForegroundColor Gray
        }
        else {
            # If it doesn't exist, create it
            Write-Host "   -> Creating webhook for '$eventType'..." -ForegroundColor Yellow
            Create-WebhookSubscription -OrgName $Organization -PAT $PersonalAccessToken -ProjectID $projectId -EventType $eventType -JellyfishUrl $jellyfishUrl -BearerToken $jellyfishBearerToken
        }
    }
    Write-Host "" # Add a space after each project
}

Write-Host "=== Webhook Creation Complete ===" -ForegroundColor Cyan
Write-Host "Processed $($selectedProjects.Count) selected project(s):" -ForegroundColor Green
foreach ($project in $selectedProjects) {
    Write-Host "  • $($project.name)" -ForegroundColor White
}
