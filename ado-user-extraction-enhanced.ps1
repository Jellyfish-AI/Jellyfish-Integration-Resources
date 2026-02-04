# Azure DevOps User and Team Extraction Script - Enhanced Version
# This script pulls active users and their team memberships from Azure DevOps
# with flexible project selection options (select specific, exclude by name, or all).

# --- Configuration ---
$orgName = "YourOrgName" 
$orgUrl = "https://dev.azure.com/$orgName"
$outputPath = "$home\Desktop\ADO_Active_User_Team_Mapping.csv"

Write-Host "=== Azure DevOps User and Team Extraction Script - Enhanced ===" -ForegroundColor Cyan
Write-Host "Organization: $orgName"
Write-Host "Output file: $outputPath"
Write-Host ""

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

# Function to select projects by number
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

# --- Main Script Execution ---

# Check if Azure CLI is installed and logged in
try {
    $null = az account show 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Azure CLI is not logged in. Please run 'az login' first."
        exit 1
    }
} catch {
    Write-Error "Azure CLI is not installed or not accessible. Please install Azure CLI first."
    exit 1
}

Write-Host "Fetching active user list to verify status..." -ForegroundColor Yellow
# Get active user entitlements (this helps us filter out inactive/disabled users)
try {
    $activeUsers = az devops user list --org $orgUrl | ConvertFrom-Json
    $activeEmails = $activeUsers.members.user.uniqueName
    Write-Host "✓ Found $($activeEmails.Count) active users" -ForegroundColor Green
} catch {
    Write-Error "Failed to fetch active users. Please check your organization name and permissions."
    exit 1
}

# Get all projects
Write-Host "Getting all projects in organization..." -ForegroundColor Yellow
try {
    $allProjectsResponse = az devops project list --org $orgUrl | ConvertFrom-Json
    $allProjects = $allProjectsResponse.value
    Write-Host "✓ Found $($allProjects.Count) projects" -ForegroundColor Green
} catch {
    Write-Error "Failed to fetch projects. Please check your organization name and permissions."
    exit 1
}

Write-Host ""

# Let user select which projects to process
$selectedProjects = Select-Projects -AllProjects $allProjects
if ($selectedProjects.Count -eq 0) {
    Write-Host "No projects selected. Exiting." -ForegroundColor Yellow
    exit 0
}

# Initialize results array
$results = @()
$totalProjects = $selectedProjects.Count
$counter = 0

Write-Host ""
Write-Host "=== Processing Selected Projects ===" -ForegroundColor Cyan

foreach ($project in $selectedProjects) {
    $counter++
    $percent = [Math]::Round(($counter / $totalProjects) * 100, 2)
    Write-Progress -Activity "Extracting Team Data" -Status "Project $counter of $totalProjects ($percent%) - $($project.name)" -PercentComplete $percent
    Write-Host "--- Processing Project: $($project.name) ---" -ForegroundColor Cyan

    try {
        # Get all teams in the project
        $teams = az devops team list --org $orgUrl --project $project.id | ConvertFrom-Json
        
        if ($teams -and $teams.Count -gt 0) {
            Write-Host "  Found $($teams.Count) team(s)" -ForegroundColor Gray
            
            foreach ($team in $teams) {
                try {
                    # Get members of the specific team
                    $members = az devops team member list --org $orgUrl --project $project.id --team $team.id | ConvertFrom-Json
                    
                    $activeMembers = 0
                    foreach ($member in $members) {
                        # FILTER: Only add if the user is in our active list
                        if ($activeEmails -contains $member.identity.uniqueName) {
                            $obj = [PSCustomObject]@{
                                DisplayName = $member.identity.displayName
                                Email       = $member.identity.uniqueName
                                Project     = $project.name
                                TeamName    = $team.name
                                Status      = "Active"
                            }
                            $results += $obj
                            $activeMembers++
                        }
                    }
                    
                    if ($activeMembers -gt 0) {
                        Write-Host "    ✓ Team '$($team.name)': $activeMembers active member(s)" -ForegroundColor Green
                    } else {
                        Write-Host "    • Team '$($team.name)': no active members" -ForegroundColor Gray
                    }
                    
                    # Rate-limit buffer: Small sleep after every team fetch
                    Start-Sleep -Milliseconds 150
                } catch {
                    Write-Host "    ❌ Failed to get members for team '$($team.name)'" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  No teams found in this project" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  ❌ Failed to get teams for project '$($project.name)'" -ForegroundColor Red
    }
    
    Write-Host ""
}

Write-Progress -Activity "Extracting Team Data" -Completed

# Export to CSV
Write-Host "Exporting results to CSV..." -ForegroundColor Yellow
try {
    $results | Export-Csv -Path $outputPath -NoTypeInformation
    Write-Host "✓ Successfully exported $($results.Count) user-team mappings" -ForegroundColor Green
} catch {
    Write-Error "Failed to export CSV: $($_.Exception.Message)"
    exit 1
}

Write-Host ""
Write-Host "=== User and Team Extraction Complete ===" -ForegroundColor Cyan
Write-Host "Processed $($selectedProjects.Count) selected project(s):" -ForegroundColor Green
foreach ($project in $selectedProjects) {
    Write-Host "  • $($project.name)" -ForegroundColor White
}
Write-Host ""
Write-Host "Results:"
Write-Host "  • Total user-team mappings: $($results.Count)" -ForegroundColor Green
Write-Host "  • Unique users found: $(($results | Select-Object -Unique Email).Count)" -ForegroundColor Green
Write-Host "  • File saved to: $outputPath" -ForegroundColor Green