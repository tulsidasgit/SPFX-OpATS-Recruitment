#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
    Creates all SharePoint lists, document libraries, and sample data
    for the Recruitment & Application Tracking System (OpRea).

.PARAMETER SiteUrl
    The full URL of the target SharePoint site.
    Example: https://yourtenant.sharepoint.com/sites/recruitment

.EXAMPLE
    .\CreateLists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/recruitment"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  OpRea — SharePoint Setup Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ── Connect ──────────────────────────────────────────────────────────────────

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive
Write-Host "Connected." -ForegroundColor Green

# ── Helper ───────────────────────────────────────────────────────────────────

function Ensure-List {
    param([string]$ListTitle, [string]$Template = "GenericList")
    $existing = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        New-PnPList -Title $ListTitle -Template $Template -OnQuickLaunch
        Write-Host "  Created list: $ListTitle" -ForegroundColor Green
    } else {
        Write-Host "  List already exists: $ListTitle (skipped)" -ForegroundColor DarkGray
    }
    return Get-PnPList -Identity $ListTitle
}

function Ensure-Field {
    param([string]$ListTitle, [string]$InternalName, [string]$DisplayName, [string]$Type, [hashtable]$Extra = @{})
    $existing = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        Write-Host "    Field '$InternalName' already exists (skipped)" -ForegroundColor DarkGray
        return
    }
    $params = @{
        List         = $ListTitle
        InternalName = $InternalName
        DisplayName  = $DisplayName
        Type         = $Type
        AddToDefaultView = $true
    }
    foreach ($k in $Extra.Keys) { $params[$k] = $Extra[$k] }
    Add-PnPField @params | Out-Null
    Write-Host "    Added field: $DisplayName ($Type)" -ForegroundColor Green
}

# ── 1. Departments ────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "1. Creating Departments list..." -ForegroundColor Cyan
Ensure-List -ListTitle "Departments"
Ensure-Field -ListTitle "Departments" -InternalName "IsActive" -DisplayName "IsActive" -Type "Boolean"

# ── 2. JobTitles ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "2. Creating JobTitles list..." -ForegroundColor Cyan
Ensure-List -ListTitle "JobTitles"
Ensure-Field -ListTitle "JobTitles" -InternalName "DepartmentId" -DisplayName "DepartmentId" -Type "Number"

# ── 3. JobOpenings ────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "3. Creating JobOpenings list..." -ForegroundColor Cyan
Ensure-List -ListTitle "JobOpenings"
Ensure-Field -ListTitle "JobOpenings" -InternalName "Department"      -DisplayName "Department"      -Type "Text"
Ensure-Field -ListTitle "JobOpenings" -InternalName "JobTitle"        -DisplayName "JobTitle"        -Type "Text"
Ensure-Field -ListTitle "JobOpenings" -InternalName "RequiredSkills"  -DisplayName "RequiredSkills"  -Type "Note"
Ensure-Field -ListTitle "JobOpenings" -InternalName "Experience"      -DisplayName "Experience"      -Type "Text"
Ensure-Field -ListTitle "JobOpenings" -InternalName "DueDate"         -DisplayName "DueDate"         -Type "DateTime"
Ensure-Field -ListTitle "JobOpenings" -InternalName "PostedBy"        -DisplayName "PostedBy"        -Type "Text"
Ensure-Field -ListTitle "JobOpenings" -InternalName "LinkedInUrl"     -DisplayName "LinkedInUrl"     -Type "Text"

# Status choice field — add via XML for choices
$statusFieldXml = @"
<Field Type="Choice" DisplayName="Status" Name="Status" Required="FALSE" AddToDefaultView="TRUE">
  <CHOICES>
    <CHOICE>Open</CHOICE>
    <CHOICE>In Progress</CHOICE>
    <CHOICE>Closed</CHOICE>
  </CHOICES>
  <Default>Open</Default>
</Field>
"@
$existingStatus = Get-PnPField -List "JobOpenings" -Identity "Status" -ErrorAction SilentlyContinue
if ($null -eq $existingStatus) {
    Add-PnPFieldFromXml -List "JobOpenings" -FieldXml $statusFieldXml | Out-Null
    Write-Host "    Added field: Status (Choice)" -ForegroundColor Green
} else {
    Write-Host "    Field 'Status' already exists (skipped)" -ForegroundColor DarkGray
}

# ── 4. AllowedPosters ─────────────────────────────────────────────────────────

Write-Host ""
Write-Host "4. Creating AllowedPosters list..." -ForegroundColor Cyan
Ensure-List -ListTitle "AllowedPosters"
Ensure-Field -ListTitle "AllowedPosters" -InternalName "UserEmail" -DisplayName "UserEmail" -Type "Text"

# ── 5. Candidates ─────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "5. Creating Candidates list..." -ForegroundColor Cyan
Ensure-List -ListTitle "Candidates"
Ensure-Field -ListTitle "Candidates" -InternalName "JobOpeningId"    -DisplayName "JobOpeningId"    -Type "Number"
Ensure-Field -ListTitle "Candidates" -InternalName "CandidateName"   -DisplayName "CandidateName"   -Type "Text"
Ensure-Field -ListTitle "Candidates" -InternalName "Email"           -DisplayName "Email"           -Type "Text"
Ensure-Field -ListTitle "Candidates" -InternalName "Phone"           -DisplayName "Phone"           -Type "Text"
Ensure-Field -ListTitle "Candidates" -InternalName "ResumeUrl"       -DisplayName "ResumeUrl"       -Type "Text"
Ensure-Field -ListTitle "Candidates" -InternalName "FitmentScore"    -DisplayName "FitmentScore"    -Type "Number"
Ensure-Field -ListTitle "Candidates" -InternalName "MatchingSkills"  -DisplayName "MatchingSkills"  -Type "Note"
Ensure-Field -ListTitle "Candidates" -InternalName "MissingSkills"   -DisplayName "MissingSkills"   -Type "Note"
Ensure-Field -ListTitle "Candidates" -InternalName "AISummary"       -DisplayName "AISummary"       -Type "Note"
Ensure-Field -ListTitle "Candidates" -InternalName "HRFeedback"      -DisplayName "HRFeedback"      -Type "Note"

$recommendationXml = @"
<Field Type="Choice" DisplayName="Recommendation" Name="Recommendation" Required="FALSE" AddToDefaultView="TRUE">
  <CHOICES>
    <CHOICE>Recommended</CHOICE>
    <CHOICE>Maybe</CHOICE>
    <CHOICE>Not Recommended</CHOICE>
  </CHOICES>
</Field>
"@
$existingRec = Get-PnPField -List "Candidates" -Identity "Recommendation" -ErrorAction SilentlyContinue
if ($null -eq $existingRec) {
    Add-PnPFieldFromXml -List "Candidates" -FieldXml $recommendationXml | Out-Null
    Write-Host "    Added field: Recommendation (Choice)" -ForegroundColor Green
} else {
    Write-Host "    Field 'Recommendation' already exists (skipped)" -ForegroundColor DarkGray
}

$experienceMatchXml = @"
<Field Type="Choice" DisplayName="ExperienceMatch" Name="ExperienceMatch" Required="FALSE" AddToDefaultView="TRUE">
  <CHOICES>
    <CHOICE>meets</CHOICE>
    <CHOICE>exceeds</CHOICE>
    <CHOICE>below</CHOICE>
  </CHOICES>
</Field>
"@
$existingExp = Get-PnPField -List "Candidates" -Identity "ExperienceMatch" -ErrorAction SilentlyContinue
if ($null -eq $existingExp) {
    Add-PnPFieldFromXml -List "Candidates" -FieldXml $experienceMatchXml | Out-Null
    Write-Host "    Added field: ExperienceMatch (Choice)" -ForegroundColor Green
} else {
    Write-Host "    Field 'ExperienceMatch' already exists (skipped)" -ForegroundColor DarkGray
}

# ── 6. Interviews ─────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "6. Creating Interviews list..." -ForegroundColor Cyan
Ensure-List -ListTitle "Interviews"
Ensure-Field -ListTitle "Interviews" -InternalName "CandidateId"      -DisplayName "CandidateId"      -Type "Number"
Ensure-Field -ListTitle "Interviews" -InternalName "JobOpeningId"     -DisplayName "JobOpeningId"     -Type "Number"
Ensure-Field -ListTitle "Interviews" -InternalName "InterviewerEmail" -DisplayName "InterviewerEmail" -Type "Text"
Ensure-Field -ListTitle "Interviews" -InternalName "ScheduledDate"    -DisplayName "ScheduledDate"    -Type "DateTime"
Ensure-Field -ListTitle "Interviews" -InternalName "Feedback"         -DisplayName "Feedback"         -Type "Note"
Ensure-Field -ListTitle "Interviews" -InternalName "HRNotes"          -DisplayName "HRNotes"          -Type "Note"

$roundXml = @"
<Field Type="Choice" DisplayName="InterviewRound" Name="InterviewRound" Required="FALSE" AddToDefaultView="TRUE">
  <CHOICES>
    <CHOICE>1</CHOICE>
    <CHOICE>2</CHOICE>
    <CHOICE>3</CHOICE>
  </CHOICES>
</Field>
"@
$existingRound = Get-PnPField -List "Interviews" -Identity "InterviewRound" -ErrorAction SilentlyContinue
if ($null -eq $existingRound) {
    Add-PnPFieldFromXml -List "Interviews" -FieldXml $roundXml | Out-Null
    Write-Host "    Added field: InterviewRound (Choice)" -ForegroundColor Green
} else {
    Write-Host "    Field 'InterviewRound' already exists (skipped)" -ForegroundColor DarkGray
}

$feedbackStatusXml = @"
<Field Type="Choice" DisplayName="FeedbackStatus" Name="FeedbackStatus" Required="FALSE" AddToDefaultView="TRUE">
  <CHOICES>
    <CHOICE>Pending</CHOICE>
    <CHOICE>Submitted</CHOICE>
  </CHOICES>
  <Default>Pending</Default>
</Field>
"@
$existingFS = Get-PnPField -List "Interviews" -Identity "FeedbackStatus" -ErrorAction SilentlyContinue
if ($null -eq $existingFS) {
    Add-PnPFieldFromXml -List "Interviews" -FieldXml $feedbackStatusXml | Out-Null
    Write-Host "    Added field: FeedbackStatus (Choice)" -ForegroundColor Green
} else {
    Write-Host "    Field 'FeedbackStatus' already exists (skipped)" -ForegroundColor DarkGray
}

# ── 7. Resumes document library + folders ─────────────────────────────────────

Write-Host ""
Write-Host "7. Creating Resumes document library and folders..." -ForegroundColor Cyan
$resumesLib = Get-PnPList -Identity "Resumes" -ErrorAction SilentlyContinue
if ($null -eq $resumesLib) {
    New-PnPList -Title "Resumes" -Template DocumentLibrary -OnQuickLaunch
    Write-Host "  Created document library: Resumes" -ForegroundColor Green
} else {
    Write-Host "  Library 'Resumes' already exists (skipped)" -ForegroundColor DarkGray
}

$folders = @(
    "Engineering",
    "Engineering/SoftwareEngineer",
    "Engineering/DevOpsEngineer",
    "Engineering/QAEngineer",
    "HR",
    "HR/HRManager",
    "HR/TalentAcquisition",
    "Finance",
    "Finance/Analyst",
    "Marketing"
)
foreach ($folder in $folders) {
    $existing = Get-PnPFolder -Url "Resumes/$folder" -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        Add-PnPFolder -Name ($folder.Split("/")[-1]) -Folder ("Resumes/" + ($folder -replace "/[^/]+$", "").TrimEnd("/")) -ErrorAction SilentlyContinue | Out-Null
        Write-Host "  Created folder: Resumes/$folder" -ForegroundColor Green
    } else {
        Write-Host "  Folder 'Resumes/$folder' already exists (skipped)" -ForegroundColor DarkGray
    }
}

# ── 8. Sample data ────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "8. Adding sample data..." -ForegroundColor Cyan

# AllowedPosters
$allowedPosters = @(
    "tulsidas.rp@sintecmedia.com",
    "dev1@company.onmicrosoft.com",
    "dev2@company.onmicrosoft.com",
    "dev3@company.onmicrosoft.com"
)
foreach ($email in $allowedPosters) {
    $exists = Get-PnPListItem -List "AllowedPosters" -Query "<View><Query><Where><Eq><FieldRef Name='UserEmail'/><Value Type='Text'>$email</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $exists -or $exists.Count -eq 0) {
        Add-PnPListItem -List "AllowedPosters" -Values @{ Title = $email; UserEmail = $email } | Out-Null
        Write-Host "  Added AllowedPoster: $email" -ForegroundColor Green
    } else {
        Write-Host "  AllowedPoster '$email' already exists (skipped)" -ForegroundColor DarkGray
    }
}

# Departments
$departments = @(
    @{ Title = "Engineering"; IsActive = $true },
    @{ Title = "HR";          IsActive = $true },
    @{ Title = "Finance";     IsActive = $true },
    @{ Title = "Marketing";   IsActive = $true }
)
$deptIds = @{}
foreach ($dept in $departments) {
    $existing = Get-PnPListItem -List "Departments" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($dept.Title)</Value></Eq></Where></Query></View>"
    if ($null -eq $existing -or $existing.Count -eq 0) {
        $item = Add-PnPListItem -List "Departments" -Values @{ Title = $dept.Title; IsActive = $dept.IsActive }
        $deptIds[$dept.Title] = $item.Id
        Write-Host "  Added Department: $($dept.Title) (ID: $($item.Id))" -ForegroundColor Green
    } else {
        $deptIds[$dept.Title] = $existing[0].Id
        Write-Host "  Department '$($dept.Title)' already exists (ID: $($existing[0].Id)) (skipped)" -ForegroundColor DarkGray
    }
}

# JobTitles
$jobTitles = @(
    @{ Title = "Software Engineer";   Dept = "Engineering" },
    @{ Title = "DevOps Engineer";     Dept = "Engineering" },
    @{ Title = "QA Engineer";         Dept = "Engineering" },
    @{ Title = "HR Manager";          Dept = "HR" },
    @{ Title = "Talent Acquisition";  Dept = "HR" },
    @{ Title = "Financial Analyst";   Dept = "Finance" }
)
foreach ($jt in $jobTitles) {
    $deptId = $deptIds[$jt.Dept]
    $existing = Get-PnPListItem -List "JobTitles" -Query "<View><Query><Where><And><Eq><FieldRef Name='Title'/><Value Type='Text'>$($jt.Title)</Value></Eq><Eq><FieldRef Name='DepartmentId'/><Value Type='Number'>$deptId</Value></Eq></And></Where></Query></View>"
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "JobTitles" -Values @{ Title = $jt.Title; DepartmentId = $deptId } | Out-Null
        Write-Host "  Added JobTitle: $($jt.Title) (Dept: $($jt.Dept))" -ForegroundColor Green
    } else {
        Write-Host "  JobTitle '$($jt.Title)' already exists (skipped)" -ForegroundColor DarkGray
    }
}

# ── Done ──────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Setup complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Verify lists in: $SiteUrl/_layouts/15/viewlsts.aspx"
Write-Host "  2. Update AllowedPosters with your real user emails"
Write-Host "  3. Configure the web part property pane with your Anthropic API key"
Write-Host ""
