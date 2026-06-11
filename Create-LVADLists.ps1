#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
    Creates the LVAD_Patients SharePoint list for the LVAD Map app.
.PARAMETER SiteUrl
    Full URL of the target SharePoint site.
    Example: https://bswh.sharepoint.com/sites/EMS
.NOTES
    Prerequisites:  Install-Module PnP.PowerShell -Scope CurrentUser
    Run:            .\Create-LVADLists.ps1 -SiteUrl "https://bswh.sharepoint.com/sites/EMS"
    Safe to re-run — existing lists and columns are skipped, not duplicated.
#>
param(
    [Parameter(Mandatory)]
    [string]$SiteUrl
)

Connect-PnPOnline -Url $SiteUrl -Interactive

function Ensure-List {
    param([string]$Name)
    $list = Get-PnPList -Identity $Name -ErrorAction SilentlyContinue
    if ($list) {
        Write-Host "  List '$Name' already exists — skipping." -ForegroundColor DarkYellow
    } else {
        New-PnPList -Title $Name -Template GenericList -OnQuickLaunch | Out-Null
        Write-Host "  Created list: $Name" -ForegroundColor Green
    }
}

function Add-Col {
    param([string]$List, [string]$Display, [string]$Internal, [string]$Type, [int]$Decimals = -1)
    if (Get-PnPField -List $List -Identity $Internal -ErrorAction SilentlyContinue) {
        Write-Host "    Column '$Internal' already exists — skipping." -ForegroundColor DarkYellow
        return
    }
    Add-PnPField -List $List -DisplayName $Display -InternalName $Internal -Type $Type -AddToDefaultView | Out-Null
    if ($Decimals -ge 0) {
        Set-PnPField -List $List -Identity $Internal -Values @{ Decimals = $Decimals }
    }
    Write-Host "    Added column: $Internal ($Type)" -ForegroundColor Cyan
}

# ─── LVAD_Patients ────────────────────────────────────────────────────────────
Write-Host "`nSetting up LVAD_Patients..." -ForegroundColor White
Ensure-List -Name "LVAD_Patients"

# Title is already present — we use it as the patient name (Last, First)
Add-Col -List "LVAD_Patients" -Display "Patient UID"       -Internal "PatientUID"       -Type Text
Add-Col -List "LVAD_Patients" -Display "Hospital ID"       -Internal "HospitalId"       -Type Text
Add-Col -List "LVAD_Patients" -Display "Phone"             -Internal "Phone"            -Type Text
Add-Col -List "LVAD_Patients" -Display "Emergency Contact" -Internal "EmergencyContact" -Type Text
Add-Col -List "LVAD_Patients" -Display "LVAD Settings"     -Internal "LVADSettings"     -Type Text
Add-Col -List "LVAD_Patients" -Display "Home Address"      -Internal "HomeAddress"      -Type Text
Add-Col -List "LVAD_Patients" -Display "Latitude"          -Internal "Latitude"         -Type Number -Decimals 6
Add-Col -List "LVAD_Patients" -Display "Longitude"         -Internal "Longitude"        -Type Number -Decimals 6
Add-Col -List "LVAD_Patients" -Display "Notes"             -Internal "Notes"            -Type Note
Add-Col -List "LVAD_Patients" -Display "Added At"          -Internal "AddedAt"          -Type DateTime

# LVAD Device as a Choice column
if (Get-PnPField -List "LVAD_Patients" -Identity "LVADDevice" -ErrorAction SilentlyContinue) {
    Write-Host "    Column 'LVADDevice' already exists — skipping." -ForegroundColor DarkYellow
} else {
    Add-PnPFieldFromXml -List "LVAD_Patients" -FieldXml @"
<Field Type="Choice" DisplayName="LVAD Device" Name="LVADDevice" Required="FALSE">
  <CHOICES>
    <CHOICE>HeartMate 3</CHOICE>
    <CHOICE>HeartMate II</CHOICE>
    <CHOICE>HeartWare HVAD</CHOICE>
    <CHOICE>Jarvik 2000</CHOICE>
    <CHOICE>EVAHEART 2</CHOICE>
  </CHOICES>
</Field>
"@ | Out-Null
    Write-Host "    Added column: LVADDevice (Choice)" -ForegroundColor Cyan
}

Write-Host "`nLVAD_Patients list is ready." -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Set SHAREPOINT_SITE_URL=$SiteUrl in your .env or Vercel env vars."
Write-Host "  2. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET from your App Registration."
Write-Host "  3. Restart the app — it will automatically use this SharePoint list."
Write-Host ""
Write-Host "REMINDER: Confirm your M365 tenant has an active Microsoft HIPAA BAA" -ForegroundColor Yellow
Write-Host "before storing live patient data." -ForegroundColor Yellow
