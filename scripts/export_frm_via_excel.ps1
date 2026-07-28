# Genere un couple .frm/.frx valide pour un UserForm VBA via l'editeur VBA d'Excel.
# Voir docs\Guide_UserForm_FRM_FRX.md pour le mode d'emploi complet.
#
# 1. Active temporairement AccessVBOM dans le registre Excel (restaure a la fin)
# 2. Cree un UserForm vide nomme $FormName, injecte le code lu dans $SrcFrm
#    (en-tete designer et lignes Attribute filtres), exporte le couple
#    $FormName_export.frm + $FormName_export.frx dans $OutDir
# 3. Renommer ensuite le couple ET corriger la ligne OleObjectBlob du .frm
#
# Exemple :
#   .\scripts\export_frm_via_excel.ps1 -SrcFrm "src\frmUniformiser.frm" `
#       -OutDir "src" -FormName "frmUniformiser"
param(
    [string]$SrcFrm   = "src\frmUniformiser.frm",
    [string]$OutDir   = "src",
    [string]$FormName = "frmUniformiser"
)
$ErrorActionPreference = "Stop"

$SrcFrm = (Resolve-Path $SrcFrm).Path
$OutDir = (Resolve-Path $OutDir).Path

# --- Extraire le code : tout apres l'en-tete designer, sans les lignes Attribute ---
$lines = [IO.File]::ReadAllLines($SrcFrm)
$start = 0
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "^Attribute VB_Exposed") { $start = $i + 1; break }
}
$codeLines = $lines[$start..($lines.Count - 1)] | Where-Object { $_ -notmatch "^Attribute " }
$code = ($codeLines -join "`r`n")

# --- Version d'Excel installee (ex. "Excel.Application.12" -> "12.0") ---
$curVer = (Get-ItemProperty "HKLM:\SOFTWARE\Classes\Excel.Application\CurVer")."(default)"
$verNum = ($curVer -split "\.")[-1] + ".0"

# --- Registre : AccessVBOM (acces approuve au modele d'objet VBA) ---
$regPath = "HKCU:\Software\Microsoft\Office\$verNum\Excel\Security"
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
$old = $null
try { $old = (Get-ItemProperty $regPath -Name AccessVBOM -ErrorAction Stop).AccessVBOM } catch {}
Set-ItemProperty $regPath -Name AccessVBOM -Value 1 -Type DWord
Write-Output "AccessVBOM : ancienne valeur = $old, mise a 1"

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $vbp = $wb.VBProject
    $comp = $vbp.VBComponents.Add(3)   # 3 = vbext_ct_MSForm
    $comp.Name = $FormName
    $comp.CodeModule.AddFromString($code) | Out-Null

    $comp.Export("$OutDir\$($FormName)_export.frm")
    Write-Output "Export OK : $OutDir\$($FormName)_export.frm + .frx"
    Write-Output "RAPPEL : renommer le couple et corriger la ligne OleObjectBlob du .frm"

    $wb.Close($false)
} finally {
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    # Restaurer le registre
    if ($null -eq $old) { Remove-ItemProperty $regPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty $regPath -Name AccessVBOM -Value $old -Type DWord }
    Write-Output "AccessVBOM restaure ($old)"
}
