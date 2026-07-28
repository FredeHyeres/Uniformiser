# ==============================================================================
# Uniformiser - installation automatique (MicroStation V8i SS3)
#
# Copie le projet VBA (.mvba) au bon endroit, puis configure le chargement
# automatique du projet au demarrage.
#
# Usage : clic droit sur install.cmd > Executer  (ou .\install.ps1 en PowerShell)
# Relancable sans risque : chaque etape est ignoree si deja faite.
# ==============================================================================

$ErrorActionPreference = "Stop"
Write-Host "=== Installation Uniformiser ===" -ForegroundColor Cyan

# --- 0. Emplacements ---------------------------------------------------------
$Source    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Workspace = Join-Path $env:USERPROFILE "Documents\MicroStV8i\WorkSpace"

if (-not (Test-Path $Workspace)) {
    # Autre emplacement possible : installation par defaut de V8i
    $Alt = "C:\ProgramData\Bentley\MicroStation V8i (SELECTseries)\WorkSpace"
    if (Test-Path $Alt) { $Workspace = $Alt }
    else {
        Write-Host "ERREUR : workspace MicroStation introuvable." -ForegroundColor Red
        Write-Host "Cherche : $Workspace"
        Write-Host "Modifiez la variable `$Workspace en tete de script si votre"
        Write-Host "workspace est ailleurs, puis relancez."
        exit 1
    }
}
Write-Host "Workspace : $Workspace"

# --- 1. Installer le projet VBA (Uniformiser.mvba dedie) ---------------------
# Le Default.mvba existe deja sur tout MicroStation et ne doit jamais etre
# ecrase : la macro est donc livree dans son propre projet Uniformiser.mvba,
# copie dans le workspace et charge automatiquement via le .ucf (etape 3).
$Mvba = Join-Path $Source "Uniformiser.mvba"
$VbaDir = Join-Path $Workspace "Standards\vba"
$MvbaOk = $false
if (Test-Path $Mvba) {
    New-Item -ItemType Directory -Force $VbaDir | Out-Null
    Copy-Item $Mvba $VbaDir -Force
    $MvbaOk = $true
    Write-Host "[OK] Uniformiser.mvba copie vers $VbaDir" -ForegroundColor Green
} else {
    Write-Host "[!] Uniformiser.mvba absent du dossier d'installation :" -ForegroundColor Yellow
    Write-Host "    importez les fichiers de src\ dans un projet VBA (voir README)."
}

# --- 1b. Neutraliser les copies masquantes dans les projets MicroStation -----
# MicroStation cherche le .mvba d'abord dans le dossier vba du projet actif
# (WorkSpace\Projects\<projet>\vba) : une vieille copie y masquerait la version
# installee dans Standards\vba. On les renomme en .bak (aucune suppression).
if ($MvbaOk) {
    $ProjetsDir = Join-Path $Workspace "Projects"
    if (Test-Path $ProjetsDir) {
        $Masquantes = @(Get-ChildItem $ProjetsDir -Recurse -Filter "Uniformiser.mvba" -File -ErrorAction SilentlyContinue)
        foreach ($M in $Masquantes) {
            $Bak = "$($M.FullName).bak"
            if (Test-Path $Bak) { Remove-Item $Bak -Force }
            Rename-Item $M.FullName $Bak
            Write-Host "[OK] Copie masquante neutralisee : $($M.FullName) -> .bak" -ForegroundColor Yellow
        }
        if ($Masquantes.Count -eq 0) {
            Write-Host "[OK] Aucune copie masquante dans $ProjetsDir" -ForegroundColor Green
        }
    }
}

# --- 2. Chargement automatique du projet (fichier .ucf utilisateur) ----------
if ($MvbaOk) {
    $UsersDir = Join-Path $Workspace "Users"
    $UcfFiles = @()
    if (Test-Path $UsersDir) {
        $UcfFiles = @(Get-ChildItem $UsersDir -Filter *.ucf -File)
    }
    if ($UcfFiles.Count -eq 0) {
        Write-Host "[!] Aucun fichier .ucf dans $UsersDir : configurez l'autoload" -ForegroundColor Yellow
        Write-Host "    dans MicroStation (Utilities > Macros > Project Manager >"
        Write-Host "    clic droit sur Uniformiser > Autoload)."
    } else {
        foreach ($UcfFile in $UcfFiles) {
            $Contenu = Get-Content $UcfFile.FullName -Raw
            if ($Contenu -match "MS_VBAAUTOLOADPROJECTS.*Uniformiser") {
                Write-Host "[OK] Autoload deja configure dans $($UcfFile.Name)" -ForegroundColor Green
            } else {
                $Lignes = "`r`n# --- Uniformiser (ajoute par install.ps1) ---`r`n" + `
                          "MS_VBASEARCHDIRECTORIES < $VbaDir\`r`n" + `
                          "MS_VBAAUTOLOADPROJECTS > Uniformiser.mvba`r`n"
                Add-Content -Path $UcfFile.FullName -Value $Lignes -Encoding ASCII
                Write-Host "[OK] Autoload configure dans $($UcfFile.Name)" -ForegroundColor Green
            }
        }
    }
}

# --- 3. Recapitulatif --------------------------------------------------------
Write-Host ""
Write-Host "=== Installation terminee ===" -ForegroundColor Cyan
Write-Host "1. (Re)demarrez MicroStation"
Write-Host "2. Key-in : vba run [UniformiserV1]Uniformiser"
Write-Host "3. Affectation touche : Utilitaires > Touches de fonction (ex. F8)"
Write-Host ""
Read-Host "Appuyez sur Entree pour fermer"
