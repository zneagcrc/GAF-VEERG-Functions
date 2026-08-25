<#
.SYNOPSIS
  Export the Copilot repository memory (/memories/repo) into the workspace as
  plain markdown, so the notes are readable outside Copilot chat and can be
  version-controlled.

.DESCRIPTION
  Copilot stores repo memory under VS Code user storage, OUTSIDE the repo:
    <AppData>\Code\User\workspaceStorage\<hash>\GitHub.copilot-chat\memory-tool\memories\repo
  This finds the <hash> folder whose workspace.json points at THIS repo, then
  copies every *.md into the output folder (default: chat-memory\repo-memory).
  Re-run any time to refresh after the agent updates its memory.

.PARAMETER RepoRoot
  Repository root. Defaults to the parent of the scripts folder.

.PARAMETER OutputDir
  Where to write the exported markdown. Defaults to chat-memory\repo-memory.
#>
param(
  [string] $RepoRoot = (Split-Path $PSScriptRoot -Parent),
  [string] $OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoPath = (Resolve-Path -LiteralPath $RepoRoot).Path.TrimEnd('\', '/')
if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = Join-Path $repoPath 'chat-memory\repo-memory' }

# VS Code stable and Insiders storage roots.
$storageRoots = @(
  @(
    (Join-Path $env:APPDATA 'Code\User\workspaceStorage'),
    (Join-Path $env:APPDATA 'Code - Insiders\User\workspaceStorage')
  ) | Where-Object { Test-Path -LiteralPath $_ }
)
if ($storageRoots.Count -eq 0) { throw 'No VS Code workspaceStorage folder found under %APPDATA%.' }

# Find the memory-tool\memories\repo whose workspaceStorage entry belongs to this
# repo. Prefer an exact file: path match; otherwise match the workspace folder NAME
# (handles vscode-remote://.../workspaces/<name> Codespaces URIs), tie-broken by the
# most recently written memory (i.e. the active workspace).
$repoLeaf = Split-Path $repoPath -Leaf
$candidates = New-Object System.Collections.Generic.List[object]
foreach ($root in $storageRoots) {
  foreach ($d in Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue) {
    $memRepoDir = Join-Path $d.FullName 'GitHub.copilot-chat\memory-tool\memories\repo'
    if (-not (Test-Path -LiteralPath $memRepoDir)) { continue }
    $folder = $null
    $wsJson = Join-Path $d.FullName 'workspace.json'
    if (Test-Path -LiteralPath $wsJson) { try { $folder = (Get-Content -LiteralPath $wsJson -Raw | ConvertFrom-Json).folder } catch { } }
    $exact = $false; $leafMatch = $false
    if (-not [string]::IsNullOrWhiteSpace($folder)) {
      if ($folder -like 'file:*') { try { if (([Uri] $folder).LocalPath.TrimEnd('\', '/') -ieq $repoPath) { $exact = $true } } catch { } }
      $leaf = [Uri]::UnescapeDataString(($folder.TrimEnd('/', '\') -split '[\\/]')[-1])
      if ($leaf -ieq $repoLeaf) { $leafMatch = $true }
    }
    $lastWrite = (Get-ChildItem -LiteralPath $memRepoDir -Filter '*.md' -File -ErrorAction SilentlyContinue | Measure-Object -Property LastWriteTime -Maximum).Maximum
    $candidates.Add([pscustomobject]@{ Path = $memRepoDir; Exact = $exact; Leaf = $leafMatch; LastWrite = $lastWrite })
  }
}
$pick = @($candidates | Where-Object { $_.Exact }) | Select-Object -First 1
if ($null -eq $pick) { $pick = @($candidates | Where-Object { $_.Leaf } | Sort-Object LastWrite -Descending) | Select-Object -First 1 }
if ($null -eq $pick) { throw "Could not locate this workspace's Copilot repo-memory folder (no workspaceStorage entry matching '$repoLeaf' with memory-tool\memories\repo)." }
$memRepo = $pick.Path

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$files = @(Get-ChildItem -LiteralPath $memRepo -Filter '*.md' -File)
foreach ($f in $files) { Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $OutputDir $f.Name) -Force }

Write-Host ("Exported {0} memory file(s)" -f $files.Count)
Write-Host ("  from: {0}" -f $memRepo)
Write-Host ("  to:   {0}" -f $OutputDir)
