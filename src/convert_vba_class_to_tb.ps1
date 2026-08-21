# Converts important VBA-sourced classes to native tB format

function Convert-VBAClassToTwinBasic {
    param(
        [string] $InputPath,
        [string] $OutputPath,
        [string] $ClassId,
        [string] $InterfaceId
    )

    Write-Verbose "Reading input: $InputPath"
    if (!(Test-Path -LiteralPath $InputPath)) { throw "InputPath not found: $InputPath" }

    # Read and normalize line endings
    $raw  = [System.IO.File]::ReadAllText($InputPath, [System.Text.Encoding]::UTF8)
    $text = $raw -replace "`r`n","`n" -replace "`r","`n"

    # Capture class name and flags BEFORE removal
    $className = ($text | Select-String -Pattern 'Attribute\s+VB_Name\s*=\s*"([^"]+)"' -AllMatches).Matches.Groups[1].Value
    if ([string]::IsNullOrWhiteSpace($className)) { throw 'Could not find Attribute VB_Name = "..." in input.' }

    $creatableMatch   = ($text | Select-String -Pattern 'Attribute\s+VB_Creatable\s*=\s*(True|False)').Matches
    $predeclaredMatch = ($text | Select-String -Pattern 'Attribute\s+VB_PredeclaredId\s*=\s*(True|False)').Matches
    $exposedMatch     = ($text | Select-String -Pattern 'Attribute\s+VB_Exposed\s*=\s*(True|False)').Matches

    $creatable   = if ($creatableMatch.Count   -gt 0) { $creatableMatch[0].Groups[1].Value } else { 'False' }
    $predeclared = if ($predeclaredMatch.Count -gt 0) { $predeclaredMatch[0].Groups[1].Value } else { 'False' }
    $exposed     = if ($exposedMatch.Count     -gt 0) { $exposedMatch[0].Groups[1].Value } else { 'False' }

    $isPublicClass = ($exposed -eq 'True')
    $hiddenValue   = if ($isPublicClass) { 'False' } else { 'True' }
    $scopeKeyword  = if ($isPublicClass) { 'Public' } else { 'Private' }

    # Capture module-level description, then remove it
    $moduleDescMatch = [regex]::Match(
        $text,
        'Attribute\s+VB_Description\s*=\s*"([^"]+)"',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $moduleDesc = if ($moduleDescMatch.Success) { $moduleDescMatch.Groups[1].Value } else { $null }
    if ($moduleDesc) {
        $text = [regex]::Replace(
            $text,
            'Attribute\s+VB_Description\s*=\s*"[^"]+"',
            '',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
    }

    # Strip VBA header block (VERSION ... END) if present
    $text = $text -replace '(?s)^VERSION\s+.*?END\s*', ''

    # Remove all legacy Attribute VB_* lines using multiline anchor
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*Attribute\s+VB_Name\s*=\s*".*?"\s*$',
        ''
    )
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*Attribute\s+VB_GlobalNameSpace\s*=\s*(True|False)\s*$',
        ''
    )
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*Attribute\s+VB_Creatable\s*=\s*(True|False)\s*$',
        ''
    )
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*Attribute\s+VB_PredeclaredId\s*=\s*(True|False)\s*$',
        ''
    )
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*Attribute\s+VB_Exposed\s*=\s*(True|False)\s*$',
        ''
    )
    # Remove VB_UserMemId attributes for NewEnum and Item
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*''?\s*Attribute\s+NewEnum\.VB_UserMemId\s*=\s*-4\s*$',
        ''
    )
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*''?\s*Attribute\s+Item\.VB_UserMemId\s*=\s*0\s*$',
        ''
    )

    # Remove @ModuleDescription, @Exposed, @folder lines (commented or not)
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*''?\s*@ModuleDescription\b.*$',
        ''
    )
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*''?\s*@Exposed\b.*$',
        ''
    )
    $text = [regex]::Replace(
        $text,
        '(?mi)^\s*''?\s*@folder\b.*$',
        ''
    )

    # Marker cleanup and remaps
    $text = $text -replace '(?im)^\s*''?@PredeclaredId\s*$', ''
    $text = $text -replace '(?im)^\s*''?@Enumerator\s*$', "[Enumerator]`n[Hidden]"
    $text = $text -replace '(?im)^\s*''?@DefaultMember\s*$', "[DefaultMember]"
    # Remove any @Description(...) comment lines
    $text = $text -replace '(?im)^\s*''?@Description\(".*"\)\s*$', ''

    # Collapse extra blank lines created by removals
    $text = [regex]::Replace($text, "(\n\s*){3,}", "`n`n")

    # Rewrite Replace() calls to VBA.Replace()
    $text = [regex]::Replace(
        $text,
        '(?<!\w)Replace\(',
        'VBA.Replace(',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $text = [regex]::Replace(
        $text,
        '(?<!\w)Replace\$\(',
        'VBA.Replace$(',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    # Rewrite Split() calls to VBA.Split()
    $text = [regex]::Replace(
        $text,
        '(?<!\w)Split\(',
        'VBA.Split(',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    # Build twinBASIC header
    function Test-Guid([string]$g) {
        if ([string]::IsNullOrWhiteSpace($g)) { return $false }
        try { [Guid]::Parse($g) | Out-Null; return $true } catch { return $false }
    }

    $header = @()
    if ($moduleDesc) { $header += '[Description("' + $moduleDesc + '")]' }
    $header += "[COMCreatable($creatable)]"
    $header += "[PredeclaredId($predeclared)]"
    $header += "[Hidden($hiddenValue)]"
    if (Test-Guid $ClassId)     { $header += "[ClassId(""$ClassId"")]" }
    if (Test-Guid $InterfaceId) { $header += "[InterfaceId(""$InterfaceId"")]" }
    $header += "$scopeKeyword Class $className"
    $headerBlock = ($header -join "`n") + "`n`n"

    # Method-level description: move Attribute <Member>.VB_Description to [Description] above
    $lines = $text -split "`n", -1
    $out   = New-Object 'System.Collections.Generic.List[string]'
    $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        $methodDeclMatch = [regex]::Match(
            $line,
            '^\s*(Public|Private)\s+(Sub|Function|Property)\b',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        if ($methodDeclMatch.Success) {
            $peek = if ($i + 1 -lt $lines.Count) { $lines[$i+1] } else { "" }
            $attrMatch = [regex]::Match(
                $peek,
                '^\s*Attribute\s+[A-Za-z0-9_]+\.\s*VB_Description\s*=\s*"([^"]+)"\s*$',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
            if ($attrMatch.Success) {
                $descText = $attrMatch.Groups[1].Value
                $descLine = '[Description("' + $descText + '")]'
                $out.Add($descLine)
                $out.Add($line)
                $i += 2
                continue
            }
        }
        $out.Add($line)
        $i++
    }

    $body = ($out -join "`n")

    # Assemble final text
    $tbText = $headerBlock + $body
    if ($tbText -notmatch '(?mi)^\s*End\s+Class\s*$') {
        $tbText = $tbText.TrimEnd() + "`n`nEnd Class`n"
    }
    $tbText = $tbText -replace "(\n\s*){3,}", "`n`n"

    # Write output with Windows CRLF
    $dir = Split-Path -Parent $OutputPath
    if (![string]::IsNullOrWhiteSpace($dir) -and !(Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
    $tbText = $tbText -replace "`n", "`r`n"
    [System.IO.File]::WriteAllText($OutputPath, $tbText, [System.Text.Encoding]::UTF8)
}


# --- Run line ---
Convert-VBAClassToTwinBasic `
  -InputPath .\Dictionary.cls `
  -OutputPath .\Dictionary.twin `
  -ClassId 'AE4D1399-8AFA-4866-BD86-96AB8CA9ECCE' `
  -InterfaceId 'AE4D1399-8AFA-4866-BD86-DE5BD5312E7C' `
  -Verbose
