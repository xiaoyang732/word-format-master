param(
  [Parameter(Mandatory=$true)][string]$InputPath,
  [string]$OutputPath,
  [ValidateSet('inspect','render','refresh','convert')][string]$Mode = 'inspect',
  [ValidateSet('word','wps')][string]$Backend = 'word',
  [Parameter(Mandatory=$true)][string]$EvidencePath
)
$ErrorActionPreference = 'Stop'

# Word / Office COM constants
$WD_UNDEFINED = 9999999
$WD_FORMAT_XML_DOCUMENT = 12
$WD_EXPORT_FORMAT_PDF = 17
$WD_STATISTIC_PAGES = 2
$WD_INFO_ADJUSTED_PAGE = 1
$WD_INFO_PHYSICAL_PAGE = 3
$WD_COLLAPSE_START = 1
$MSO_SHAPE_GROUP = 6
$MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
$MAX_PAGINATION_PASSES = 3
$MAX_REASONABLE_FONT_PT = 1000
$HEADER_FOOTER_STORY_TYPES = @(6, 7, 8, 9, 10, 11)

$app = $null
$document = $null
$owned = $false
$ownedProcesses = @()
$syncWarnings = @()
$syncAdjustments = 0
$processName = if ($Backend -eq 'word') { 'WINWORD' } else { 'wps' }
$existingIds = @(Get-Process -Name $processName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

function Write-Evidence($Path, $Value) {
  $Value | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Get-StoryRanges($Doc) {
  foreach ($rootStory in $Doc.StoryRanges) {
    $story = $rootStory
    while ($null -ne $story) {
      Write-Output $story
      $story = $story.NextStoryRange
    }
  }
}

function Get-ShapeTextRanges($Shapes) {
  foreach ($shape in $Shapes) {
    if ([int]$shape.Type -eq $MSO_SHAPE_GROUP) {
      Get-ShapeTextRanges $shape.GroupItems
    }
    try {
      if ($shape.TextFrame.HasText -ne 0) { Write-Output $shape.TextFrame.TextRange }
    } catch {
      # Pictures and other non-text shapes do not expose a usable TextFrame.
    }
  }
}

function Get-HeaderFooterRanges($Doc) {
  foreach ($section in $Doc.Sections) {
    foreach ($header in $section.Headers) {
      if ($header.Exists -and ([int]$section.Index -eq 1 -or -not $header.LinkToPrevious)) {
        Write-Output $header.Range
      }
    }
    foreach ($footer in $section.Footers) {
      if ($footer.Exists -and ([int]$section.Index -eq 1 -or -not $footer.LinkToPrevious)) {
        Write-Output $footer.Range
      }
    }
  }
}

function Get-FieldRanges($Doc) {
  Get-StoryRanges $Doc | Where-Object { [int]$_.StoryType -notin $HEADER_FOOTER_STORY_TYPES }
  Get-HeaderFooterRanges $Doc
  Get-ShapeTextRanges $Doc.Shapes
  foreach ($section in $Doc.Sections) {
    foreach ($header in $section.Headers) { Get-ShapeTextRanges $header.Shapes }
    foreach ($footer in $section.Footers) { Get-ShapeTextRanges $footer.Shapes }
  }
}

function Sync-TocEntryFonts($Doc) {
  foreach ($toc in $Doc.TablesOfContents) {
    foreach ($tocPara in $toc.Range.Paragraphs) {
      try {
        $styleObj = $tocPara.Range.Style
        if ($null -eq $styleObj) { continue }
        $styleName = [string]$styleObj.NameLocal
        if ($styleName -notmatch '^(?i)(?:toc|目录)\s*[1-9]$') { continue }
        $sf = $styleObj.Font
        $pf = $tocPara.Range.Font
        $fe = [string]$sf.NameFarEast
        $la = [string]$sf.NameAscii
        if ($fe -and -not $fe.StartsWith('+') -and [string]$pf.NameFarEast -ne $fe) {
          $pf.NameFarEast = $fe
          $script:syncAdjustments++
        }
        if ($la -and -not $la.StartsWith('+')) {
          $pf.NameAscii = $la
          $pf.NameOther = $la
        }
        $sz = [double]$sf.Size
        if ($sz -gt 0 -and $sz -lt $MAX_REASONABLE_FONT_PT) { $pf.Size = $sz }
        if ([int]$sf.Bold -ne $WD_UNDEFINED) { $pf.Bold = [int]$sf.Bold }
        if ([int]$sf.Italic -ne $WD_UNDEFINED) { $pf.Italic = [int]$sf.Italic }
        if ([int]$sf.Color -ne $WD_UNDEFINED) { $pf.Color = [int]$sf.Color }
      } catch {
        $script:syncWarnings += "TOC font sync warning: $($_.Exception.Message)"
      }
    }
  }
}

function Sync-HeaderFooterFieldFonts($Doc) {
  foreach ($story in @(Get-HeaderFooterRanges $Doc)) {
    foreach ($field in $story.Fields) {
      try {
        $code = [string]$field.Code.Text
        if ($code -notmatch '^\s*(?i)(?:STYLEREF|PAGE|NUMPAGES|SECTIONPAGES)\b') { continue }
        $pRange = $field.Result.Paragraphs.Item(1).Range
        $markFont = $pRange.Characters.Last.Font
        $fe = [string]$markFont.NameFarEast
        $la = [string]$markFont.NameAscii
        $sz = [double]$markFont.Size
        $hasExplicit = ($fe -and -not $fe.StartsWith('+')) -or ($la -and -not $la.StartsWith('+')) -or ($sz -gt 0 -and $sz -lt $MAX_REASONABLE_FONT_PT)
        if ($hasExplicit -and $code -match '^\s*(?i)STYLEREF\b' -and $code -notmatch '(?i)\\\*\s*CHARFORMAT\b') {
          $field.Code.Text = ($code.TrimEnd() + ' \* CHARFORMAT ')
          $script:syncAdjustments++
        }
        foreach ($targetFont in @($field.Code.Font, $field.Result.Font)) {
          if ($fe -and -not $fe.StartsWith('+')) { $targetFont.NameFarEast = $fe }
          if ($la -and -not $la.StartsWith('+')) {
            $targetFont.NameAscii = $la
            $targetFont.NameOther = $la
          }
          if ($sz -gt 0 -and $sz -lt $MAX_REASONABLE_FONT_PT) { $targetFont.Size = $sz }
          if ([int]$markFont.Bold -ne $WD_UNDEFINED) { $targetFont.Bold = [int]$markFont.Bold }
          if ([int]$markFont.Italic -ne $WD_UNDEFINED) { $targetFont.Italic = [int]$markFont.Italic }
          if ([int]$markFont.Color -ne $WD_UNDEFINED) { $targetFont.Color = [int]$markFont.Color }
        }
        if ($code -match '^\s*(?i)STYLEREF\b') {
          $null = $field.Update()
        }
      } catch {
        $script:syncWarnings += "Header/footer field font sync warning: $($_.Exception.Message)"
      }
    }
  }
}

try {
  if ($Backend -eq 'word') {
    $app = New-Object -ComObject Word.Application
    $created = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
      Where-Object { $existingIds -notcontains $_.Id })
    if ($created.Count -ne 1) { throw 'Cannot identify a dedicated new Word instance; no documents were opened' }
    $wordProcess = $created[0]
    $ownedProcesses = @($wordProcess)
    $owned = $true
    $identity = [ordered]@{
      name = [string]$app.Name; path = [string]$app.Path
      version = [string]$app.Version; build = [string]$app.Build
      executable = [string]$wordProcess.Path; pid = [int]$wordProcess.Id; owned = $owned
      started_ticks = $wordProcess.StartTime.ToUniversalTime().Ticks.ToString()
    }
    Write-Evidence ([IO.Path]::ChangeExtension($EvidencePath, '.identity.json')) $identity
    if ($identity.name -ne 'Microsoft Word' -or
        [IO.Path]::GetFileName($identity.executable) -ine 'WINWORD.EXE') {
      throw "Actual application is not Microsoft Word: $($identity.name) $($identity.executable)"
    }
  } else {
    $progUsed = $null
    foreach ($progId in @('KWPS.Application', 'WPS.Application', 'Word.Application')) {
      try {
        $candidate = New-Object -ComObject $progId
        $created = @(Get-Process -Name wps -ErrorAction SilentlyContinue |
          Where-Object { $existingIds -notcontains $_.Id })
        if ($progId -eq 'Word.Application' -and $created.Count -eq 0 -and $existingIds.Count -eq 0) {
          try { $candidate.Quit() } catch {}
          continue
        }
        $app = $candidate
        $progUsed = $progId
        if ($existingIds.Count -eq 0 -and $created.Count -ge 1) {
          $owned = $true
          $ownedProcesses = $created
        }
        break
      } catch {}
    }
    if ($null -eq $app) { throw 'No WPS Office COM server available' }
    $primaryProc = if ($ownedProcesses.Count -gt 0) { $ownedProcesses[0] } else {
      Get-Process -Name wps -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    $identity = [ordered]@{
      name = 'WPS Office'; prog_id = $progUsed; version = [string]$app.Version
      executable = if ($null -ne $primaryProc) { [string]$primaryProc.Path } else { 'wps.exe' }
      pid = if ($null -ne $primaryProc) { [int]$primaryProc.Id } else { 0 }
      owned = $owned
      owned_items = @($ownedProcesses | ForEach-Object {
        @{ pid = [int]$_.Id; executable = [string]$_.Path; started_ticks = $_.StartTime.ToUniversalTime().Ticks.ToString() }
      })
    }
    Write-Evidence ([IO.Path]::ChangeExtension($EvidencePath, '.identity.json')) $identity
  }

  $app.Visible = $false
  try { $app.DisplayAlerts = 0 } catch {}
  try { $app.AutomationSecurity = $MSO_AUTOMATION_SECURITY_FORCE_DISABLE } catch {}
  $readOnly = $Mode -ne 'refresh'
  $document = $app.Documents.Open($InputPath, $false, $readOnly, $false)
  if ($Mode -eq 'convert') {
    try {
      $document.SaveAs2($OutputPath, $WD_FORMAT_XML_DOCUMENT)
    } catch {
      $document.SaveAs($OutputPath, $WD_FORMAT_XML_DOCUMENT)
    }
    Write-Evidence $EvidencePath @{ application = $identity; mode = $Mode; status = 'completed' }
    return
  }
  $updates = @()
  # Bounded repeated passes let TOC and PAGE settle after pagination changes.
  $previousPageCount = -1
  $stable = $false
  for ($pass = 1; $pass -le $MAX_PAGINATION_PASSES; $pass++) {
    foreach ($toc in $document.TablesOfContents) { $null = $toc.Update() }
    foreach ($story in @(Get-FieldRanges $document)) {
      $updateResult = [int]$story.Fields.Update()
      if ($updateResult -ne 0) {
        $updates += @{ story = [int]$story.StoryType; first_error_index = $updateResult; pass = $pass }
      }
    }
    try { $document.Repaginate() } catch {}
    foreach ($toc in $document.TablesOfContents) {
      try { $null = $toc.UpdatePageNumbers() } catch {}
    }
    Sync-TocEntryFonts $document
    Sync-HeaderFooterFieldFonts $document
    $pageCount = [int]$document.ComputeStatistics($WD_STATISTIC_PAGES)
    if ($pageCount -eq $previousPageCount) { $stable = $true; break }
    $previousPageCount = $pageCount
  }
  Sync-TocEntryFonts $document
  Sync-HeaderFooterFieldFonts $document
  if ($Mode -eq 'refresh') { $document.Save() }
  if ($Mode -eq 'render') { $document.ExportAsFixedFormat($OutputPath, $WD_EXPORT_FORMAT_PDF) }

  $fields = @()
  $storyTypes = @()
  foreach ($story in @(Get-FieldRanges $document)) {
    $storyTypes += [int]$story.StoryType
    foreach ($field in $story.Fields) {
      $code = [string]$field.Code.Text
      $result = [string]$field.Result.Text
      $target = $null
      $targetExists = $null
      $targetLabel = $null
      $targetPage = $null
      if ($code -match '^\s*(?:REF|PAGEREF)\s+(?:"([^"]+)"|(\S+))') {
        $target = if ($Matches[1]) { $Matches[1] } else { $Matches[2] }
        $targetExists = [bool]$document.Bookmarks.Exists($target)
        if ($targetExists) {
          $targetRange = $document.Bookmarks.Item($target).Range
          $targetLabel = [string]$targetRange.ListFormat.ListString
          $targetRange.Collapse($WD_COLLAPSE_START)
          $targetPage = [int]$targetRange.Information($WD_INFO_ADJUSTED_PAGE)
        }
      }
      $fields += [ordered]@{
        code = $code; result = $result; story = [int]$story.StoryType
        start = [int]$field.Result.Start; page = [int]$field.Result.Information($WD_INFO_ADJUSTED_PAGE)
        target = $target; target_exists = $targetExists; target_label = $targetLabel
        target_page = $targetPage
      }
    }
  }
  $paragraphs = @()
  if ($Backend -eq 'word') {
    foreach ($paragraph in $document.Content.Paragraphs) {
      $range = $paragraph.Range
      $startRange = $range.Duplicate
      $startRange.Collapse($WD_COLLAPSE_START)
      $paragraphs += [ordered]@{
        text = [string]$range.Text; start = [int]$range.Start
        list_string = [string]$range.ListFormat.ListString
        list_value = [int]$range.ListFormat.ListValue
        list_level = [int]$range.ListFormat.ListLevelNumber
        outline = [int]$paragraph.OutlineLevel; style = [string]$range.Style.NameLocal
        page = [int]$startRange.Information($WD_INFO_ADJUSTED_PAGE)
        physical_page = [int]$startRange.Information($WD_INFO_PHYSICAL_PAGE)
        font_east_asia = [string]$range.Font.NameFarEast; font_latin = [string]$range.Font.NameAscii
        font_size_pt = [double]$range.Font.Size
        alignment = [int]$range.ParagraphFormat.Alignment
        first_line_indent_pt = [double]$range.ParagraphFormat.FirstLineIndent
        first_line_indent_chars = [double]$range.ParagraphFormat.CharacterUnitFirstLineIndent
        left_indent_pt = [double]$range.ParagraphFormat.LeftIndent
        left_indent_chars = [double]$range.ParagraphFormat.CharacterUnitLeftIndent
        line_spacing_rule = [int]$range.ParagraphFormat.LineSpacingRule
        line_spacing_pt = [double]$range.ParagraphFormat.LineSpacing
      }
    }
  }
  $sections = @()
  foreach ($section in $document.Sections) {
    $sectionStart = $section.Range.Duplicate
    $sectionStart.Collapse($WD_COLLAPSE_START)
    $sections += [ordered]@{
      index = [int]$section.Index; start = [int]$section.Range.Start
      adjusted_page = [int]$sectionStart.Information($WD_INFO_ADJUSTED_PAGE)
      physical_page = [int]$sectionStart.Information($WD_INFO_PHYSICAL_PAGE)
    }
  }
  $tocEntries = @()
  foreach ($toc in $document.TablesOfContents) {
    foreach ($link in $toc.Range.Hyperlinks) {
      $anchor = [string]$link.SubAddress
      if ($anchor -and $document.Bookmarks.Exists($anchor)) {
        $tocTarget = $document.Bookmarks.Item($anchor).Range.Duplicate
        $tocTarget.Collapse($WD_COLLAPSE_START)
        $tocEntries += @{ text = [string]$link.Range.Text; target = $anchor
          page = [int]$tocTarget.Information($WD_INFO_ADJUSTED_PAGE) }
      }
    }
  }
  $fonts = if ($Backend -eq 'word') { @($app.FontNames | ForEach-Object { [string]$_ }) } else { @() }
  Write-Evidence $EvidencePath ([ordered]@{
    application = $identity; mode = $Mode; status = 'completed'
    field_update_errors = @($updates); fields = @($fields); paragraphs = @($paragraphs)
    update_passes = $pass
    sections = @($sections); page_count = $pageCount; pagination_stable = $stable
    toc_entries = @($tocEntries); available_fonts = $fonts
    sync_adjustments = [int]$syncAdjustments; sync_warnings = @($syncWarnings)
    coverage = @{ story_types = @($storyTypes | Sort-Object -Unique)
      textboxes = 'Text-frame stories plus text-bearing shapes/groups in the body and headers/footers' }
  })
} finally {
  if ($null -ne $document) { try { $document.Close(0) } catch {} }
  if ($null -ne $app -and $owned) { try { $app.Quit() } catch {} }
  if ($Backend -eq 'wps' -and $owned) {
    foreach ($proc in $ownedProcesses) {
      Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
  }
  if ($null -ne $document) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
  if ($null -ne $app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
