param(
  [Parameter(Mandatory=$true)][string]$InputPath,
  [string]$OutputPath,
  [ValidateSet('inspect','render','refresh','convert')][string]$Mode = 'inspect',
  [Parameter(Mandatory=$true)][string]$EvidencePath
)
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
$owned = $false
$existingIds = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

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
    if ([int]$shape.Type -eq 6) {
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
  Get-StoryRanges $Doc | Where-Object { [int]$_.StoryType -notin @(6, 7, 8, 9, 10, 11) }
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
        if ($fe -and -not $fe.StartsWith('+')) { $pf.NameFarEast = $fe }
        if ($la -and -not $la.StartsWith('+')) {
          $pf.NameAscii = $la
          $pf.NameOther = $la
        }
        $sz = [double]$sf.Size
        if ($sz -gt 0 -and $sz -lt 1000) { $pf.Size = $sz }
        if ([int]$sf.Bold -ne 9999999) { $pf.Bold = [int]$sf.Bold }
        if ([int]$sf.Italic -ne 9999999) { $pf.Italic = [int]$sf.Italic }
        if ([int]$sf.Color -ne 9999999) { $pf.Color = [int]$sf.Color }
      } catch {}
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
        $hasExplicit = ($fe -and -not $fe.StartsWith('+')) -or ($la -and -not $la.StartsWith('+')) -or ($sz -gt 0 -and $sz -lt 1000)
        if ($hasExplicit -and $code -match '^\s*(?i)STYLEREF\b' -and $code -notmatch '(?i)\\\*\s*CHARFORMAT\b') {
          $field.Code.Text = ($code.TrimEnd() + ' \* CHARFORMAT ')
        }
        foreach ($targetFont in @($field.Code.Font, $field.Result.Font)) {
          if ($fe -and -not $fe.StartsWith('+')) { $targetFont.NameFarEast = $fe }
          if ($la -and -not $la.StartsWith('+')) {
            $targetFont.NameAscii = $la
            $targetFont.NameOther = $la
          }
          if ($sz -gt 0 -and $sz -lt 1000) { $targetFont.Size = $sz }
          if ([int]$markFont.Bold -ne 9999999) { $targetFont.Bold = [int]$markFont.Bold }
          if ([int]$markFont.Italic -ne 9999999) { $targetFont.Italic = [int]$markFont.Italic }
          if ([int]$markFont.Color -ne 9999999) { $targetFont.Color = [int]$markFont.Color }
        }
        if ($code -match '^\s*(?i)STYLEREF\b') {
          $null = $field.Update()
        }
      } catch {}
    }
  }
}

try {
  $word = New-Object -ComObject Word.Application
  $created = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $existingIds -notcontains $_.Id })
  if ($created.Count -ne 1) { throw 'Cannot identify a dedicated new Word instance; no documents were opened' }
  $wordProcess = $created[0]
  $ownedId = $wordProcess.Id
  $owned = $true
  $identity = [ordered]@{
    name = [string]$word.Name; path = [string]$word.Path
    version = [string]$word.Version; build = [string]$word.Build
    executable = [string]$wordProcess.Path; pid = [int]$ownedId; owned = $owned
    started_ticks = $wordProcess.StartTime.ToUniversalTime().Ticks.ToString()
  }
  Write-Evidence ([IO.Path]::ChangeExtension($EvidencePath, '.identity.json')) $identity
  if ($identity.name -ne 'Microsoft Word' -or
      [IO.Path]::GetFileName($identity.executable) -ine 'WINWORD.EXE') {
    throw "Actual application is not Microsoft Word: $($identity.name) $($identity.executable)"
  }
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = 3
  $readOnly = $Mode -ne 'refresh'
  $document = $word.Documents.Open($InputPath, $false, $readOnly, $false)
  if ($Mode -eq 'convert') {
    $document.SaveAs2($OutputPath, 12)
    Write-Evidence $EvidencePath @{ application = $identity; mode = $Mode; status = 'completed' }
    return
  }
  $updates = @()
  # Bounded repeated passes let TOC and PAGE settle after pagination changes.
  $previousPageCount = -1
  $stable = $false
  for ($pass = 1; $pass -le 3; $pass++) {
    foreach ($toc in $document.TablesOfContents) { $null = $toc.Update() }
    foreach ($story in @(Get-FieldRanges $document)) {
      $updateResult = [int]$story.Fields.Update()
      if ($updateResult -ne 0) {
        $updates += @{ story = [int]$story.StoryType; first_error_index = $updateResult; pass = $pass }
      }
    }
    $document.Repaginate()
    foreach ($toc in $document.TablesOfContents) { $null = $toc.UpdatePageNumbers() }
    Sync-TocEntryFonts $document
    Sync-HeaderFooterFieldFonts $document
    $pageCount = [int]$document.ComputeStatistics(2)
    if ($pageCount -eq $previousPageCount) { $stable = $true; break }
    $previousPageCount = $pageCount
  }
  Sync-TocEntryFonts $document
  Sync-HeaderFooterFieldFonts $document
  if ($Mode -eq 'refresh') { $document.Save() }
  if ($Mode -eq 'render') { $document.ExportAsFixedFormat($OutputPath, 17) }

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
          $targetRange.Collapse(1)
          $targetPage = [int]$targetRange.Information(1)
        }
      }
      $fields += [ordered]@{
        code = $code; result = $result; story = [int]$story.StoryType
        start = [int]$field.Result.Start; page = [int]$field.Result.Information(1)
        target = $target; target_exists = $targetExists; target_label = $targetLabel
        target_page = $targetPage
      }
    }
  }
  $paragraphs = @()
  foreach ($paragraph in $document.Content.Paragraphs) {
    $range = $paragraph.Range
    $startRange = $range.Duplicate
    $startRange.Collapse(1)
    $paragraphs += [ordered]@{
      text = [string]$range.Text; start = [int]$range.Start
      list_string = [string]$range.ListFormat.ListString
      list_value = [int]$range.ListFormat.ListValue
      list_level = [int]$range.ListFormat.ListLevelNumber
      outline = [int]$paragraph.OutlineLevel; style = [string]$range.Style.NameLocal
      page = [int]$startRange.Information(1); physical_page = [int]$startRange.Information(3)
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
  $sections = @()
  foreach ($section in $document.Sections) {
    $sectionStart = $section.Range.Duplicate
    $sectionStart.Collapse(1)
    $sections += [ordered]@{
      index = [int]$section.Index; start = [int]$section.Range.Start
      adjusted_page = [int]$sectionStart.Information(1)
      physical_page = [int]$sectionStart.Information(3)
    }
  }
  $tocEntries = @()
  foreach ($toc in $document.TablesOfContents) {
    foreach ($link in $toc.Range.Hyperlinks) {
      $anchor = [string]$link.SubAddress
      if ($anchor -and $document.Bookmarks.Exists($anchor)) {
        $tocTarget = $document.Bookmarks.Item($anchor).Range.Duplicate
        $tocTarget.Collapse(1)
        $tocEntries += @{ text = [string]$link.Range.Text; target = $anchor
          page = [int]$tocTarget.Information(1) }
      }
    }
  }
  Write-Evidence $EvidencePath ([ordered]@{
    application = $identity; mode = $Mode; status = 'completed'
    field_update_errors = @($updates); fields = @($fields); paragraphs = @($paragraphs)
    update_passes = $pass
    sections = @($sections); page_count = $pageCount; pagination_stable = $stable
    toc_entries = @($tocEntries); available_fonts = @($word.FontNames | ForEach-Object { [string]$_ })
    coverage = @{ story_types = @($storyTypes | Sort-Object -Unique)
      textboxes = 'Text-frame stories plus text-bearing shapes/groups in the body and headers/footers' }
  })
} finally {
  if ($null -ne $document) { try { $document.Close(0) } catch {} }
  if ($null -ne $word -and $owned) { try { $word.Quit() } catch {} }
  if ($null -ne $document) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
  if ($null -ne $word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
