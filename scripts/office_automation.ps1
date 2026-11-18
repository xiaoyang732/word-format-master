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

function Get-FieldRanges($Doc) {
  Get-StoryRanges $Doc
  Get-ShapeTextRanges $Doc.Shapes
  foreach ($section in $Doc.Sections) {
    foreach ($header in $section.Headers) { Get-ShapeTextRanges $header.Shapes }
    foreach ($footer in $section.Footers) { Get-ShapeTextRanges $footer.Shapes }
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
    $pageCount = [int]$document.ComputeStatistics(2)
    if ($pageCount -eq $previousPageCount) { $stable = $true; break }
    $previousPageCount = $pageCount
  }
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
