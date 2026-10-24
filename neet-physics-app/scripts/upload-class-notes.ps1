# Upload class notes from Documents/physics 11/class notes into Supabase.
$ErrorActionPreference = "Stop"
$notesDir = "c:\Users\DC7586\Documents\physics 11\class notes"
$projectRef = "ctyclopfajckrlignqoq"
$base = "https://$projectRef.supabase.co"
$chapterId = "bcedd27e-c888-4dbd-9d31-563a73655d3e"
$teacherId = "b5c24a95-d99a-4cdf-a20e-8b1a664a043d"

$keys = npx supabase projects api-keys --project-ref $projectRef --output json | ConvertFrom-Json
$service = ($keys | Where-Object { $_.name -eq "service_role" } | Select-Object -First 1).api_key
if (-not $service) { throw "service_role key not found" }

$headers = @{ apikey = $service; Authorization = "Bearer $service" }
$jsonHeaders = @{
  apikey = $service
  Authorization = "Bearer $service"
  "Content-Type" = "application/json"
  Prefer = "return=representation"
}

$items = @(
  @{ n = 1; file = "Lecture_01_Trigonometry_Ratios_Identities-class notes.pdf"; title = "Trigonometric Ratios and Identities"; date = "2026-09-01"; material = "Lecture 01 Class Notes" },
  @{ n = 2; file = "Lecture_02_Trigonometry_ASTC_Height_Radian class notes.pdf"; title = "Height, ASTC and Radian"; date = "2026-09-02"; material = "Lecture 02 Class Notes" },
  @{ n = 3; file = "Lecture_03_Trigonometry_Feel_Compound_Small_Angle.pdf"; title = "Feel of Trigonometry and Small Angles"; date = "2026-09-03"; material = "Lecture 03 Class Notes" },
  @{ n = 4; file = "Lecture_04_ASTC_MaxMin_Phase_Phasor - class notes.pdf"; title = "ASTC, Maximum-Minimum and Phase"; date = "2026-09-04"; material = "Lecture 04 Class Notes" },
  @{ n = 5; file = "Lecture_05_MaxMin class notes.pdf"; title = "Phase, A.P., G.P. and Graph Basics"; date = "2026-09-07"; material = "Lecture 05 Class Notes" },
  @{ n = 6; file = "Lecture_06_AP_GP class notes.pdf"; title = "A.P., G.P. and Straight-Line Graphs"; date = "2026-09-08"; material = "Lecture 06 Class Notes" },
  @{ n = 7; file = "Lecture_07_Complete_GP_and_Straight_Line_Graphs class notes.pdf"; title = "Complete G.P. and Straight-Line Graphs"; date = "2026-09-09"; material = "Lecture 07 Class Notes" },
  @{ n = 8; file = "Lecture_08_Graph_Shapes_and_Physics_Applications class notes.pdf"; title = "Graph Shapes and Physics Applications"; date = "2026-09-10"; material = "Lecture 08 Class Notes" },
  @{ n = 9; file = "Lecture_09_Circle_Ellipse_Exponential_Differentiation class notes.pdf"; title = "Circle, Ellipse, Exponential and Differentiation"; date = "2026-09-11"; material = "Lecture 09 Class Notes" }
)

function Ensure-Class($item) {
  $existing = Invoke-RestMethod -Uri "$base/rest/v1/classes?select=id,status&lecture_number=eq.$($item.n)" -Headers $headers
  if ($existing -and $existing.Count -gt 0) {
    $classId = $existing[0].id
    if ($existing[0].status -ne "published") {
      Invoke-RestMethod -Method Patch -Uri "$base/rest/v1/classes?id=eq.$classId" -Headers $jsonHeaders -Body (@{
        status = "published"
        published_at = (Get-Date).ToUniversalTime().ToString("o")
      } | ConvertTo-Json) | Out-Null
    }
    return $classId
  }

  $created = Invoke-RestMethod -Method Post -Uri "$base/rest/v1/classes" -Headers $jsonHeaders -Body (@{
    chapter_id = $chapterId
    lecture_number = $item.n
    title = $item.title
    description = "Class notes for Lecture $($item.n)."
    class_date = $item.date
    duration_minutes = 90
    created_by = $teacherId
    status = "published"
    published_at = (Get-Date).ToUniversalTime().ToString("o")
  } | ConvertTo-Json)
  return $created[0].id
}

foreach ($item in $items) {
  $path = Join-Path $notesDir $item.file
  if (-not (Test-Path -LiteralPath $path)) {
    Write-Host "MISSING: $($item.file)"
    continue
  }

  $classId = Ensure-Class $item
  $safeName = "lecture-{0:D2}-class-notes.pdf" -f $item.n
  $storagePath = "classes/$classId/$safeName"
  $bytes = [System.IO.File]::ReadAllBytes($path)

  # Upsert file in storage
  $uploadHeaders = @{
    apikey = $service
    Authorization = "Bearer $service"
    "Content-Type" = "application/pdf"
    "x-upsert" = "true"
  }
  Invoke-RestMethod -Method Post -Uri "$base/storage/v1/object/class-materials/$storagePath" -Headers $uploadHeaders -Body $bytes | Out-Null

  $existingMaterial = Invoke-RestMethod -Uri "$base/rest/v1/materials?select=id&class_id=eq.$classId&kind=eq.class_notes" -Headers $headers
  $payload = @{
    class_id = $classId
    title = $item.material
    kind = "class_notes"
    storage_path = $storagePath
    mime_type = "application/pdf"
    file_size = $bytes.Length
    created_by = $teacherId
    status = "published"
    published_at = (Get-Date).ToUniversalTime().ToString("o")
  }

  if ($existingMaterial -and $existingMaterial.Count -gt 0) {
    $materialId = $existingMaterial[0].id
    Invoke-RestMethod -Method Patch -Uri "$base/rest/v1/materials?id=eq.$materialId" -Headers $jsonHeaders -Body ($payload | ConvertTo-Json) | Out-Null
    Write-Host "UPDATED L$($item.n): $($item.file) ($([math]::Round($bytes.Length/1KB)) KB)"
  } else {
    Invoke-RestMethod -Method Post -Uri "$base/rest/v1/materials" -Headers $jsonHeaders -Body ($payload | ConvertTo-Json) | Out-Null
    Write-Host "CREATED L$($item.n): $($item.file) ($([math]::Round($bytes.Length/1KB)) KB)"
  }
}

Write-Host "Done."
