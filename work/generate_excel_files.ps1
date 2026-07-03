$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$RequiredSheets = [ordered]@{
  "Master_CPL" = @("Kode CPL", "Rumusan CPL", "GA IABEE")
  "Master_IK" = @("Kode IK", "Rumusan IK", "Kode CPL")
  "Mapping_CPMK" = @("Semester", "Kode MK", "Mata Kuliah", "Kode CPMK", "Rumusan CPMK", "Kode IK", "Kode CPL", "Bobot CPMK")
  "Nilai_CPMK" = @("NIM", "Nama Mahasiswa", "Kode MK", "Mata Kuliah", "Kode CPMK", "Nilai")
  "Target" = @("Parameter", "Nilai")
}

function Escape-XmlValue($Value) {
  if ($null -eq $Value) { return "" }
  return [System.Security.SecurityElement]::Escape([string]$Value)
}

function Get-ColName([int]$Index) {
  $Name = ""
  $Number = $Index + 1
  while ($Number -gt 0) {
    $Remainder = ($Number - 1) % 26
    $Name = [char](65 + $Remainder) + $Name
    $Number = [math]::Floor(($Number - 1) / 26)
  }
  return $Name
}

function New-SheetXml($Rows) {
  $XmlRows = New-Object System.Collections.Generic.List[string]
  for ($r = 0; $r -lt $Rows.Count; $r++) {
    $Cells = New-Object System.Collections.Generic.List[string]
    $Row = $Rows[$r]
    for ($c = 0; $c -lt $Row.Count; $c++) {
      $Ref = "$(Get-ColName $c)$($r + 1)"
      $Value = $Row[$c]
      if ($null -eq $Value -or $Value -eq "") {
        $Cells.Add("<c r=`"$Ref`"/>")
      } elseif ($Value -is [int] -or $Value -is [double]) {
        $Cells.Add("<c r=`"$Ref`"><v>$Value</v></c>")
      } else {
        $Cells.Add("<c r=`"$Ref`" t=`"inlineStr`"><is><t>$(Escape-XmlValue $Value)</t></is></c>")
      }
    }
    $XmlRows.Add("<row r=`"$($r + 1)`">$($Cells -join '')</row>")
  }
  return "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheetData>$($XmlRows -join '')</sheetData></worksheet>"
}

function Add-ZipEntry($Zip, [string]$Name, [string]$Content) {
  $Entry = $Zip.CreateEntry($Name)
  $Stream = $Entry.Open()
  $Encoding = New-Object System.Text.UTF8Encoding $false
  $Writer = New-Object System.IO.StreamWriter $Stream, $Encoding
  $Writer.Write($Content)
  $Writer.Dispose()
}

function Write-Workbook([string]$Path, $Sheets) {
  if (Test-Path -LiteralPath $Path) {
    [System.IO.File]::Delete($Path)
  }
  $FileStream = [System.IO.File]::Create($Path)
  $Zip = New-Object System.IO.Compression.ZipArchive $FileStream, ([System.IO.Compression.ZipArchiveMode]::Create)

  $SheetOverrides = New-Object System.Collections.Generic.List[string]
  $WorkbookSheets = New-Object System.Collections.Generic.List[string]
  $WorkbookRels = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $Sheets.Count; $i++) {
    $SheetId = $i + 1
    $SheetName = Escape-XmlValue $Sheets[$i].Name
    $SheetOverrides.Add("<Override PartName=`"/xl/worksheets/sheet$SheetId.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
    $WorkbookSheets.Add("<sheet name=`"$SheetName`" sheetId=`"$SheetId`" r:id=`"rId$SheetId`"/>")
    $WorkbookRels.Add("<Relationship Id=`"rId$SheetId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$SheetId.xml`"/>")
  }
  $WorkbookRels.Add("<Relationship Id=`"rId$($Sheets.Count + 1)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>")

  Add-ZipEntry $Zip "[Content_Types].xml" "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Override PartName=`"/xl/workbook.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`"/>$($SheetOverrides -join '')<Override PartName=`"/xl/styles.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml`"/></Types>"
  Add-ZipEntry $Zip "_rels/.rels" "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument`" Target=`"xl/workbook.xml`"/></Relationships>"
  Add-ZipEntry $Zip "xl/workbook.xml" "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets>$($WorkbookSheets -join '')</sheets></workbook>"
  Add-ZipEntry $Zip "xl/_rels/workbook.xml.rels" "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$($WorkbookRels -join '')</Relationships>"
  Add-ZipEntry $Zip "xl/styles.xml" "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><styleSheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><fonts count=`"1`"><font><sz val=`"11`"/><name val=`"Calibri`"/></font></fonts><fills count=`"1`"><fill><patternFill patternType=`"none`"/></fill></fills><borders count=`"1`"><border/></borders><cellStyleXfs count=`"1`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`"/></cellStyleXfs><cellXfs count=`"1`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`" xfId=`"0`"/></cellXfs></styleSheet>"

  for ($i = 0; $i -lt $Sheets.Count; $i++) {
    Add-ZipEntry $Zip "xl/worksheets/sheet$($i + 1).xml" (New-SheetXml $Sheets[$i].Rows)
  }
  $Zip.Dispose()
  $FileStream.Dispose()
}

function New-Sheet([string]$Name, $Rows) {
  return [pscustomobject]@{ Name = $Name; Rows = $Rows }
}

function New-RowsFromObjects($Columns, $Objects) {
  $Rows = New-Object System.Collections.Generic.List[object]
  $Rows.Add([object[]]$Columns)
  foreach ($Object in $Objects) {
    $Row = New-Object System.Collections.Generic.List[object]
    foreach ($Column in $Columns) {
      $Row.Add($Object[$Column])
    }
    $Rows.Add([object[]]$Row)
  }
  return $Rows
}

$TemplateSheets = New-Object System.Collections.Generic.List[object]
foreach ($SheetName in $RequiredSheets.Keys) {
  $Rows = New-Object System.Collections.Generic.List[object]
  $Rows.Add([object[]]$RequiredSheets[$SheetName])
  $TemplateSheets.Add((New-Sheet $SheetName $Rows))
}

$MasterCpl = @(
  @{ "Kode CPL" = "CPL1"; "Rumusan CPL" = "Mampu menerapkan matematika, sains, dan konsep teknik elektro terapan."; "GA IABEE" = "GA1" },
  @{ "Kode CPL" = "CPL2"; "Rumusan CPL" = "Mampu menganalisis dan menyelesaikan masalah teknik elektro well-defined."; "GA IABEE" = "GA2" },
  @{ "Kode CPL" = "CPL3"; "Rumusan CPL" = "Mampu menggunakan instrumen, perangkat lunak, dan teknologi modern."; "GA IABEE" = "GA5" }
)
$MasterIk = @(
  @{ "Kode IK" = "IK1.1"; "Rumusan IK" = "Menggunakan konsep matematika untuk rangkaian listrik dasar."; "Kode CPL" = "CPL1" },
  @{ "Kode IK" = "IK1.2"; "Rumusan IK" = "Menerapkan konsep sains dan elektro dalam tugas terapan."; "Kode CPL" = "CPL1" },
  @{ "Kode IK" = "IK2.1"; "Rumusan IK" = "Mengidentifikasi gejala gangguan pada sistem elektronika."; "Kode CPL" = "CPL2" },
  @{ "Kode IK" = "IK2.2"; "Rumusan IK" = "Menentukan alternatif solusi troubleshooting."; "Kode CPL" = "CPL2" },
  @{ "Kode IK" = "IK3.1"; "Rumusan IK" = "Mengoperasikan instrumen pengukuran elektronika."; "Kode CPL" = "CPL3" },
  @{ "Kode IK" = "IK3.2"; "Rumusan IK" = "Menggunakan software simulasi dan analisis rangkaian."; "Kode CPL" = "CPL3" }
)
$Mapping = @(
  @{ "Semester" = 1; "Kode MK" = "TE101"; "Mata Kuliah" = "Rangkaian Listrik"; "Kode CPMK" = "CPMK1"; "Rumusan CPMK" = "Mahasiswa mampu menghitung parameter rangkaian DC."; "Kode IK" = "IK1.1"; "Kode CPL" = "CPL1"; "Bobot CPMK" = 2 },
  @{ "Semester" = 1; "Kode MK" = "TE102"; "Mata Kuliah" = "Fisika Terapan"; "Kode CPMK" = "CPMK2"; "Rumusan CPMK" = "Mahasiswa mampu menerapkan konsep fisika pada sistem elektro."; "Kode IK" = "IK1.2"; "Kode CPL" = "CPL1"; "Bobot CPMK" = 1 },
  @{ "Semester" = 2; "Kode MK" = "TE201"; "Mata Kuliah" = "Dasar Elektronika"; "Kode CPMK" = "CPMK3"; "Rumusan CPMK" = "Mahasiswa mampu mengidentifikasi gangguan komponen elektronika."; "Kode IK" = "IK2.1"; "Kode CPL" = "CPL2"; "Bobot CPMK" = 1 },
  @{ "Semester" = 2; "Kode MK" = "TE202"; "Mata Kuliah" = "Troubleshooting Elektronika"; "Kode CPMK" = "CPMK4"; "Rumusan CPMK" = "Mahasiswa mampu memilih solusi troubleshooting."; "Kode IK" = "IK2.2"; "Kode CPL" = "CPL2"; "Bobot CPMK" = 2 },
  @{ "Semester" = 3; "Kode MK" = "TE301"; "Mata Kuliah" = "Pengukuran Elektronika"; "Kode CPMK" = "CPMK5"; "Rumusan CPMK" = "Mahasiswa mampu menggunakan osiloskop dan multimeter."; "Kode IK" = "IK3.1"; "Kode CPL" = "CPL3"; "Bobot CPMK" = 1 },
  @{ "Semester" = 3; "Kode MK" = "TE302"; "Mata Kuliah" = "Simulasi Rangkaian"; "Kode CPMK" = "CPMK6"; "Rumusan CPMK" = "Mahasiswa mampu menggunakan software simulasi rangkaian."; "Kode IK" = "IK3.2"; "Kode CPL" = "CPL3"; "Bobot CPMK" = 1 }
)
$Students = @(
  @("2303001", "Andi Saputra"), @("2303002", "Budi Santoso"), @("2303003", "Citra Lestari"),
  @("2303004", "Dewi Anggraini"), @("2303005", "Eko Prasetyo"), @("2303006", "Fitri Nabila"),
  @("2303007", "Gilang Ramadhan"), @("2303008", "Hana Maharani"), @("2303009", "Indra Wijaya"),
  @("2303010", "Joko Permana")
)
$Scores = [ordered]@{
  "CPMK1" = @(82, 76, 69, 88, 73, 65, 91, 70, 67, 80)
  "CPMK2" = @(75, 72, 68, 81, 66, 62, 85, 74, 69, 78)
  "CPMK3" = @(71, 64, 60, 78, 69, 55, 82, 67, 63, 74)
  "CPMK4" = @(68, 66, 58, 75, 62, 57, 79, 65, 61, 72)
  "CPMK5" = @(88, 84, 76, 90, 79, 72, 92, 81, 77, 86)
  "CPMK6" = @(80, 74, 70, 86, 71, 68, 89, 75, 73, 82)
}
$MappingByCpmk = @{}
foreach ($Row in $Mapping) { $MappingByCpmk[$Row["Kode CPMK"]] = $Row }
$Nilai = New-Object System.Collections.Generic.List[object]
foreach ($Cpmk in $Scores.Keys) {
  for ($i = 0; $i -lt $Students.Count; $i++) {
    $Nilai.Add(@{
      "NIM" = $Students[$i][0]
      "Nama Mahasiswa" = $Students[$i][1]
      "Kode MK" = $MappingByCpmk[$Cpmk]["Kode MK"]
      "Mata Kuliah" = $MappingByCpmk[$Cpmk]["Mata Kuliah"]
      "Kode CPMK" = $Cpmk
      "Nilai" = $Scores[$Cpmk][$i]
    })
  }
}
$Target = @(
  @{ "Parameter" = "Batas Nilai Minimum"; "Nilai" = 70 },
  @{ "Parameter" = "Target Ketercapaian CPL"; "Nilai" = 70 }
)

$Objects = @{
  "Master_CPL" = $MasterCpl
  "Master_IK" = $MasterIk
  "Mapping_CPMK" = $Mapping
  "Nilai_CPMK" = $Nilai
  "Target" = $Target
}
$SampleSheets = New-Object System.Collections.Generic.List[object]
foreach ($SheetName in $RequiredSheets.Keys) {
  $SampleSheets.Add((New-Sheet $SheetName (New-RowsFromObjects $RequiredSheets[$SheetName] $Objects[$SheetName])))
}

Write-Workbook (Join-Path $Root "template_input.xlsx") $TemplateSheets
Write-Workbook (Join-Path $Root "sample_data.xlsx") $SampleSheets
Write-Output "created template_input.xlsx and sample_data.xlsx"
