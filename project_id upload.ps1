$userProfile = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile)
$excelFile = Join-Path $userProfile "PCE Ltd\Process & Systems - General\BST Workspace\Database\Data Dump\Book1.xlsx"
$sheetName = "Sheet1"
$sqlConnectionString = "Server=PCE-SQL-DEV;Database=pce;Integrated Security=True;"

$data = Import-Excel -Path $excelFile -WorksheetName $sheetName

$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $sqlConnectionString
$connection.Open()

$command = $connection.CreateCommand()

foreach ($row in $data) {
    $project_concat = if ([string]::IsNullOrEmpty($row.project_concat)) { "NULL" } else { "'$($row.project_concat)'" }
    $project_number = if ([string]::IsNullOrEmpty($row.project_number)) { "NULL" } else { "'$($row.project_number)'" }
    $ynomia_type = if ([string]::IsNullOrEmpty($row.ynomia_type)) { "NULL" } else { "'$($row.ynomia_type)'" }
    $novade_site_id = if ([string]::IsNullOrEmpty($row.novade_site_id)) { "NULL" } else { "'$($row.novade_site_id)'" }
    $ynomia_tennant = if ([string]::IsNullOrEmpty($row.ynomia_tennant)) { "NULL" } else { "'$($row.ynomia_tennant)'" }
    $ynomia_project_code = if ([string]::IsNullOrEmpty($row.ynomia_project_code)) { "NULL" } else { "'$($row.ynomia_project_code)'" }
    $novade_product_project_id = if ([string]::IsNullOrEmpty($row.novade_product_project_id)) { "NULL" } else { "'$($row.novade_product_project_id)'" }

    $command.CommandText = @"
    INSERT INTO project_id (project_concat, project_number, ynomia_type, novade_site_id, ynomia_tennant, ynomia_project_code, novade_product_project_id)
    VALUES ($project_concat, $project_number, $ynomia_type, $novade_site_id, $ynomia_tennant, $ynomia_project_code, $novade_product_project_id)
"@
    $command.ExecuteNonQuery()
}

$connection.Close()
