<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

$reader = new Xlsx();
$spreadsheet = $reader->load("th_district.xlsx");
$aData = null;

try {
    $aData = $spreadsheet->getSheet(1)->toArray();
} catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
    echo '<pre>', $e->getMessage(), '</pre>';
}

//echo '<pre>', print_r($aData), '</pre>';

if (is_array($aData) && count($aData) > 0) {
    $user = 'root';
    $password = 'root';
    $dsn = "mysql:host=localhost;port=3306;dbname=th_district;charset=UTF8";
    $pdo = null;

    try {
        $pdo = new PDO($dsn, $user, $password);
    } catch (PDOException $e) {
        echo $e->getMessage();
    }

    $sqlRegion = "INSERT INTO region (name, name_en) VALUES (?, ?) ON DUPLICATE KEY UPDATE name_en = VALUES (name_en), id = LAST_INSERT_ID(id)";
    $stmtRegion = $pdo->prepare($sqlRegion);

    $sqlProvince = "INSERT INTO province (name, name_en, region_id) VALUES (?, ?, ?) ON DUPLICATE KEY UPDATE name_en = VALUES (name_en), id = LAST_INSERT_ID(id)";
    $stmtProvince = $pdo->prepare($sqlProvince);

    $sqlDistrict = "INSERT INTO district (name, name_en, province_id) VALUES (?, ?, ?) ON DUPLICATE KEY UPDATE name_en = VALUES (name_en), id = LAST_INSERT_ID(id)";
    $stmtDistrict = $pdo->prepare($sqlDistrict);

    $sqlSubDistrict = "INSERT INTO sub_district (name, name_en, district_id) VALUES (?, ?, ?) ON DUPLICATE KEY UPDATE name_en = VALUES (name_en), id = LAST_INSERT_ID(id)";
    $stmtSubDistrict = $pdo->prepare($sqlSubDistrict);

    foreach ($aData as $idx => $row) {
        if ($idx > 0) { // row 0 is column header
            $stmtRegion->execute([$row[16], '']);
            $regionId = $pdo->lastInsertId();

            $stmtProvince->execute([$row[14], $row[15], $regionId]);
            $provinceId = $pdo->lastInsertId();

            $stmtDistrict->execute([$row[9], $row[10], $provinceId]);
            $districtId = $pdo->lastInsertId();

            $stmtSubDistrict->execute([$row[3], $row[4], $districtId]);
        }
    }

    echo 'DONE!!';
} else {
    echo 'Read excel data fail.';
}