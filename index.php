<?php
date_default_timezone_set('Asia/Taipei');
require_once __DIR__ . '/Classes/PHPExcel.php';

if (!empty($_FILES['file']['size'])) {
    $objPHPExcel = PHPExcel_IOFactory::load($_FILES['file']['tmp_name']);
    $fh = fopen($_FILES['file']['tmp_name'], 'w');
    $sheets = $objPHPExcel->getAllSheets();
    foreach ($sheets AS $sheet) {
        $rows = $sheet->toArray();
        foreach ($rows AS $row) {
            fputcsv($fh, $row);
        }
        unset($rows);
        unset($sheet);
    }
    unset($sheets);
    unset($objPHPExcel);
    fclose($fh);
    $length = filesize($_FILES['file']['tmp_name']);
    header("Content-Disposition: attachment; filename=output.csv;");
    header('Content-Type: application/octet-stream');
    header('Content-Length: ' . $length);
    readfile($_FILES['file']['tmp_name']);
    exit();
} else {
    ?><html>
        <head>
            <title>Excel 分頁合併工具</title>
        </head>
        <body>
            <h3>上傳的檔案系統會直接將所有分頁合併後轉換為 CSV 格式，產出的檔案依然需要做些整理</h3>
            <form method="post" enctype="multipart/form-data">
                <input type="file" name="file" />
                <input type="submit" value="上傳" />
            </form>
        </body>
    </html>
    <?php
}
