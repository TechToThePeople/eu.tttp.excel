<?php
  
function writeHTMLFile($fileName, &$header, &$rows, $titleHeader = NULL, $outputHeader = TRUE) {
  if ($outputHeader) {
    header('Content-language: en');
    CRM_Utils_System::download(CRM_Utils_String::munge($fileName),
      'application/vnd.ms-excel',
      CRM_Core_DAO::$_nullObject,
      'xls',
      FALSE
    );
    echo "<table><thead><tr>";
    foreach ($header as $field) {
      echo "<th>$field</th>";
    }
    echo "</tr></thead><tbody>";
  }
  $i = 0;
  $fields_cnt = count($header);

  foreach ($rows as $row) {
    $schema_insert = '';
    $colNo = 0;
    echo "\n<tr>";
    foreach ($row as $j => $value) {
      echo "<td>" . htmlentities($value, ENT_COMPAT, 'UTF-8') . "</td>";
    }
    echo "</tr>";
  }
//  echo "</tbody></table>";
}

function exportexcel_civicrm_export( $exportTempTable, $headerRows, $sqlColumns, $exportMode ) {

  $writeHeader = true;
  $offset = 0;
  $limit  = 10;

  $query = "SELECT * FROM $exportTempTable";
  require_once 'CRM/Core/Report/Excel.php';
  while ( 1 ) {
    $limitQuery = $query . " LIMIT $offset, $limit;";
    $dao = CRM_Core_DAO::executeQuery( $limitQuery );

    if ( $dao->N <= 0 ) {
      break;
    }

    $componentDetails = array( );
    while ( $dao->fetch( ) ) {
      $row = array( );
      foreach ( $sqlColumns as $column => $dontCare ) {
        $row[$column] = $dao->$column;
      }
      $componentDetails[] = $row;
    }
 
    //CRM_Core_Report_Excel::
    writeHTMLFile( CRM_Export_BAO_Export::getExportFileName('xls', $exportMode)." ".date("Y-m-d_H-i"), $headerRows,$componentDetails, null, $writeHeader );
    $writeHeader = false;
    $offset += $limit;
  }
  echo "</tbody></table>";

  CRM_Utils_System::civiExit( );
}

