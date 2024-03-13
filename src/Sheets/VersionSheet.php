<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Version sheet for various spreadsheets.
 */
class VersionSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $sheet->setTitle('Version');

    $title = explode('_', $this->options['filename']);
    $title = end($title);
    $title = explode('.', $title)[0];
    $title = ucfirst($title);
    $title = "Drupal " . $title . " Structure";

    $this->setCell($sheet, "B", 1, "Project:");
    $this->setCell($sheet, "B", 2, "Title:");
    $this->setCell($sheet, "C", 2, $this->translate($title));
    $this->setCell($sheet, "E", 2, $title);

    $headers = [
      "No.", "Date", "Author", "Sheets", "Note",
    ];
    $sheet->fromArray($headers, NULL, 'A4');
    $row = 5;

    $this->setRows($sheet, $row);

    $this->setStyle($sheet, 4);

    $this->setBorders("B1", "E2");
    $this->setBorders("A4");

    $sheet->getColumnDimension('A')->setWidth(5);
    $sheet->getColumnDimension('B')->setWidth(12);
    $sheet->getColumnDimension('C')->setWidth(13.5);
    $sheet->getColumnDimension('D')->setWidth(30);
    $sheet->getColumnDimension('E')->setWidth(60);

    $titleStyle = [
      'font' => [
        'bold' => TRUE,
      ],
      'alignment' => [
        'wrapText' => FALSE,
      ],
    ];
    $sheet->getStyle('A1:E5')->applyFromArray($titleStyle);
  }

  /**
   * Sets rows on the Version sheet.
   */
  protected function setRows(Worksheet $sheet, int $row, int $rows = 1): int {
    $lastRow = $row + $rows;

    for ($row; $row < $lastRow; $row++) {
      $this->setCell($sheet, 'A', $row, '=ROW()-4');
      $this->setCell($sheet, 'B', $row, '');
      $this->setCell($sheet, 'C', $row, '');
      $this->setCell($sheet, 'D', $row, '');
      $this->setCell($sheet, 'E', $row, '');
    }

    return $lastRow;
  }

}
