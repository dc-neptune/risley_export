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

    $row = $this->setHeaders([
      [5, 12, 13.5, 30, 60],
      [[], ['value' => 'Project:'], [], [], []],
      [[], ['value' => "Title:"], ['value' => $this->translate($title)], [], ['value' => $title]],
      [[], [], [], [], []],
      ["No.", "Date", "Author", "Sheets", "Note"],
    ]);

    $this->merge(['C1:D1', 'C2:D2']);

    $this->setRows($sheet, $row);

    $this->setStyle();

    $this->setBorders("B1", "E2");
    $this->setBorders("A4");

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
  protected function setRows(Worksheet $sheet, int $row, int $rows = 1): void {
    $lastRow = $row + $rows;

    for (; $row < $lastRow; $row++) {
      $this->setCell($sheet, 'A', $row, '=ROW()-4');
      $this->setCell($sheet, 'B', $row, '');
      $this->setCell($sheet, 'C', $row, '');
      $this->setCell($sheet, 'D', $row, '');
      $this->setCell($sheet, 'E', $row, '');
    }
  }

}
