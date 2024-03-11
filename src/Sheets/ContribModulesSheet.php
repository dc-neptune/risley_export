<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Contrib Modules sheet for modules spreadsheet.
 */
class ContribModulesSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $sheet->setTitle('Contribモジュール | Contrib modules');
    $headers = [
      "No.", "モジュール\nModule", "システム内部名称\nMachine Name", "ON/OFF", "備考\nRemarks",
    ];
    $sheet->fromArray($headers, NULL, 'A1');
    $row = 2;

    $this->setRows($sheet, $row);

    $this->setStyle($sheet);

    $centerAlignmentStyle = [
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
    ];
    $sheet->getStyle('D:D')->applyFromArray($centerAlignmentStyle);

    $this->setBorders();

    $sheet->getColumnDimension('A')->setWidth(5);
    $sheet->getColumnDimension('B')->setWidth(30);
    $sheet->getColumnDimension('C')->setWidth(30);
    $sheet->getColumnDimension('D')->setWidth(10);
    $sheet->getColumnDimension('E')->setWidth(46);
  }

  /**
   * Sets rows on the Version sheet.
   */
  protected function setRows(Worksheet $sheet, int $row): int {
    $modules = $this->getModules('contrib');

    foreach ($modules as $module) {
      $number = '=ROW()-1';
      $label = $module->info['name'] ?? '';
      $machineName = $module->getName();
      $status = $this->buildCheck($this->moduleHandler->moduleExists($machineName));
      $remark = $this->getModuleDescription($module);

      $this->setCell($sheet, 'A', $row, $number);
      $this->setCell($sheet, 'B', $row, $label);
      $this->setCell($sheet, 'C', $row, $machineName);
      $this->setCell($sheet, 'D', $row, $status);
      $this->setCell($sheet, 'E', $row, $remark);

      $row++;
    }

    return $row;
  }

}
