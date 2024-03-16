<?php

namespace Drupal\risley_export\Sheets;

use Drupal\Core\Extension\Extension;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Costom Modules sheet for modules spreadsheet.
 */
class CustomModulesSheet extends BaseSheet {

  /**
   * A list of custom modules across all sites.
   *
   * @var array|null
   */
  protected $modules;

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $this->modules = $this->getModulesAcrossSites('custom');

    $sheet->setTitle('カスタムモジュール | Custom modules');
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
    $modules = $this->getAllModules();

    foreach ($modules as $module) {
      $number = '=ROW()-1';
      $label = $module->info['name'] ?? '';
      $machineName = $module instanceof Extension ? $module->getName() : $module['machine_name'];
      $status = $this->getModuleStatus($module);
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
