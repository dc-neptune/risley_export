<?php

namespace Drupal\risley_export\Sheets;

use Drupal\Core\Extension\Extension;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Contrib Modules sheet for modules spreadsheet.
 */
class ContribModulesSheet extends BaseSheet {

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
    $this->modules = $this->getModulesAcrossSites('contrib');
    $sheet->setTitle('Contribモジュール | Contrib modules');
    $row = $this->setHeaders([
      [5, 30, 30, 10, 46],
      [
        "番号", "モジュール", "システム内部名称", "有効/無効", "備考",
      ],
      [
        "No.", "Module", "Machine Name", "ON/OFF", "Remarks",
      ],
    ]);

    $this->setRows($sheet, $row);

    $this->setStyle($sheet);

    $this->setStyleCenter('D:D');

    $this->setBorders();
  }

  /**
   * Sets rows on the Version sheet.
   */
  protected function setRows(Worksheet $sheet, int $row): int {
    $modules = $this->getAllModules();

    foreach ($modules as $module) {
      $number = '=ROW()-1';
      $label = $module->info['name'] ?? $module['info']['name'] ?? '';
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
