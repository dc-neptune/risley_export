<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Roles sheet for the modules spreadsheet.
 */
class RolesSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $sheet->setTitle('役割 | Roles');

    $row = $this->setHeaders([
      [5, 20, 20, 25, 25, 30, 40],
      [
        "番号", "ロール名（日本語）", "ロール名（英語）", "システム内部名称", "説明", "利用可能機能の例", "備考",
      ],
      [
        "No.", "Role Name (Japanese)", "Role Name (English)", "Machine Name", "Description", "Example Uses", "Remarks",
      ],
    ]);

    $this->setRows($sheet, $row);

    $this->setStyle($sheet);

    $this->setBorders();
  }

  /**
   * Sets rows on the Roles sheet.
   */
  protected function setRows(Worksheet $sheet, int $row): int {
    $roles = $this->getPermissionRoles();

    foreach ($roles as $role) {
      $number = '=ROW()-1';
      $japaneseLabel = $this->translate($role, '');
      $label = $role->get('label');
      $machine_name = $role->get('id');
      $description = $this->translate("roles.description.$machine_name", '');
      $uses = $this->translate("roles.uses.$machine_name", '');
      $remarks = $this->translate("roles.remarks.$machine_name", '');

      $this->setCell($sheet, 'A', $row, $number);
      $this->setCell($sheet, 'B', $row, $japaneseLabel !== $label ? $japaneseLabel : '');
      $this->setCell($sheet, 'C', $row, $label);
      $this->setCell($sheet, 'D', $row, $machine_name);
      $this->setCell($sheet, 'E', $row, $description);
      $this->setCell($sheet, 'F', $row, $uses);
      $this->setCell($sheet, 'G', $row, $remarks);

      $row++;
    }

    return $row;
  }

}
