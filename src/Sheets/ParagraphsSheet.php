<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Paragraphs sheet for data spreadsheet.
 */
class ParagraphsSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $sheet->setTitle('パラグラフ | Paragraphs');

    $sheet->fromArray([
      "No.", "Paragraph", "日本語名", "システム内部名称\nMachine name", "翻訳可\nMultilingual", "Remarks",
    ]);
    $row = $this->setHeaders([
      [5, 30, 30, 45, 13, 30],
      [
        "番号", "パラグラフ", "日本語名", "システム内部名称", "翻訳可", "備考",
      ],
      [
        "No.", "Paragraph", "Japanese Name", "Machine name", "Multilingual", "Remarks",
      ],
    ]);

    $this->setEntities($sheet, $row);

    $lastRow = $sheet->getHighestDataRow("A");
    $lastColumn = $sheet->getHighestDataColumn("1");

    $this->setCell($sheet, 'B', $lastRow + 2, 'パラグラフ共通設定');
    $this->setCell($sheet, 'B', $lastRow + 3, '設定項目');
    $this->setCell($sheet, 'C', $lastRow + 3, '設定値');
    $this->setCell($sheet, 'B', $lastRow + 4, '非公開のパラグラフを閲覧可能にする');
    $state = $this->canTranslateAll('paragraph') ? 'ON' : 'OFF';
    $this->setCell($sheet, 'C', $lastRow + 4, $state);
    $this->setCell($sheet, 'D', $lastRow + 4, 'パラグラフライブラリーアイテム再利用のためにURL有効にする。');

    $this->setCell($sheet, 'B', $lastRow + 6, 'パラグラフライブラリー項目');
    $this->setCell($sheet, 'B', $lastRow + 7, '設定項目');
    $this->setCell($sheet, 'C', $lastRow + 7, '設定値');
    $this->setCell($sheet, 'B', $lastRow + 8, '翻訳可能');
    $this->setStyleHeader([
      "B" . ($lastRow + 3) . ":C" . ($lastRow + 3),
      "B" . ($lastRow + 7) . ":C" . ($lastRow + 7),
    ]);
    $state = '!!TODO!!';
    $this->setCell($sheet, 'C', $lastRow + 8, $state);

    $this->setCell($sheet, 'B', $lastRow + 2, 'パラグラフ共通設定');
    $this->setStyle();
    $this->setBorders("A1", $lastColumn . $lastRow);
    $this->setBorders('B' . ($lastRow + 3), 'C' . ($lastRow + 4));
    $this->setBorders('B' . ($lastRow + 6), 'C' . ($lastRow + 8));

    $centerAlignmentStyle = [
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
    ];
    $sheet->getStyle('E:E')->applyFromArray($centerAlignmentStyle);

  }

  /**
   * Sets rows for entities on the Paragraphs sheet.
   */
  protected function setRows(Worksheet $sheet, string $entityTypeId, array $entities, int $row): int {
    foreach ($entities as $entity) {
      $entityLabel = $entity->label();
      // Usually the machine name of the content type.
      $bundle = $entity->id();
      $multilingual = $this->buildCheck($this->canTranslate($entityTypeId, $bundle));
      $japaneseLabel = $this->translate($entity, '');

      $this->setCell($sheet, 'A', $row, '=ROW()-1');
      // Or other types based on the field.
      $this->setCell($sheet, 'B', $row, $entityLabel);
      $this->setCell($sheet, 'C', $row, $japaneseLabel === $entityLabel ? '' : $japaneseLabel);
      $this->setCell($sheet, 'D', $row, $bundle);
      $this->setCell($sheet, 'E', $row, $multilingual);

      $row++;
    }

    return $row;
  }

  /**
   * Sets entities.
   */
  private function setEntities(Worksheet $sheet, int $row): void {
    $entities = $this->entityTypeManager->getStorage('paragraphs_type')->loadMultiple();
    $this->setRows($sheet, 'paragraph', $entities, $row);
  }

}
