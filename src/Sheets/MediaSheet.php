<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Media sheet for data spreadsheet.
 */
class MediaSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize(): void {
    $sheet = $this->sheet;
    $sheet->setTitle('メディア | Media');

    $sheet->fromArray([
      "No.", "Media", "日本語名", "システム内部名称\nMachine name", "説明\nDescription", "翻訳可\nMultilingual", "最大アップロードサイズ\nMaximum Upload Size", "言語設定\nLanguage Settings", "Remarks",
    ], NULL, 'A1');
    $row = 2;

    $row = $this->setEntities($sheet, $row);

    $this->setStyle($sheet);
    $this->setBorders();

    $sheet->getColumnDimension('A')->setWidth(5);
    $sheet->getColumnDimension('B')->setWidth(20);
    $sheet->getColumnDimension('C')->setWidth(10);
    $sheet->getColumnDimension('D')->setWidth(30);
    $sheet->getColumnDimension('E')->setWidth(35);
    $sheet->getColumnDimension('F')->setWidth(13);
    $sheet->getColumnDimension('G')->setWidth(15);
    $sheet->getColumnDimension('H')->setWidth(30);
    $sheet->getColumnDimension('I')->setWidth(30);

    $centerAlignmentStyle = [
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
    ];
    $sheet->getStyle('E:G')->applyFromArray($centerAlignmentStyle);
  }

  /**
   * Sets rows for entities on the Media sheet.
   */
  protected function setRows(Worksheet $sheet, string $entityCategory, string $entityTypeId, array $entities, int $row): int {
    foreach ($entities as $entity) {
      $entityLabel = $entity->label();
      // Usually the machine name of the content type.
      $bundle = $entity->id();
      $entityDescription = $this->translate($entity->getDescription());
      $multilingual = $this->buildCheck($this->canTranslate($entityTypeId, $bundle));
      $language = $this->getLanguageSettings($entityTypeId, $bundle);
      $japaneseLabel = $this->translate($entity, '');

      $fieldDefinitions = $this->entityFieldManager->getFieldDefinitions($entityTypeId, $bundle);
      $uploadSize = '';
      foreach ($fieldDefinitions as $fieldDefinition) {
        if (method_exists($fieldDefinition, 'get')) {
          $settings = $fieldDefinition->get('settings');
          if (is_array($settings) && isset($settings['max_filesize']) && !empty($settings['max_filesize'])) {
            $uploadSize = $settings['max_filesize'];
            break;
          }
        }
      }
      $this->setCell($sheet, 'A', $row, '=ROW()-1');
      // Or other types based on the field.
      $this->setCell($sheet, 'B', $row, $entityLabel);
      $this->setCell($sheet, 'C', $row, $japaneseLabel === $entityLabel ? '' : $japaneseLabel);
      $this->setCell($sheet, 'D', $row, $bundle);
      $this->setCell($sheet, 'E', $row, $entityDescription);
      $this->setCell($sheet, 'F', $row, $multilingual);
      $this->setCell($sheet, 'G', $row, $uploadSize);
      $this->setCell($sheet, 'H', $row, $language);

      $row++;
    }

    return $row;
  }

  /**
   * Sets entities.
   */
  private function setEntities(Worksheet $sheet, int $row, string $entityCategory = 'media_type', string $entityTypeId = 'media'): int {
    $entities = $this->entityTypeManager->getStorage($entityCategory)->loadMultiple();
    $row = $this->setRows($sheet, $entityCategory, $entityTypeId, $entities, $row);
    return $row;
  }

}
