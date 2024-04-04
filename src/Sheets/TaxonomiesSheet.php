<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Taxonomies sheet for data spreadsheet.
 */
class TaxonomiesSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $sheet->setTitle('タクソノミー | Taxonomies');

    $row = $this->setHeaders([
      [5, 30, 30, 35, 12, 15, 15, 30, 30],
      [
        "番号", "ターム", "システム内部名称", "説明", "翻訳可", "XMLサイトマップ", "ページ表示", "言語設定", "備考",
      ],
      [
        "No.", "Vocabulary", "Machine Name", "Description", "Multilingual", "XMLSiteMap", "Rabbit Hole", "Language Settings", "Remarks",
      ],
    ]);

    $row = $this->setEntities($sheet, $row);

    $this->setStyle($sheet);
    $this->setBorders();

    $centerAlignmentStyle = [
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
    ];
    $sheet->getStyle('E:G')->applyFromArray($centerAlignmentStyle);
  }

  /**
   * Sets rows for entities on the Taxonomies sheet.
   */
  protected function setRows(Worksheet $sheet, string $entityCategory, string $entityTypeId, array $entities, int $row):int {
    foreach ($entities as $entity) {
      $entityLabel = $entity->label();
      // Usually the machine name of the content type.
      $bundle = $entity->id();

      $entityDescription = $this->translate($entity->getDescription());

      $multilingual = $this->buildCheck($this->canTranslate($entityTypeId, $bundle));
      $language = $this->getLanguageSettings($entityTypeId, $bundle);
      $sitemap = $this->buildCheck($this->isShownOnSimpleSitemap($entityTypeId, $bundle));
      $rabbitHole = $this->getRabbitholeSetting($entityCategory, $bundle);

      $this->setCell($sheet, 'A', $row, '=ROW()-1');
      // Or other types based on the field.
      $this->setCell($sheet, 'B', $row, $entityLabel);
      $this->setCell($sheet, 'C', $row, $bundle);
      $this->setCell($sheet, 'D', $row, $entityDescription);
      $this->setCell($sheet, 'E', $row, $multilingual);
      $this->setCell($sheet, 'F', $row, $sitemap);
      $this->setCell($sheet, 'G', $row, $rabbitHole);
      $this->setCell($sheet, 'H', $row, $language);

      $row++;
    }

    return $row;
  }

  /**
   * Sets entities.
   */
  private function setEntities(Worksheet $sheet, int $row, string $entityCategory = 'taxonomy_vocabulary', string $entityTypeId = 'taxonomy_term'): int {
    $entities = $this->entityTypeManager->getStorage($entityCategory)->loadMultiple();
    $row = $this->setRows($sheet, $entityCategory, $entityTypeId, $entities, $row);
    return $row;
  }

}
