<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Content sheet for data spreadsheet.
 */
class ContentSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize(): void {
    $sheet = $this->sheet;
    $sheet->setTitle('コンテンツ | Content');

    $row = $this->setHeaders([
      [5, 24, 22.5, 35, 12, 12, 12, 12, 20, 12, 12, 12, 12, 12, 30, 12, 12, 12, 15, 8, 8],
      [
        "番号", "コンテンツタイプ　Content Type", "", "", "デフォルト設定 Default Settings", "", "", "", "標準機能設定 Basic Settings", "", "", "", "", "", "", "拡張機能設定 Extended Settings", "", "", "", "カスタム機能設定", "",
      ],
      [
        "", "ラベル", "システム内部名称", "説明", "掲載", "フロントページへ掲載", "リスト上部に固定", "新しいリビジョンの作成", "メニューリンクを生成", "URLエイリアス", "投稿前にプレビュー", "翻訳可", "ワークフロー適用", "作成者と日付情報を表示", "言語設定", "メタタグ設定", "スケジュール設定", "XMLサイトマップ", "ページ表示", "", "",
      ],
      [
        "No.", "Label", "Machine Name", "Description", "Published", "Promoted to front page", "Sticky at top of lists", "Create new revision", "Menu", "URL alias", "Preview before submitting", "Multilingual", "Workflow", "Display author and date information", "Language settings", "Metatag", "Scheduler", "XMLsitemap", "Rabbit Hole", "", "",
      ],
    ]);

    $this->setEntities($sheet, $row);

    $this->setStyle();
    $this->setBorders();

    $this->setStyleCenter('E:T');
    $this->merge(['A1:A2', 'B1:D1', 'E1:H1', 'I1:O1', 'P1:S1', 'T1:U1']);
  }

  /**
   * Sets rows for entities.
   */
  protected function setRows(Worksheet $sheet, string $entityCategory, array $entities, int $row): int {
    foreach ($entities as $entity) {
      $entityLabel = $entity->label();
      // Usually the machine name of the content type.
      $bundle = $entity->id();

      $entityDescription = $entity->getDescription();

      $published = $this->buildCheck($this->getDefaultValue($this->entityFieldManager->getFieldDefinitions('node', $bundle)['status']));
      $promoted = $this->buildCheck($this->getDefaultValue($this->entityFieldManager->getFieldDefinitions('node', $bundle)['promote']));
      $stickied = $this->buildCheck($this->getDefaultValue($this->entityFieldManager->getFieldDefinitions('node', $bundle)['sticky']));
      $created = $this->buildCheck($entity->get('new_revision'));

      $menu = $this->getEnabledMenus($entity);
      $urlAlias = $this->buildCheck($this->hasPathautoPattern($bundle));
      // Preview Mode can be 0, 1, or 2.
      $preview = $this->buildCheck($entity->get('preview_mode') !== 0);
      $multilingual = $this->buildCheck($this->canTranslate('node', $bundle));
      $workflow = $this->buildCheck($this->hasWorkflow('node', $bundle));
      $displayAuthor = $this->buildCheck($entity->get('display_submitted'));
      $language = $this->getLanguageSettings('node', $bundle);
      $metatag = '@TODO';
      $scheduler = $this->buildCheck(in_array('scheduler', $entity->get('dependencies')['module'] ?? []));
      $sitemap = $this->buildCheck($this->isShownOnSimpleSitemap('node', $bundle));
      $rabbitHole = $this->getRabbitholeSetting($entityCategory, $bundle);

      $this->setCell($sheet, 'A', $row, '=ROW()-3');
      // Or other types based on the field.
      $this->setCell($sheet, 'B', $row, $entityLabel);
      $this->setCell($sheet, 'C', $row, $bundle);
      $this->setCell($sheet, 'D', $row, $entityDescription);
      $this->setCell($sheet, 'E', $row, $published);
      $this->setCell($sheet, 'F', $row, $promoted);
      $this->setCell($sheet, 'G', $row, $stickied);
      $this->setCell($sheet, 'H', $row, $created);
      $this->setCell($sheet, 'I', $row, $menu);
      $this->setCell($sheet, 'J', $row, $urlAlias);
      $this->setCell($sheet, 'K', $row, $preview);
      $this->setCell($sheet, 'L', $row, $multilingual);
      $this->setCell($sheet, 'M', $row, $workflow);
      $this->setCell($sheet, 'N', $row, $displayAuthor);
      $this->setCell($sheet, 'O', $row, $language);
      $this->setCell($sheet, 'P', $row, $metatag);
      $this->setCell($sheet, 'Q', $row, $scheduler);
      $this->setCell($sheet, 'R', $row, $sitemap);
      $this->setCell($sheet, 'S', $row, $rabbitHole);

      $row++;
    }

    return $row;
  }

  /**
   * Sets entities.
   */
  private function setEntities(Worksheet $sheet, int $row, string $entityCategory = 'node_type'): int {
    $entities = $this->entityTypeManager->getStorage($entityCategory)->loadMultiple();
    $row = $this->setRows($sheet, $entityCategory, $entities, $row);
    return $row;
  }

}
