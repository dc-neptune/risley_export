<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Costom Modules sheet for modules spreadsheet.
 */
class MenusSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $sheet->setTitle('メニュー | Menus');
    // Run this stupid thing twice b/c too lazy to refactor setRows()
    $maxDepth = $this->findMaxDepth();
    $customFieldLabels = $this->getCustomMenuLinkFields('label');

    $settings = array_merge(
      [5, 40],
      array_fill(0, $maxDepth + 1, 30),
      [60, 15],
      array_fill(0, count($customFieldLabels), 15),
      [15, 60]
    );
    $japaneseHeaders = array_merge(
      ["番号", "メニュー"],
      array_fill(0, $maxDepth + 1, "メニュー項目"),
      ["ユーアーエル", "拡張表示"],
      $customFieldLabels,
      ["サイトマップ反映", "備考"]
    );
    $englishHeaders = array_merge(
      ["No.", "Menu"],
      array_fill(0, $maxDepth + 1, "Menu Item"),
      ["URL", "Show As Expanded"],
      $customFieldLabels,
      ["Show on Site Map", "Remarks"]
    );

    $row = $this->setHeaders([
      $settings,
      $japaneseHeaders,
      $englishHeaders,
    ]);

    // Loop through headers to merge each across rows 1 and 2.
    $j = 1;
    for ($i = 0; $i < count($englishHeaders); $i++) {
      if ($englishHeaders[$i] === "Menu Item") {
        $this->setCell($sheet, $this->intToCol($i), 2, "第" . $j++ . "階層");
      }
    }

    // Merge Menu items.
    $sheet->mergeCells("C1:" . $this->intToCol(2 + $maxDepth) . "1");

    $this->setRows($sheet, $row);

    $this->setStyle($sheet);

    $this->setBorders();

    $this->setStyleCenter($this->intToCol(3 + $maxDepth + 1) . ":" . $this->incrementColumn($sheet->getHighestDataColumn(1), -1));
  }

  /**
   * Recursively sets rows for the page.
   */
  private function setRow(array $binaries, int $row, int $col, array $columns, Worksheet $sheet, int $maxDepth):int {
    if (count($columns) > $maxDepth + 2) {
      return $row;
    }

    foreach ($binaries as [$item, $children]) {
      $nextColumns = array_merge($columns, [$item->getTitle()]);
      $currentColumns = array_merge(
            $nextColumns,
            array_fill(0, ($maxDepth + 2) - count($columns), ' '),
            [
              explode(':', $item->get('link')->uri)[1],
              $this->buildCheck((bool) $item->get('expanded')->value),
            ],
            array_map(function ($machineName) use ($item) {
              return $this->buildCheck($item->hasField($machineName) && !$item->get($machineName)->isEmpty());
            }, $this->getCustomMenuLinkFields()),
            [
              $this->buildCheck($this->isShownOnSiteMap($item->getMenuName())),
              $this->translate("menus.{$item->getMenuName()}.remarks.{$this->toKebabCase($item->getTitle())}", ''),
            ]
          );
      foreach ($currentColumns as $i => $column) {
        $this->setCell($sheet, $this->intToCol($i), $row, $column, $i > 0 && $i < 2 + $maxDepth);
      }
      $row++;

      $row = $this->setRow($children, $row, $col, $nextColumns, $sheet, $maxDepth);
    }
    return $row;
  }

  /**
   * Sets rows on the sheet.
   */
  protected function setRows(Worksheet $sheet, int $row): int {

    // Ignore account, admin, tools.
    $menus = array_diff_key(
      $this->entityTypeManager->getStorage('menu')->loadMultiple(),
      array_flip(['admin', 'tools', 'account'])
    );
    $maxDepth = $this->findMaxDepth();
    foreach ($menus as $machineName => $menu) {
      $binaries = $this->buildMenuTree($this->entityTypeManager->getStorage('menu_link_content')->loadByProperties(['menu_name' => $machineName]));
      $row = $this->setRow($binaries, $row, 0, ['=ROW()-1', "{$menu->label()} ({$machineName})"], $sheet, $maxDepth);
    }

    return $row;
  }

  /**
   * Builds a nested tree for sane navigations.
   *
   * @return array
   *   An array.
   */
  private function buildMenuTree(array $menuItems, string $parentId = ''): array {
    $tree = [];
    foreach ($menuItems as $item) {
      $_parentId = explode(':', $item->getParentId());
      $_parentId = end($_parentId);
      if ($parentId === $_parentId) {
        /** @var array{\Drupal\menu_link_content\Entity\MenuLinkContent, array<\Drupal\menu_link_content\Entity\MenuLinkContent>} $binary */
        $binary = [
          0 => $item,
          1 => $this->buildMenuTree($menuItems, $item->uuid()),
        ];
        $tree[] = $binary;
      }
    }
    return $tree;
  }

  /**
   * Gets the max depth of a menu.
   *
   * Unoptimized because refactoring into driver would be annoying.
   */
  private function findDepth(array $tree, int $currentDepth = 0): int {
    $maxDepth = $currentDepth;
    foreach ($tree as $node) {
      if (!empty($node[1])) {
        $subtreeDepth = $this->findDepth($node[1], $currentDepth + 1);
        if ($subtreeDepth > $maxDepth) {
          $maxDepth = $subtreeDepth;
        }
      }
    }
    return $maxDepth;
  }

  /**
   * Gets the max depth of a menu.
   *
   * Unoptimized because refactoring into driver would be annoying.
   */
  private function findMaxDepth(): int {
    $menus = array_diff_key(
      $this->entityTypeManager->getStorage('menu')->loadMultiple(),
      array_flip(['admin', 'tools', 'account'])
    );
    $maxDepth = 0;
    foreach ($menus as $machineName => $menu) {
      $binaries = $this->buildMenuTree($this->entityTypeManager->getStorage('menu_link_content')->loadByProperties(['menu_name' => $machineName]));
      $depth = $this->findDepth($binaries);
      if ($depth > $maxDepth) {
        $maxDepth = $depth;
      }
    }
    return $maxDepth;
  }

  /**
   * Gets all custom fields on menu links.
   */
  private function getCustomMenuLinkFields(string $ret = 'machineName'): array {
    $field_definitions = $this->entityFieldManager->getFieldDefinitions('menu_link_content', 'menu_link_content');
    $fields_array = [];
    foreach ($field_definitions as $field_name => $field_definition) {
      if (strpos($field_name, 'field_') === 0) {
        if ($ret === 'machineName') {
          $fields_array[] = $field_name;
        }
        elseif ($ret === 'label') {
          $fields_array[] = $field_definition->getLabel();
        }
      }
    }
    return $fields_array;
  }

  /**
   * Checks whether the menu is enabled on sitemap.
   */
  private function isShownOnSiteMap(string $menuName):bool {
    $plugins = $this->configFactory->get('sitemap.settings')->get('plugins');
    if (!isset($plugins) || !is_array($plugins)) {
      return FALSE;
    }

    $plugin = $plugins["menu:{$menuName}"] ?? NULL;
    if (!isset($plugin)) {
      return FALSE;
    }

    return !!$plugin['enabled'];
  }

}
