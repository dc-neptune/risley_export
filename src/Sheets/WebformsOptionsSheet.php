<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Webforms Options sheet for data spreadsheet.
 */
class WebformsOptionsSheet extends BaseSheet {

  /**
   * A list of webforms across all sites.
   *
   * @var array<mixed>|null
   */
  protected array|NULL $webforms;

  /**
   * A merged webform containing all fields across all sites.
   *
   * @var array<mixed>|null
   */
  public array|NULL $allWebforms;

  /**
   * Initializes the sheet.
   */
  protected function initialize(): void {
    $sheet = $this->sheet;
    $this->webforms = $this->getWebformFieldsAcrossSites();
    $sheet->setTitle('項目 | Fields');
    $row = $this->setHeaders([
      [5, 15, 15, 30, 25, 25, 11, 11, 11, 20, 30, 40, 15, 35],
      [
        "番号", "カテゴリ", "ウエブフォーム", "項目名（英語）",
        "システム内部名称\n（26文字以内）", "フィールドタイプ", "必須", "個数", "事前入力", "初期値", "プレースホルダー（英語）", "取りうる値/制限", "適用サイト", "備考",
      ],
      [
        "No.", "Category", "Webform", "Field Name (English)",
        "Machine Name\n(under 27 characters)", "Field Type", "Required", "Number", "Prepopulate", "Default Value", "Placeholder (English)", "Field Settings", "Sites", "Remarks",
      ],
    ]);

    // Content types.
    $this->setWebforms($sheet, $row, 'node_type', 'node');

    // $sheet->getRowDimension(1)->setRowHeight(48);
    $this->setStyle();

    $this->setStyleCenter(['G:I', 'M:M']);

    $this->setBorders();
  }

  /**
   * Gets all webform elements from all sites.
   */
  protected function getWebformFieldsAcrossSites():array|NULL {
    return array_reduce($this->sites, function ($result, $site) {
      // $command = "/opt/drupal/vendor/bin/drush $site ev 'echo json_encode(array_map(function(\$webform) { return \$webform->toArray(); }, \\Drupal::service(\"entity_type.manager\")->getStorage(\"webform\")->loadMultiple()));'";
      $siteObj = $this->siteAliasManager->get($site);

      if (!$siteObj) {
        return $result;
      }

      $uri = $siteObj->uri();
      if (empty($uri)) {
        return $result;
      }

      $command = <<<EOT
          /opt/drupal/vendor/bin/drush --uri="$uri" ev '
          if(!\\Drupal::service("entity_type.manager")->hasDefinition("webform")) return [];
          \$webforms = \\Drupal::service("entity_type.manager")->getStorage("webform")->loadMultiple();
          \$webformsArray = array_map(function(\$webform) {
              \$array = \$webform->getElementsInitialized();
              \$array["#title"] = \$webform->get("title");
              \$array["#categories"] = \$webform->get("categories");
              return \$array;
          }, \$webforms);
          echo json_encode(\$webformsArray);
          '
          EOT;

      $jsonWebforms = shell_exec($command);

      if (!is_string($jsonWebforms)) {
        return $result;
      }

      $webforms = json_decode($jsonWebforms, TRUE);

      if (!is_array($webforms)) {
        return $result;
      }

      $result[$site] = $webforms;
      return $result;
    }, []);
  }

  /**
   * Helper.
   *
   * Gets the multiple value for various elements.
   */
  private function getEmptyLabelForMultiple(string $type): string {
    $locale = [
      'fieldset' => '-',
      'webform_email_confirm' => '-',
      'webform_email' => '-',
      'email' => '-',
      'checkbox' => '-',
      'captcha' => '-',
      'webform_flexbox' => '-',
    ];

    return $locale[$type] ?? 1;
  }

  /**
   * Helper.
   *
   * Gets the default value for various elements.
   */
  private function getEmptyLabelForDefault(string $type): string {
    $locale = [
      'fieldset' => '-',
      'textfield' => '空欄',
      'radios' => '選択無し',
      'webform_email_confirm' => '空欄',
      'webform_email' => '空欄',
      'email' => '空欄',
      'checkbox' => '空欄',
      'captcha' => '空欄',
      'date' => '空欄',
      'webform_flexbox' => '空欄',
    ];

    return $locale[$type] ?? 'Webform自動入力';
  }

  /**
   * Helper.
   *
   * Gets the real children of element.
   */
  private function getChildren(array $array): array {
    return array_filter($array, function ($key) {
      return !str_starts_with($key, '#');
    }, ARRAY_FILTER_USE_KEY);
  }

  /**
   * Helper.
   *
   * Gets the #attributes of element.
   */
  private function getAttributes(array $array): array {
    return array_filter($array, function ($key) {
      return str_starts_with($key, '#');
    }, ARRAY_FILTER_USE_KEY);
  }

  /**
   * Helper.
   *
   * Gets the real children of element.
   */
  private function hasDescendant(array $array, $key): bool {
    $children = $this->getChildren($array);

    foreach ($children as $_key => $child) {
      if ( $key === $_key || $this->hasDescendant($child, $key)) return TRUE;
    }

    return FALSE;
  }

  /**
   * Helper.
   *
   * Gets a list of sites that have the field.
   */
  private function getSitesFor(string $webformName, string $elementName): string {
    $sites = [];

    foreach ($this->webforms as $site => $webforms) {
      if (!isset($webforms[$webformName])) continue;
      $array = $webforms[$webformName];
      if ($this->hasDescendant($array, $elementName)) {
        $sites[] = $site;
      }
    }

    return $this->buildSites($sites);
  }

  /**
   * Helper.
   *
   * Merges two trees together.
   */
  private function mergeArrays(array &$array1, array $array2, string $parentKey = ''): array {
    [$children1, $children2] = array_map(function ($array) {
      return $this->getChildren($array);
    }, [$array1, $array2]);

    /*
     *  Since the spec sheet expects all values to be the same across sites,
     *  it would be nice to at least warn if a value is different somewhere
     */
    [$attributes1, $attributes2] = array_map(function ($array) {
      return $this->getAttributes($array);
    }, [$array1, $array2]);
    foreach ($attributes2 as $key => $value) {
      if (!isset($attributes1[$key]) || $attributes1[$key] !== $value)
        $this->logger->notice(dt("Webform element key '!key' of '!parentKey' differs on various sites. Found '!value'", ['!parentKey' => $parentKey, '!key' => $key, '!value' => $value]));
    }

    if (!isset($children2)) {
      return $array1;
    }
    //
    //    var_dump('Merging ' . ($tree1['key'] ?? 'UNSET') . ' and ' . ($tree2['key'] ?? 'UNSET'));
    //    var_dump('Original ' . ($tree1['key'] ?? 'UNSET') . ' has ' . count($tree1['children']) . ' children');
    //    var_dump('New ' . ($tree2['key'] ?? 'UNSET') . ' has ' . count($tree2['children']) . ' children');
    //    if (isset($tree2['key']) && $tree2['key'] === 'regional_questions') {
    //      var_dump('Merging regional_questions');
    //      var_dump($tree2['children']);
    //    }.
    foreach ($children2 as $key => $child) {
      if (!isset($array1[$key])) {
        // var_dump('The original ' . ($tree1['key'] ?? 'UNSET') . ' needs the child ' . $child['key']);.
        $array1[$key] = $child;
      }
      else {
        $array1[$key] = $this->mergeArrays($array1[$key], $child, $key);
      }
    }

    return $array1;
  }

  /**
   * Sets entities in the Fields sheet.
   */
  private function setWebforms(Worksheet $sheet, int $row, string $entityCategory, string $entityTypeId): int {
    $sites = $this->webforms;
    // Build a master list of all webforms, each with all possible fields.
    $allWebforms = [];
    foreach ($sites as $site => $webforms) {
      foreach ($webforms as $id => $elements) {
        if (!isset($allWebforms[$id])) {
          $allWebforms[$id] = $elements;
          continue;
        }

        $this->mergeArrays($allWebforms[$id], $elements);
      }
    }

    uasort($allWebforms, function ($a, $b) {
      return strcmp($a['#categories'][0] ?? '', $b['#categories'][0] ?? '');
    });

    foreach ($allWebforms as $webformName => $allElements) {
      $title = $allElements['#title'];
      $category = $allElements['#categories'][0] ?? '';
      $row = $this->setElements($allElements, $row, $webformName, $title, $category);
    }
    return $row;
  }

  /**
   * Sets rows for a group of elements.
   */
  private function setElements(array $allElements, int $row, string $webformName, string $title = '', string $category = ''): int {
    foreach ($allElements as $machineName => $element) {
      if (str_starts_with($machineName, '#')) {
        continue;
      }

      $cells = [
        // Number.
        '=ROW()-2',
        // Category.
        $category,
        // Webform.
        $title,
        // Title.
        $element['#title'],
        // Machine name.
        $machineName,
        // Field type.
        $this->webformElementManager->getDefinitions()[$element['#type']]['label']?->__toString() ?? $element['#type'],
        // Required.
        $this->buildCheck(isset($element['#required']) && $element['#required'] === TRUE),
        // Number.
        $element['#multiple'] ?? $this->getEmptyLabelForMultiple($element['#type']),
        // Prepopulate.
        $this->buildCheck(isset($element['#prepopulate']) && $element['#prepopulate'] === TRUE),
        // Default value.
        $element['#empty_option'] ?? $this->getEmptyLabelForDefault($element['#type']),
        // Placeholder.
        $element['#placeholder'] ?? '-',
        // Field settings.
        '@todo',
        // Site.
        $this->getSitesFor($webformName, $machineName),
        // Remarks.
        '',
      ];
      foreach ($cells as $i => $content) {
        $this->setCell($this->sheet, $this->intToCol($i), $row, $content, in_array($i, [1, 2]));
      }

      $row++;

      $row = $this->setElements($element, $row, $webformName, $title, $category);
    }
    return $row;
  }

}
