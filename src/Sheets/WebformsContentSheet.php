<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Content sheet for webforms spreadsheet.
 */
class WebformsContentSheet extends BaseSheet {

  /**
   * A list of webforms across all sites.
   *
   * @var array|null
   */
  protected $webforms;

  /**
   * A list of all nodes that contain a webform in the layout builder.
   *
   * @var array
   */
  protected $webformNodes;

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $this->webforms = $this->getWebformsAcrossSites();
    $this->webformNodes = $this->getNodesWithWebform();
    $sheet->setTitle('ウェブフォーム | Content');
    $row = $this->setHeaders([
      [
        5,
        30,
        30,
        30,
        60,
        12,
        15,
        60,
        15,
        15,
        15,
        15,
      ],
      ["番号", "Webform", "", "", "", "Multisite", "標準機能設定 Basic Settings", "", "", "", "", ""],
      ["",
        "ラベル",
        "タイトル",
        "システム内部名称",
        "説明",
        "適用サイト",
        "カテゴリ",
        "URLエイリアス",
        "Warn users about unsaved changes",
        "Disable client-side validation",
        "Enable preview page ",
        "Form method",
      ],
      ["No.",
        "Name",
        "Title",
        "Machine name",
        "Administrative description",
        "Apply site",
        "Category",
        "URL alias",
        "",
        "",
        "",
        "",
      ],
    ]);

    $this->merge(['A1:A2', 'B1:E1', 'G1:L1', 'I2:I3', 'J2:J3', 'K2:K3', 'L2:L3']);

    $this->setRows($sheet, $row);

    $this->setStyle();

    foreach (['I:L', 'F:F'] as $range) {
      $this->setStyleCenter($range);
    }

    $this->setBorders();

  }

  /**
   * Sets rows on the sheet.
   */
  protected function setRows(Worksheet $sheet, int $row): int {
    $webforms = $this->getAllWebforms();

    foreach ($webforms as $webform) {
      $cells = [
      // Number.
        '=ROW()-1',
      // Name.
        $this->translate($webform['title']),
      // Title.
        $webform['title'] ?? '',
      // Machine name.
        $webform['id'] ?? '',
      // Administrative description.
        $this->translate($webform['description'] ?? ''),
      // Apply site.
        $this->getWebformStatus($webform),
      // Category.
        implode(", ", $webform['categories'] ?? []),
      // URL alias.
        $this->getUrlAliasFor($webform),
      // Warn users about unsaved changes.
        $this->buildCheck($webform['settings']['form_unsaved']),
      // Disable client-side validation.
        $this->buildCheck($webform['settings']['form_novalidate']),
      // Enable preview page .
        $this->buildCheck(!empty($webform['settings']['preview'])),
      /*
       * Form method.
       * "" (post default)
       * "post" (post custom)
       * "get" (get custom)
       */
        ['' => 'POST', 'post' => 'POST', 'get' => 'GET'][$webform['settings']['form_method']],
      ];
      foreach ($cells as $i => $content) {
        $this->setCell($sheet, $this->intToCol($i), $row, $content);
      }
      $row++;
    }

    return $row;
  }

  /**
   * Gets all nodes containing a webform block in their block layout.
   */
  private function getNodesWithWebform():array {
    $webformNodes = [];

    $nodes = $this->entityTypeManager->getStorage('node')->loadMultiple();

    foreach ($nodes as $node) {
      if ($node->hasField('layout_builder__layout') && !$node->get('layout_builder__layout')->isEmpty()) {
        /** @var \Drupal\layout_builder\Field\LayoutSectionItemList $layout */
        $layout = $node->get('layout_builder__layout');
        $sections = $layout->getSections();
        foreach ($sections as $section) {
          $components = $section->getComponents();
          foreach ($components as $component) {
            /*
             * In drupal 10 and earlier, getConfiguration is
             * protected for some dumb reason.
             */
            $reflectionClass = new \ReflectionClass(get_class($component));
            $method = $reflectionClass->getMethod('getConfiguration');
            $method->setAccessible(TRUE);
            $configuration = $method->invoke($component);
            if (
              isset($configuration['provider']) &&
              $configuration['provider'] === 'webform' &&
              isset($configuration['webform_id'])
            ) {
              $webformId = $configuration['webform_id'];

              if (!isset($webformNodes[$webformId])) {
                $webformNodes[$webformId] = [];
              }

              $webformNodes[$webformId][] = $node;
              break 2;
            }
          }
        }
      }
    }
    return $webformNodes;
  }

  /**
   * Gets the url aliases for a given webform.
   */
  private function getUrlAliasFor(array $webform):string {
    $webformId = $webform['id'];

    if (!isset($this->webformNodes[$webformId])) {
      return $webform['_url_alias'] ?? '';
    }

    $nodes = $this->webformNodes[$webformId];
    $languages = array_keys($this->languageManager->getNativeLanguages());

    foreach ($nodes as $i => $node) {
      $path = $node->toUrl()->toString();
      $url_alias = $this->pathAliasManager->getAliasByPath($path);
      foreach ($languages as $language) {
        $url_alias = str_replace("/$language/", "/", $url_alias);
      }
      $nodes[$i] = $url_alias;
    }

    return implode("\n", $nodes);
  }

}
