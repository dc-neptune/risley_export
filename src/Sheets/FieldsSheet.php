<?php

namespace Drupal\risley_export\Sheets;

use Drupal\Core\Datetime\DrupalDateTime;
use Drupal\Core\Entity\EntityInterface;
use Drupal\Core\Field\FieldConfigInterface;
use Drupal\Core\Field\FieldDefinitionInterface;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Fields sheet for data spreadsheet.
 */
class FieldsSheet extends BaseSheet {
  /**
   * Format field settings.
   *
   * @var array<array[]>
   */
  protected array $fieldformatter;

  /**
   * Initializes the sheet.
   */
  protected function initialize(): void {
    $this->fieldformatter = [
      'link' => [
        'title' => [
          [
            'compare' => fn($arg) => TRUE,
            'return' => fn($arg) => "リンクテキストを許可: " . ['無効', '任意', '必須'][$arg],
          ],
        ],
        'link_type' => [
          [
            'compare' => fn($arg) => TRUE,
            'return' => fn($arg) => "リンクテキストを許可: " . (['内部リンクのみ', '外部リンクのみ'][$arg] ?? '内部と外部リンクの両方'),
          ],
        ],
      ],
      'link_destination' => [
        'target_options' => [
          [
            'compare' => fn($arg) => empty(array_values($arg)),
            'return' => fn($arg) => "ファイルディレクトリー: $arg",
          ],
          [
            'compare' => fn($arg) => !empty(array_values($arg)),
            'return' => fn($arg) => "参照リンクタイプ: " . implode(", ", array_filter(array_map(fn($key) => ['self' => 'このページ', 'field_document_file' => 'ドキュメントファイル', 'field_target_url' => '指定URL', 'nolink' => 'リンク無し'][$key] ?? NULL, array_keys($arg)), fn($key) => !empty($key))),
          ],
        ],
      ],
      'file' => [
        'file_directory' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "ファイルディレクトリー: $arg",
          ],
        ],
        'file_extensions' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "許可拡張子: " . implode(", ", explode(" ", $arg)),
          ],
        ],
        'max_filesize' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "サイズ上限: $arg",
          ],
        ],
        'description_field' => [
          [
            'compare' => fn($arg) => $arg === TRUE,
            'return' => fn($arg) => "説明の入力欄を表示する",
          ],
        ],
        'display_field' => [
          [
            'compare' => fn($arg) => $arg === TRUE,
            'return' => fn($arg) => " 非表示を可能にするチェックボックスを追加する",
          ],
        ],
        'display_default' => [
          [
            'compare' => fn($arg, $settings) => $settings['display_field'] === TRUE && $arg === TRUE,
            'return' => fn($arg) => "デフォルト: 表示",
          ],
          [
            'compare' => fn($arg, $settings) => $settings['display_field'] === TRUE && $arg === FALSE,
            'return' => fn($arg) => "デフォルト: 非表示",
          ],
        ],
      ],
      'image' => [
        'file_directory' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "ファイルディレクトリー: $arg",
          ],
        ],
        'file_extensions' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "許可拡張子: " . implode(", ", explode(" ", $arg)),
          ],
        ],
        'max_resolution' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "最大解像度: $arg",
          ],
        ],
        'min_resolution' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "最小解像度: $arg",
          ],
        ],
        'alt_field' => [
          [
            'compare' => fn($arg) => $arg === FALSE,
            'return' => fn($arg) => 'altの入力欄を表示しない',
          ],
        ],
        'alt_field_required' => [
          [
            'compare' => fn($arg, $settings) => $settings['alt_field'] === TRUE && $arg === TRUE,
            'return' => fn($arg) => 'altを必須にする',
          ],
        ],
        'title_field' => [
          [
            'compare' => fn($arg) => $arg === TRUE,
            'return' => fn($arg) => "タイトルの入力欄を表示する",
          ],
        ],
        'title_field_required' => [
          [
            'compare' => fn($arg, $settings) => $settings['alt_field'] === TRUE && $arg === TRUE,
            'return' => fn($arg) => 'タイトルを必須にする',
          ],
        ],
        'max_filesize' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "サイズ上限: $arg",
          ],
        ],
        'description_field' => [
          [
            'compare' => fn($arg) => $arg === TRUE,
            'return' => fn($arg) => "説明の入力欄を表示する",
          ],
        ],
        'display_field' => [
          [
            'compare' => fn($arg) => $arg === TRUE,
            'return' => fn($arg) => " 非表示を可能にするチェックボックスを追加する",
          ],
        ],
        'display_default' => [
          [
            'compare' => fn($arg, $settings) => $settings['display_field'] === TRUE && $arg === TRUE,
            'return' => fn($arg) => $arg ? "デフォルト: 表示" : "デフォルト: 非表示",
          ],
        ],
      ],

      'svg_image_field' => [
        'file_directory' => [
                [
                  'compare' => fn($arg) => !empty($arg),
                  'return' => fn($arg) => "ファイルディレクトリー: $arg",
                ],
        ],
        'file_extensions' => [
                [
                  'compare' => fn($arg) => !empty($arg),
                  'return' => fn($arg) => "許可拡張子: " . implode(", ", explode(" ", $arg)),
                ],
        ],
        'alt_field' => [
                [
                  'compare' => fn($arg) => $arg === FALSE,
                  'return' => fn($arg) => 'altの入力欄を表示しない',
                ],
        ],
        'alt_field_required' => [
                [
                  'compare' => fn($arg, $settings) => $settings['alt_field'] === TRUE && $arg === TRUE,
                  'return' => fn($arg) => 'altを必須にする',
                ],
        ],
        'title_field' => [
                [
                  'compare' => fn($arg) => $arg === TRUE,
                  'return' => fn($arg) => "タイトルの入力欄を表示する",
                ],
        ],
        'title_field_required' => [
                [
                  'compare' => fn($arg, $settings) => $settings['alt_field'] === TRUE && $arg === TRUE,
                  'return' => fn($arg) => 'タイトルを必須にする',
                ],
        ],
        'max_filesize' => [
                [
                  'compare' => fn($arg) => !empty($arg),
                  'return' => fn($arg) => "サイズ上限: $arg",
                ],
        ],
        'description_field' => [
                [
                  'compare' => fn($arg) => $arg === TRUE,
                  'return' => fn($arg) => "説明の入力欄を表示する",
                ],
        ],
        'display_field' => [
                [
                  'compare' => fn($arg) => $arg === TRUE,
                  'return' => fn($arg) => " 非表示を可能にするチェックボックスを追加する",
                ],
        ],
        'display_default' => [
                [
                  'compare' => fn($arg, $settings) => $settings['display_field'] === TRUE && $arg === TRUE,
                  'return' => fn($arg) => $arg ? "デフォルト: 表示" : "デフォルト: 非表示",
                ],
        ],
      ],
      'list_string' => [
        'allowed_values' => [
          [
            'compare' => fn($arg) => TRUE,
            'return' => fn($arg) => implode("\n", array_map(fn ($key, $value) => "$key|$value", array_keys($arg), $arg)),
          ],
        ],
      ],
      'list_integer' => [
        'allowed_values' => [
          [
            'compare' => fn($arg) => TRUE,
            'return' => fn($arg) => implode("\n", array_map(fn ($key, $value) => "$key|$value", array_keys($arg), $arg)),
          ],
        ],
      ],
      'integer' => [
        'min' => [
          [
            'compare' => fn($arg) => is_numeric($arg),
            'return' => fn($arg) => "最小値: $arg",
          ],
        ],
        'max' => [
          [
            'compare' => fn($arg) => is_numeric($arg),
            'return' => fn($arg) => "最大値: $arg",
          ],
        ],
        'prefix' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "接頭辞: $arg",
          ],
        ],
        'suffix' => [
          [
            'compare' => fn($arg) => !empty($arg),
            'return' => fn($arg) => "接尾辞: $arg",
          ],
        ],
      ],
      'decimal' => [
        'min' => [
          [
            'compare' => fn($arg) => is_numeric($arg),
            'return' => fn($arg) => "最小値: $arg",
          ],
        ],
        'max' => [
            [
              'compare' => fn($arg) => is_numeric($arg),
              'return' => fn($arg) => "最大値: $arg",
            ],
        ],
        'prefix' => [
              [
                'compare' => fn($arg) => !empty($arg),
                'return' => fn($arg) => "接頭辞: $arg",
              ],
        ],
        'suffix' => [
                [
                  'compare' => fn($arg) => !empty($arg),
                  'return' => fn($arg) => "接尾辞: $arg",
                ],
        ],
        'precision' => [
                  [
                    'compare' => fn($arg) => $arg > 10,
                    'return' => fn($arg) => "有効桁数: $arg",
                  ],
        ],
        'scale' => [
                    [
                      'compare' => fn($arg) => $arg !== 2,
                      'return' => fn($arg) => "小数点以下桁数: $arg",
                    ],
        ],
      ],
      'boolean' => [
        'on_label' => [
          [
            'compare' => fn($arg) => TRUE,
            'return' => fn($arg) => "ONの名称: $arg",
          ],
        ],
        'off_label' => [
          [
            'compare' => fn($arg) => TRUE,
            'return' => fn($arg) => "OFFの名称: $arg",
          ],
        ],
      ],
      'text_with_summary' => [
        'display_summary' => [
          [
            'compare' => (fn($arg) => $arg === FALSE),
            'return' => '概要の入力欄を表示しない',
          ],
        ],
        'required_summary' => [
          [
            'compare' => (fn($arg) => $arg === TRUE),
            'return' => '概要を必須にする',
          ],
        ],
        'allowed_formats' => [
          [
            'compare' => (fn($arg) => is_array($arg) && !empty($arg)),
            'return' => (fn($arg) => "許可するフォーマット: " . implode(", ", $arg)),
          ],
        ],
      ],
      'video_embed_field' => [
        'allowed_formats' => [
              [
                'compare' => (fn($arg) => is_array($arg) && !empty($arg)),
                'return' => (fn($arg) => "許可する提供者: " . implode(", ", $arg)),
              ],
        ],
      ],
      'text_long' => [
        'allowed_formats' => [
          [
            'compare' => (fn($arg) => is_array($arg) && !empty($arg)),
            'return' => (fn($arg) => "許可するフォーマット: " . implode(", ", $arg)),
          ],
        ],
      ],
      'text' => [
        'max_length' => [
          [
            'compare' => (fn($arg) => is_numeric($arg)),
            'return' => (fn($arg) => "文字数上限: " . $arg),
          ],
        ],
        'allowed_formats' => [
          [
            'compare' => (fn($arg) => is_array($arg) && !empty($arg)),
            'return' => (fn($arg) => "許可するフォーマット: " . implode(", ", $arg)),
          ],
        ],
      ],
      'string' => [
        'max_length' => [
                [
                  'compare' => (fn($arg) => is_numeric($arg)),
                  'return' => (fn($arg) => "文字数上限: " . $arg),
                ],
        ],
      ],
      'string_long' => [
        'case_sensitive' => [
                [
                  'compare' => (fn($arg) => $arg === TRUE),
                  'return' => (fn($arg) => "大文字と小文字を区別"),
                ],
        ],
      ],
    ];

    $sheet = $this->sheet;
    $sheet->setTitle('項目 | Fields');
    $row = $this->setHeaders([
      [5, 12, 15, 20, 27, 31, 33.5, 8, 7, 21, 10, 46, 30],
      [
        "番号", "分類", "コンテンツタイプ", "項目名（日本語）", "項目名（英語）",
        "システム内部名称\n（26文字以内）", "フィールドタイプ", "必須", "個数", "初期値", "翻訳可", "とりうる値/制限", "備考",
      ],
      [
        "No.", "Type", "Content Type", "Field Name (Japanese)", "Field Name (English)",
        "Machine Name\n(under 27 characters)", "Field Type", "Required", "Number", "Default Value", "Multilingual", "Field Settings", "Remarks",
      ],
    ]);

    // Content types.
    $row = $this->setEntities($sheet, $row, 'node_type', 'node');

    // Taxonomy terms.
    $row = $this->setEntities($sheet, $row, 'taxonomy_vocabulary', 'taxonomy_term');

    // Media.
    $row = $this->setEntities($sheet, $row, 'media_type', 'media');

    // Paragraph.
    $row = $this->setEntities($sheet, $row, 'paragraphs_type', 'paragraph');

    // User.
    $this->setEntities($sheet, $row, 'user_role', 'user');

    $sheet->getRowDimension(1)->setRowHeight(48);

    $this->setStyle();

    $this->setStyleCenter('G:K');

    $this->setBorders();
  }

  /**
   * Sets rows for entities.
   */
  protected function setRows(Worksheet $sheet, array $fields, int $row, string $columnCValue, string $entityTypeId, string $bundle): int {
    foreach ($fields as $field_name => $field_definition) {
      if (is_array($this->settings['blackListFields']) && in_array($field_name, $this->settings['blackListFields'])) {
        $this->logger->notice(dt("Omitting blacklisted field '!field_name'", ['!field_name' => $field_name]));
        continue;
      }

      // Necessary to get default value.
      $fieldConfig = $this->entityTypeManager->getStorage('field_config')->load($entityTypeId . '.' . $bundle . '.' . $field_name);
      $entityType = $this->translate($this->entityTypeManager->getDefinition($entityTypeId)?->getLabel() ?? '');
      $fieldTypeLabel = $this->getFieldTypeLabel($field_definition);
      $isRequired = $this->buildCheck($field_definition->isRequired());
      $translatable = $this->buildCheck($field_definition->isTranslatable());
      $cardinality = $field_definition->getFieldStorageDefinition()->getCardinality();
      $cardinality = $cardinality == -1 ? 'N' : $cardinality;
      $englishLabel = $field_definition->getLabel();
      $japaneseLabel = $this->translate($fieldConfig, '') ?: $this->translate($englishLabel, '');
      $defaultValue = $this->getDefaultFieldValue($fieldConfig, $field_definition);
      $formattedSettings = $this->formatFieldSettings($field_definition);

      $this->setCell($sheet, 'A', $row, '=ROW()-1');
      $this->setCell($sheet, 'B', $row, $entityType, TRUE);
      $this->setCell($sheet, 'C', $row, $this->translate($columnCValue), TRUE);
      $this->setCell($sheet, 'D', $row, $japaneseLabel);
      $this->setCell($sheet, 'E', $row, $englishLabel);
      $this->setCell($sheet, 'F', $row, $field_name);
      $this->setCell($sheet, 'G', $row, $fieldTypeLabel);
      $this->setCell($sheet, 'H', $row, $isRequired);
      $this->setCell($sheet, 'I', $row, $cardinality);
      $this->setCell($sheet, 'J', $row, $defaultValue);
      $this->setCell($sheet, 'K', $row, $translatable);
      $this->setCell($sheet, 'L', $row, $formattedSettings);

      $row++;
    }
    return $row;
  }

  /**
   * Sets entities in the Fields sheet.
   */
  private function setEntities(Worksheet $sheet, int $row, string $entityCategory, string $entityTypeId): int {
    $entities = $this->entityTypeManager->getStorage($entityCategory)->loadMultiple();
    $universalFields = $this->getUniversalFields($entities, $entityTypeId);

    if (!(isset($this->settings['hideReadOnly']) &&$this->settings['hideReadOnly'] === TRUE)) {
      $row = $this->setUniversalFields($sheet, $universalFields, 'コンテンツ共通項目 (Read Only)', $row, $entities, $entityTypeId);
    }

    $row = $this->setUniversalFields($sheet, $universalFields, 'コンテンツ共通項目 (Editable)', $row, $entities, $entityTypeId);
    foreach ($entities as $entity) {
      $fields = array_diff_key($this->entityFieldManager->getFieldDefinitions($entityTypeId, (string) $entity->id()), $universalFields['All']);
      $row = $this->setRows($sheet, $fields, $row, (string) $entity->label(), $entityTypeId, (string) $entity->id());
    }

    return $row;
  }

  /*
   * UTILITIES
   */

  /**
   * Gets arrays of universal field names.
   */
  private function getUniversalFields(array $entities, string $entityTypeId): array {
    $editablePatterns = ['/^body$/', '/^title$/', '/^field_.*/'];
    $editableFieldsList = [];
    $readOnlyFieldsList = [];

    foreach ($entities as $entity) {
      $editableFields = [];
      $readOnlyFields = [];
      $fields = $this->entityFieldManager->getFieldDefinitions($entityTypeId, $entity->id());

      foreach ($fields as $fieldName => $fieldDefinition) {
        $isEditable = FALSE;
        foreach ($editablePatterns as $pattern) {
          if (preg_match($pattern, $fieldName)) {
            $isEditable = TRUE;
            break;
          }
        }

        if ($isEditable) {
          $editableFields[] = $fieldName;
        }
        else {
          $readOnlyFields[] = $fieldName;
        }
      }

      $editableFieldsList[] = $editableFields;
      $readOnlyFieldsList[] = $readOnlyFields;
    }

    // Intersect keys to find common fields across all entities.
    // And return them as array values instead of keys.
    $commonEditable = $this->getFieldDefinitionsArray(
          $entities,
          array_intersect(...$editableFieldsList),
          $entityTypeId
      );
    $commonReadOnly = $this->getFieldDefinitionsArray(
          $entities,
          array_intersect(...$readOnlyFieldsList),
          $entityTypeId
      );

    return [
      'コンテンツ共通項目 (Read Only)' => $commonReadOnly,
      'コンテンツ共通項目 (Editable)' => $commonEditable,
      'All' => array_merge($commonEditable, $commonReadOnly),
    ];
  }

  /**
   * Converts array of strings into array of definitions.
   */
  private function getFieldDefinitionsArray(array $entities, array $universalFieldNames, string $entityTypeId): array {
    $definitions = [];
    foreach ($entities as $entity) {
      $fields = $this->entityFieldManager->getFieldDefinitions($entityTypeId, $entity->id());
      foreach ($fields as $fieldName => $fieldDefinition) {
        if (in_array($fieldName, $universalFieldNames)) {
          $definitions[$fieldName] = $fieldDefinition;
        }
      }
    }
    return $definitions;
  }

  /**
   * Sets universal fields for entities on the Fields sheet.
   */
  private function setUniversalFields(Worksheet $sheet, array $universalFields, string $key, int $row, array $content_types, string $entityTypeId): int {
    return $this->setRows($sheet, $universalFields[$key], $row, $key, $entityTypeId, reset($content_types)->id());
  }

  /**
   * Formats the field settings for display based on the field type.
   */
  private function formatFieldSettings(FieldDefinitionInterface $fieldDefinition): string {
    $fieldType = $fieldDefinition->getType();
    $settings = $fieldDefinition->getSettings();
    $formattedSettings = [];
    if (isset($this->fieldformatter[$fieldType])) {
      $locale = $this->fieldformatter[$fieldType];
      foreach ($settings as $key => $value) {
        if (!isset($locale[$key])) {
          continue;
        }
        foreach ($locale[$key] as $localeItem) {
          // Ensure both 'compare' and 'return' are set.
          if (!isset($localeItem['compare'], $localeItem['return'])) {
            continue;
          }

          $compare = $localeItem['compare'];
          $return = $localeItem['return'];
          if ($compare($value, $settings)) {
            $formattedSettings[] = is_callable($return) ? $return($value) : $value;
          }
        }
      }
      return implode("\n", $formattedSettings);
    }

    // depreciated. Convert to above
    // Custom formatting based on field type.
    switch ($fieldType) {
      case 'entity_reference':
        $label = '参照コンテンツタイプ: ';
        $handlerSettings = $settings['handler_settings'];
        $bundles = $handlerSettings['target_bundles'] ?? NULL;
        if (empty($bundles)) {
          return '';
        }
        $formattedSettings[] = $label . implode(", ", $bundles);
        if ($handlerSettings['auto_create']) {
          $formattedSettings[] = '参照先のエンティティが存在しなければ作成する';
        }
        if ($handlerSettings['sort']['field'] !== '_none') {
          $formattedSettings[] = "並び替え基準: " . $handlerSettings['sort']['field'];
          $formattedSettings[] = "並べ替えの向き: " . ($handlerSettings['sort']['direction'] === 'ASC' ? '昇順' : '降順');
        }
        return implode("\n", $formattedSettings);

      case 'entity_reference_revisions':
        $label = $settings['handler_settings']['negate'] ? '参照しないパラグラフタイプ: ' : '参照パラグラフタイプ: ';
        $bundles = $settings['handler_settings']['target_bundles_drag_drop'];
        $enabledBundles = [];
        foreach ($bundles as $bundle => $details) {
          if (isset($details['enabled']) && $details['enabled']) {
            $enabledBundles[] = $bundle;
          }
        }

        if (empty($enabledBundles)) {
          return '';
        }

        return $label . implode(", ", $enabledBundles);
    }

    if (!empty($settings)) {
      $formattedSettings[] = "($fieldType)";
    }
    foreach ($settings as $key => $value) {
      // Format the value as a string for readability.
      if (is_bool($value)) {
        $value = $value ? 'true' : 'false';
      }
      elseif (is_array($value)) {
        $value = json_encode($value);
      }
      else {
        $value = "$value";
      }
      $formattedSettings[] = "$key: $value";
    }
    return implode("\n", $formattedSettings);
  }

  /**
   * Gets the default value for a field config.
   */
  private function getDefaultFieldValue(EntityInterface|NULL $fieldConfig, FieldDefinitionInterface $fieldDefinition):string {
    if (!($fieldConfig instanceof FieldConfigInterface)) {
      return '';
    }

    $defaultValue = $fieldConfig->get('default_value');
    $fieldType = $fieldDefinition->getType();

    if (
      empty($defaultValue) ||
      !is_array($defaultValue) ||
      !is_array($defaultValue[0])
    ) {
      return '';
    }
    elseif ($fieldType === 'timestamp') {
      $date = DrupalDateTime::createFromTimestamp($defaultValue[0]['value']);
      return $date->format('Y-m-d\TH:i:sP');
    }
    elseif ($fieldType === 'boolean') {
      return $defaultValue[0]['value'] ? 'TRUE' : 'FALSE';
    }
    elseif (count($defaultValue) === 1) {
      return (string) $defaultValue[0]['value'] ?: '';
    }
    else {
      return json_encode($defaultValue) ?: '';
    }
  }

}
