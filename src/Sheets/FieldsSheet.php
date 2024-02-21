<?php

namespace Drupal\risley_export\Sheets;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Fields sheet for data spreadsheet.
 */
class FieldsSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize(): void {
    $sheet = $this->sheet;
    $sheet->setTitle('項目 | Fields');
    $headers = [
      "No.", "分類\nType", "コンテンツタイプ\nContent Type", "項目名\n（日本語）", "Field name\n(English)",
      "システム内部名称\nMachine Name\n(26文字以内)", "フィールドタイプ\nField Type", "必須\nRequired", "個数\nNumber", "初期値\nDefault Value", "翻訳可\nMultilingual", "とりうる値/制限\nField Settings", "備考\nRemarks",
    ];
    $sheet->fromArray($headers, NULL, 'A1');
    $row = 2;

    // Content types.
    $row = $this->setEntities($sheet, $row, 'node_type', 'node');

    // Taxonomy terms.
    $row = $this->setEntities($sheet, $row, 'taxonomy_vocabulary', 'taxonomy_term');

    // Media.
    $row = $this->setEntities($sheet, $row, 'media_type', 'media');

    // Paragraph.
    $row = $this->setEntities($sheet, $row, 'paragraphs_type', 'paragraph');

    // User.
    $row = $this->setEntities($sheet, $row, 'user_role', 'user');

    // Entity.
    // Custom table.
    // Format sheet.
    $sheet->getRowDimension(1)->setRowHeight(48);

    $sheet->getColumnDimension('A')->setWidth(5);
    $sheet->getColumnDimension('B')->setWidth(12);
    $sheet->getColumnDimension('C')->setWidth(14);
    $sheet->getColumnDimension('D')->setWidth(20);
    $sheet->getColumnDimension('E')->setWidth(27);
    $sheet->getColumnDimension('F')->setWidth(31);
    $sheet->getColumnDimension('G')->setWidth(33.5);
    $sheet->getColumnDimension('H')->setWidth(5.5);
    $sheet->getColumnDimension('I')->setWidth(4.5);
    $sheet->getColumnDimension('J')->setWidth(21);
    $sheet->getColumnDimension('K')->setWidth(8.5);
    $sheet->getColumnDimension('L')->setWidth(46);

    $this->setStyle($sheet);

    $centerAlignmentStyle = [
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
    ];
    $sheet->getStyle('G:G')->applyFromArray($centerAlignmentStyle);
    $sheet->getStyle('H:H')->applyFromArray($centerAlignmentStyle);
    $sheet->getStyle('I:I')->applyFromArray($centerAlignmentStyle);
    $sheet->getStyle('K:K')->applyFromArray($centerAlignmentStyle);
    $sheet->getStyle('J:J')->applyFromArray($centerAlignmentStyle);

    $this->setBorders();
  }

  /**
   * Sets rows for entities.
   */
  protected function setRows(Worksheet $sheet, array $fields, int $row, string $columnCValue, string $entityTypeId, string $bundle): int {
    foreach ($fields as $field_name => $field_definition) {

      // Necessary to get default value.
      $fieldConfig = $this->entityTypeManager->getStorage('field_config')->load($entityTypeId . '.' . $bundle . '.' . $field_name);
      $entityType = $this->translate($this->entityTypeManager->getDefinition($entityTypeId)?->getLabel() ?? '');
      $fieldType = $field_definition->getType();
      $fieldTypeLabel = is_array($_ = $this->fieldTypePluginManager->getDefinition($fieldType)) ? $_['label'] : '';
      $isRequired = $this->buildCheck($field_definition->isRequired());
      $translatable = $this->buildCheck($field_definition->isTranslatable());
      $cardinality = $field_definition->getFieldStorageDefinition()->getCardinality();
      $cardinality = $cardinality == -1 ? 'N' : $cardinality;
      $englishLabel = $field_definition->getLabel();
      $japaneseLabel = $this->translate($fieldConfig, '');
      $defaultValue = $fieldConfig ? $fieldConfig->get('default_value') : NULL;
      $formattedDefaultValue = $this->formatDefaultValue($defaultValue);
      $formattedSettings = $this->formatFieldSettings($fieldType, $field_definition->getSettings());

      $this->setCell($sheet, 'A', $row, '=ROW()-1');
      $this->setCell($sheet, 'B', $row, $entityType, TRUE);
      $this->setCell($sheet, 'C', $row, $this->translate($columnCValue), TRUE);
      $this->setCell($sheet, 'D', $row, $japaneseLabel === $englishLabel ? '' : $japaneseLabel);
      $this->setCell($sheet, 'E', $row, $englishLabel);
      $this->setCell($sheet, 'F', $row, $field_name);
      $this->setCell($sheet, 'G', $row, $fieldTypeLabel);
      $this->setCell($sheet, 'H', $row, $isRequired);
      $this->setCell($sheet, 'I', $row, $cardinality);
      $this->setCell($sheet, 'J', $row, $formattedDefaultValue);
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

    if (!$this->options['no-readonly']) {
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
   * Formats the default value for display in the spreadsheet.
   */
  private function formatDefaultValue(mixed $defaultValue): mixed {
    if (empty($defaultValue)) {
      return '';
    }

    // Check if the default value is an array with a single key 'value'.
    if (is_array($defaultValue) && count($defaultValue) === 1 && isset($defaultValue[0]['value'])) {
      return (string) $defaultValue[0]['value'];
    }

    // If default value is an array, format it as a string.
    if (is_array($defaultValue)) {
      return json_encode($defaultValue) ?: '';
    }

    return $defaultValue;
  }

  /**
   * Formats the field settings for display based on the field type.
   */
  private function formatFieldSettings(string $fieldType, array $settings): string {
    // Custom formatting based on field type.
    switch ($fieldType) {
      case 'entity_reference_revisions':
        $bundles = $settings['handler_settings']['target_bundles_drag_drop'];
        $enabledBundles = [];
        foreach ($bundles as $bundle => $details) {
          if (isset($details['enabled']) && $details['enabled']) {
            $enabledBundles[] = $bundle;
          }
        }

        if (empty($enabledBundles)) {
          return 'No enabled types';
        }

        return implode(", ", $enabledBundles);
    }

    $formattedSettings = [];
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

}
