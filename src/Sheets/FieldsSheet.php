<?php

namespace Drupal\risley_export\Sheets;

use Drupal\Core\Datetime\DrupalDateTime;
use Drupal\Core\Field\FieldConfigInterface;
use Drupal\Core\Field\FieldDefinitionInterface;
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
    $row = $this->setHeaders([
      [5, 12, 15, 20, 27, 31, 33.5, 8, 7, 21, 10, 46],
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

    $this->setStyle($sheet);

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

  /**
   * Gets the default value for a field config.
   */
  private function getDefaultFieldValue(FieldConfigInterface|NULL $fieldConfig, FieldDefinitionInterface $fieldDefinition):string {
    if (!isset($fieldConfig)) {
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
    elseif (is_array($defaultValue) && count($defaultValue) === 1) {
      return (string) $defaultValue[0]['value'] ?: '';
    }
    elseif (is_array($defaultValue)) {
      return json_encode($defaultValue) ?: '';
    }
  }

}
