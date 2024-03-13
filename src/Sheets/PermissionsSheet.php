<?php

namespace Drupal\risley_export\Sheets;

use Drupal\Component\Plugin\Exception\PluginNotFoundException;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Permissions sheet for the permissions spreadsheet.
 */
class PermissionsSheet extends BaseSheet {

  /**
   * Initializes the sheet.
   */
  protected function initialize():void {
    $sheet = $this->sheet;
    $sheet->setTitle('権限 | Permissions');
    $headers = [
      "No.", "", "", "権限 (日本語)\nPermission (Japanese)", "モジュール\nModule", "権限 (英語)\nPermission (English)",
    ];

    $sheet->getColumnDimension('A')->setWidth(5);
    $sheet->getColumnDimension('B')->setWidth(20);
    $sheet->getColumnDimension('C')->setWidth(20);
    $sheet->getColumnDimension('D')->setWidth(30);
    $sheet->getColumnDimension('E')->setWidth(25);
    $sheet->getColumnDimension('F')->setWidth(60);

    $column = 'G';
    foreach ($this->getPermissionRoles() as $role) {
      $headers[] = $role->get('label');
      $sheet->getColumnDimension($column)->setWidth(5);
      $column = $this->incrementColumn($column);
    }

    $sheet->fromArray($headers);
    $row = 2;

    $this->setRows($sheet, $row);

    $this->setStyle($sheet);

    $this->setBorders();
  }

  /**
   * Sets rows on the sheet.
   */
  protected function setRows(Worksheet $sheet, int $row): int {
    $priorityPermissions = [];
    $permissions = $this->userPermissions->getPermissions();
    $roles = $this->getPermissionRoles();

    // Save permission keys to permission for sorting.
    foreach ($permissions as $key => $permission) {
      $permissions[$key]['key'] = $key;
    }

    // Prioritize permissions listed in settings.
    if (
      isset($this->settings['priorityPermissions']) &&
      is_array($this->settings['priorityPermissions']) &&
      !empty($this->settings['priorityPermissions'])
    ) {
      $groups = $this->settings['priorityPermissions'];
      $priorityPermissions = array_filter($permissions, function ($permission) use ($groups) {
        return in_array($this->moduleExtensionList->getExtensionInfo($permission['provider'])['name'], $groups);
      });
      usort($priorityPermissions, function ($a, $b) use ($groups) {
        $providerA = $this->moduleExtensionList->getExtensionInfo($a['provider'])['name'] ?? $a['provider'];
        $providerB = $this->moduleExtensionList->getExtensionInfo($b['provider'])['name'] ?? $b['provider'];
        $configLabelA = $this->getConfigLabel($a);
        $configLabelB = $this->getConfigLabel($b);

        if ($providerA === $providerB) {
          // Sort by configLabel if providerLabels are equal.
          return $configLabelA <=> $configLabelB;
        }

        return array_search($providerA, $groups) <=> array_search($providerB, $groups);
      });
      $permissions = array_diff_key($permissions, $priorityPermissions);
    }

    // Sort permissions by providerLabel and then by configLabel.
    // foreach( [$priorityPermissions, $permissions] as $_permissions) {.
    usort($permissions, function ($a, $b) {
      $providerA = $this->moduleExtensionList->getExtensionInfo($a['provider'])['name'] ?? $a['provider'];
      $providerB = $this->moduleExtensionList->getExtensionInfo($b['provider'])['name'] ?? $b['provider'];
      $configLabelA = $this->getConfigLabel($a);
      $configLabelB = $this->getConfigLabel($b);

      if ($providerA === $providerB) {
        // Sort by configLabel if providerLabels are equal.
        return $configLabelA <=> $configLabelB;
      }
      // Otherwise, sort by providerLabel.
      return $providerA <=> $providerB;
    });
    // }
    $permissions = array_merge($priorityPermissions, $permissions);

    // Pre-process to determine which providers have a configLabel.
    $providersWithConfig = [];

    foreach ($permissions as $permission) {
      $provider = $permission['provider'] ?? '';
      $configLabel = $this->getConfigLabel($permission);

      if ($configLabel !== '') {
        $providersWithConfig[$provider] = TRUE;
      }
    }

    foreach ($permissions as $permission) {
      $key = $permission['key'];
      $title = isset($permission['title']) ? strip_tags((string) $permission['title']) : '';
      $japaneseTitle = $this->translate($title, '');
      $provider = isset($permission['provider']) ? (string) $permission['provider'] : '';
      $providerInfo = $this->moduleExtensionList->getExtensionInfo($provider);
      $providerLabel = $providerInfo['name'] ?? $provider;
      $providerLabel = is_object($providerLabel) && method_exists($providerLabel, '__toString') ? $providerLabel->__toString() : $providerLabel;
      $providerLabel = $this->translate((string) $providerLabel);
      $configLabel = $this->getConfigLabel($permission);
      if ($configLabel === '') {
        $configLabel = isset($providersWithConfig[$provider]) ? "{$providerLabel}共通" : $providerLabel;
      }

      $this->setCell($sheet, 'A', $row, '=ROW()-1');
      $this->setCell($sheet, 'B', $row, $providerLabel, TRUE);
      $this->setCell($sheet, 'C', $row, $configLabel, TRUE);
      $this->setCell($sheet, 'D', $row, $japaneseTitle);
      $this->setCell($sheet, 'E', $row, $providerLabel);
      $this->setCell($sheet, 'F', $row, $title);

      $column = 'G';

      $authenticatedHasPermission = in_array($key, $roles['authenticated']->get('permissions'));

      foreach ($roles as $roleName => $role) {
        $isEnabled = $roleName === 'administrator' || in_array($key, $role->get('permissions')) || ($role->id() !== 'anonymous' && $authenticatedHasPermission);
        $this->setCell($sheet, $column, $row, $this->buildCheck($isEnabled));

        $centerAlignmentStyle = [
          'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER,
          ],
        ];
        $sheet->getStyle("{$column}{$row}:{$column}{$row}")->applyFromArray($centerAlignmentStyle);

        $column = $this->incrementColumn($column);
      }

      $row++;
    }

    return $row;
  }

  /**
   * Gets the config label for the sheet.
   */
  private function getConfigLabel(array $permission): string {
    $dependencies = $permission['dependencies'] ?? [];
    if (!is_array($dependencies) || empty($dependencies['config'])) {
      return '';
    }

    $config = reset($dependencies['config']);
    [$content_type, $library, $node_type] = explode('.', $config);

    if ($library === 'type' || $library === 'vocabulary') {
      $node_type_entity = $this->entityTypeManager->getStorage("{$content_type}_{$library}")->load($node_type);
      if ($node_type_entity) {
        return $this->translate($node_type_entity);
      }
    }

    try {
      $node_type_entity = $this->entityTypeManager->getStorage($library)->load($node_type);
      if ($node_type_entity) {
        return $this->translate($node_type_entity);
      }
    }
    catch (PluginNotFoundException) {
      return $this->translate((string) $library);
    }

    return $this->translate((string) $library);
  }

}
