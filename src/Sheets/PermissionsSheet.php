<?php

namespace Drupal\risley_export\Sheets;

use Drupal\Component\Plugin\Exception\PluginNotFoundException;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
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
      [5, 20, 20, 30, 25, 60],
      ["番号", "", "", "権限 (日本語)", "モジュール", "権限 (英語)"],
      ["No.", "", "", "Permission (Japanese)", "Module", "Permission (English)"],
    ];

    // Values are.
    $roles = array_values(array_map(function ($role) {
      return $role->get('label');
    }, $this->getPermissionRoles()));
    $roleCount = count($roles);
    $startCount = count($headers[0]);
    $mergeRanges = [];
    for ($i = 0; $i < $roleCount; $i++) {
      $headers[0][] = 3;
    }
    foreach ($roles as $i => $roleLabel) {
      if (count($headers) <= $i + 1) {
        $headers[] = [];
        for ($j = 0; $j < $startCount; $j++) {
          $headers[$i + 1][] = [
            'value' => '',
            'fill' => [
              'fillType' => Fill::FILL_SOLID,
              'startColor' => [
                'argb' => 'FFBFBFBF',
              ],
            ],
            'borders' => [
              'left' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
              ],
              'right' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
              ],
              'top' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FFBFBFBF'],
              ],
              'bottom' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FFBFBFBF'],
              ],
            ],
          ];
        }
      }

      // Left of the name blank.
      for ($j = 0; $j < $i; $j++) {
        $headers[$i + 1][] = [
          'value' => '',
          'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => [
              'argb' => 'FFBFBFBF',
            ],
          ],
          'borders' => [
            'left' => [
              'borderStyle' => Border::BORDER_THIN,
              'color' => ['argb' => 'FF000000'],
            ],
            'right' => [
              'borderStyle' => Border::BORDER_THIN,
              'color' => ['argb' => 'FF000000'],
            ],
            'top' => [
              'borderStyle' => Border::BORDER_THIN,
              'color' => ['argb' => 'FFBFBFBF'],
            ],
            'bottom' => [
              'borderStyle' => Border::BORDER_THIN,
              'color' => ['argb' => 'FFBFBFBF'],
            ],
          ],
        ];
      }
      // Role name.
      $headers[$i + 1][] = [
        'value' => $roleLabel,
        'font' => [
          'name' => 'Meiryo UI',
          'size' => 10,
        ],
        'alignment' => [
        // Ensure this is false to prevent wrapping.
          'wrapText' => FALSE,
        // Example alignment.
          'horizontal' => Alignment::HORIZONTAL_LEFT,
        ],
        'fill' => [
          'fillType' => Fill::FILL_SOLID,
          'startColor' => [
            'argb' => 'FFBFBFBF',
          ],
        ],
        'borders' => [
          'left' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => 'FF000000'],
          ],
          'top' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => 'FF000000'],
          ],
          'right' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => 'FFBFBFBF'],
          ],
          'bottom' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => 'FFBFBFBF'],
          ],
        ],
      ];
      $mergeRanges[] =
        $this->intToCol(count($headers[$i + 1]) - 1) .
        ($i + 1) .
        ':' .
        $this->intToCol($roleCount + $startCount) .
        ($i + 1);
      for ($j = $i + 1; $j < $roleCount; $j++) {
        $headers[$i + 1][] = [
          'value' => '',
          'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => [
              'argb' => 'FFBFBFBF',
            ],
          ],
          'borders' => [
            'top' => [
              'borderStyle' => Border::BORDER_THIN,
              'color' => ['argb' => 'FF000000'],
            ],
            'bottom' => [
              'borderStyle' => Border::BORDER_THIN,
              'color' => ['argb' => 'FF000000'],
            ],
            'right' => [
              'borderStyle' => Border::BORDER_THIN,
              'color' => ['argb' => 'FFBFBFBF'],
            ],
            'left' => [
              'borderStyle' => Border::BORDER_THIN,
              'color' => ['argb' => 'FFBFBFBF'],
            ],
          ],
        ];
      }
    }

    $row = $this->setHeaders($headers);
    $this->merge($mergeRanges);
    if ($roleCount > 2) {
      for ($i = 0; $i < 6; $i++) {
        $col = $this->intToCol($i);
        $lastRow = $row - 1;
        $range = "{$col}2:$col$lastRow";
        $this->merge($range);
      }
    }

    $this->setRows($sheet, $row);

    // $this->setStyle();
    // $this->setBorders();
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
      $authenticatedHasPermission = in_array($key, $roles['authenticated']->get('permissions'));

      if (
        isset($this->settings['hideEmptyPermissions']) &&
        $this->settings['hideEmptyPermissions'] === TRUE &&
        !(
          is_array($this->settings['whiteListPermissions']) &&
          in_array($key, $this->settings['whiteListPermissions'])
        )
      ) {
        $some = FALSE;
        foreach ($roles as $roleName => $role) {
          if ($roleName === 'administrator' || in_array($key, $role->get('permissions')) || ($role->id() !== 'anonymous' && $authenticatedHasPermission)) {
            $some = TRUE;
            break;
          }
        }
        if (!$some) {
          continue;
        }
        $this->logger->notice(dt("Omitting empty permission '!key'", ['!key' => $key]));
      }

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

      foreach ($roles as $roleName => $role) {
        $isEnabled = $roleName === 'administrator' || in_array($key, $role->get('permissions')) || ($role->id() !== 'anonymous' && $authenticatedHasPermission);
        $this->setCell($sheet, $column, $row, $this->buildCheck($isEnabled));

        $centerAlignmentStyle = [
          'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER,
          ],
        ];
        $sheet->getStyle("$column$row:$column$row")->applyFromArray($centerAlignmentStyle);

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
      $node_type_entity = $this->entityTypeManager->getStorage("{$content_type}_$library")->load($node_type);
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
