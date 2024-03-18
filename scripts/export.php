<?php

/**
 * @file
 * A script to programmatically export drupalsettings from a script.
 */

use Drupal\risley_export\Commands\RisleyExportCommands;
use Drupal\risley_export\Sheets\Factory\BaseSheetFactory;
use Drush\Drush;

require_once 'autoload.php';
require_once DRUPAL_ROOT . '/modules/custom/risley_export/src/Commands/RisleyExportCommands.php';
require_once DRUPAL_ROOT . '/modules/custom/risley_export/src/Sheets/Factory/BaseSheetFactory.php';

// Create instances of the required services.
$baseSheetFactory = new BaseSheetFactory(
  \Drupal::service('entity_field.manager'),
  \Drupal::service('entity_type.manager'),
  \Drupal::service('plugin.manager.field.field_type'),
  \Drupal::service('language_manager'),
  \Drupal::service('config.factory'),
  \Drupal::service('entity_type.bundle.info'),
  \Drupal::service('entity.repository'),
  \Drupal::service('extension.list.module'),
  \Drupal::service('user.permissions'),
  \Drupal::service('logger.factory'),
  \Drupal::service('info_parser'),
  \Drupal::service('module_handler'),
  Drush::service('site.alias.manager')
);

try {
  $command = new RisleyExportCommands($baseSheetFactory);
  $command->exportSheetData(['path' => 'modules/custom/risley_export/files', 'file' => $_SERVER['argv'][4] ?? NULL, 'filename' => 'file.xlsx']);
}
catch (\Exception $e) {
  echo 'Exception caught: ', $e->getMessage(), "\n";
}
