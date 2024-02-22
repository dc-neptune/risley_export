<?php

/**
 * @file
 * A script to programmatically export drupalsettings from a script.
 */

use Drupal\risley_export\Commands\RisleyExportCommands;

$autoloader = require_once 'autoload.php';
require_once __DIR__ . '/../src/Commands/RisleyExportCommands.php';
$command = new RisleyExportCommands(
    \Drupal::service('risley_export.base_sheet_factory')
);
$command->exportSheetData();
