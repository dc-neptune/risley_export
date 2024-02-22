<?php

/**
 * @file
 * A script to programmatically export drupalsettings from a script.
 */

use DrupalFinder\DrupalFinder;
use Drush\Drush;

// Find and load Drupal.
require_once __DIR__ . '/vendor/autoload.php';

$finder = new DrupalFinder();
$finder->locateRoot(__DIR__);

if (!$finder->locateRoot()) {
    throw new \RuntimeException('Drupal root not found. Please check your setup.');
}

// Set up the site path and bootstrap Drupal.
$drupal_root = $finder->getDrupalRoot();
$site_path = $finder->getSitePath();

// Bootstrap Drupal.
define('DRUPAL_ROOT', $drupal_root);
require_once DRUPAL_ROOT . '/core/includes/bootstrap.inc';
drupal_bootstrap(DRUPAL_BOOTSTRAP_FULL);

// Load necessary Drupal components.
\Drupal::service('kernel')->boot();

// Initialize Drush.
Drush::bootstrap();

// Load your Drush command class.
require_once __DIR__ . '/modules/custom/risley_export/Commands/RisleyExportCommands.php';

// Create an instance of your command class.
$command = new \Drupal\risley_export\Commands\RisleyExportCommands(
    \Drupal::service('risley_export.sheet_factory')
);

// Execute the method from your command class.
$command->exportSheetData();
