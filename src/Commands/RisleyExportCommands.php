<?php

namespace Drupal\risley_export\Commands;

use Drupal\risley_export\Sheets\Factory\BaseSheetFactory;
use Drush\Commands\DrushCommands;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * A Drush commandfile.
 */
class RisleyExportCommands extends DrushCommands {

  /**
   * The spreadsheet being built.
   *
   * @var \Drupal\risley_export\Sheets\Factory\BaseSheetFactory
   */
  protected $sheetFactory;

  /**
   * The current filename.
   *
   * @var string
   */
  protected $filename;

  /**
   * The current path.
   *
   * @var string
   */
  protected $path;

  /**
   * The options passed through the drush command.
   *
   * @var array{
   *    path: string,
   *    file: string|null,
   *    filename: string
   *  }
   */
  protected $options;

  /**
   * The current spreadsheet being built.
   *
   * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
   */
  protected $spreadsheet;

  /**
   * Constructs a new RisleyExportCommands object.
   *
   * @param \Drupal\risley_export\Sheets\Factory\BaseSheetFactory $sheet_factory
   *   The base sheet factory.
   */
  public function __construct(
    BaseSheetFactory $sheet_factory
  ) {
    $this->sheetFactory = $sheet_factory;
  }

  /**
   * Exports content types and their fields to an Excel file.
   *
   * @param array $options
   *   An associative array of options.
   *
   * @option path
   *   The directory path where the Excel file will be saved.
   *   Defaults to '../dev'.
   *
   * @command risley_export:export
   * @aliases re:e
   * @usage risley_export:export --path=/path/to/directory/from/docroot
   *   Exports content types and fields to the specified directory.
   */
  public function exportSheetData(array $options = ['path' => 'modules/custom/risley_export/files', 'file' => NULL, 'filename' => 'file.xlsx']): void {
    if (!isset($options['path']) || !is_string($options['path']) ||
      (!is_string($options['file']) && $options['file'] !== NULL) ||
      !isset($options['filename']) || !is_string($options['filename'])) {
      throw new \InvalidArgumentException('Invalid options structure.');
    }
    $this->options = $options;
    $this->path = $options['path'];

    if (!$this->options['file'] || $this->options['file'] === 'data') {
      $this->buildSpreadsheet('drupalsettings_data', ['Version', 'Fields', 'Content', 'Taxonomies', 'Media', 'Paragraphs']);
    }
    if (!$this->options['file'] || $this->options['file'] === 'modules') {
      $this->buildSpreadsheet('drupalsettings_modules', ['Version', 'CoreModules', 'ContribModules', 'CustomModules']);
    }
    if (!$this->options['file'] || $this->options['file'] === 'permissions') {
      $this->buildSpreadsheet('drupalsettings_permissions', ['Version', 'Roles', 'Permissions', 'Workflows']);
    }
    if (!$this->options['file'] || $this->options['file'] === 'content') {
      $this->buildSpreadsheet('drupalsettings_content', ['Version', 'Menus', 'TaxonomiesForContent', 'ContentForContent', 'Redirects']);
    }
    if (!$this->options['file'] || $this->options['file'] === 'webforms') {
      $this->buildSpreadsheet('drupalsettings_webforms', ['Version', 'WebformsContent', 'WebformsOptions', 'WebformsPermissions']);
    }

  }

  /**
   * Builds the filename.
   */
  private function buildSpreadsheet(string $filename = 'excel_sheet', array $sheetNames = ['Version']): void {
    try {
      if (!file_exists($this->path) || !is_writable($this->path)) {
        if ($this->path === 'modules/custom/risley_export/files') {
          if (!mkdir($this->path, 0775, TRUE) && !is_dir($this->path)) {
            throw new \Exception(dt('Failed to create the directory.'));
          }
        }
        else {
          throw new \Exception(dt('Invalid directory path or the directory is not writable.'));
        }
      }

      $absolutePath = realpath($this->path);
      if ($absolutePath === FALSE) {
        throw new \Exception(dt('Unable to resolve the absolute path of the specified directory.'));
      }

      $this->options['filename'] = $absolutePath . DIRECTORY_SEPARATOR . "_$filename.xlsx";
    }
    catch (\Exception $e) {
      $this->logger()?->error($e->getMessage());
      exit;
    }
    $this->spreadsheet = new Spreadsheet();

    foreach ($sheetNames as $sheetName) {
      $this->buildSheet($sheetName, $this->options);
    }

    $this->saveSheet();
  }

  /**
   * Builds the given sheet.
   */
  private function buildSheet(string $key, array $options): Worksheet|NULL {
    $sheet = $this->sheetFactory->create($this->spreadsheet, $key, $options);
    if (!isset($sheet)) {
      return NULL;
    }

    if ($this->spreadsheet->getSheetCount() > 1) {
      $worksheet = $this->spreadsheet->getSheetByName('Worksheet');
      if ($worksheet !== NULL) {
        $this->spreadsheet->removeSheetByIndex($this->spreadsheet->getIndex($worksheet));
      }
    }

    return $sheet;
  }

  /**
   * Writes the spreadsheet to file.
   */
  private function saveSheet(): void {
    $fileType = ucfirst($this->settings['filetype'] ?? 'xlsx'); // Default to 'Xlsx' if not set
    $className = "PhpOffice\\PhpSpreadsheet\\Writer\\$fileType";

    if (!class_exists($className)) {
        throw new \Exception("Class $className does not exist.");
    }

    $writer = new $className($this->spreadsheet);
    try {
      $writer->save($this->options['filename']);
      if (!file_exists($this->options['filename'])) {
        $this->logger()?->error(dt('Failed to create the file in the specified directory. Check permissions and path.'));
      }
      elseif (file_exists($this->options['filename'])) {
        $this->logger()?->success(dt('Content types and fields have been exported to !filename', ['!filename' => $this->options['filename']]));
      }
    }
    catch (\Exception $e) {
      $this->logger()?->error(dt('An error occurred while saving the file: !message', ['!message' => $e->getMessage()]));
    }
  }

}
