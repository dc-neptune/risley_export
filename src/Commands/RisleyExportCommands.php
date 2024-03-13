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
   * @var array<mixed>
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
   * @option no-readonly
   *   If set, readonly fields will not be included. Defaults to false.
   *
   * @command risley_export:export
   * @aliases re:e
   * @usage risley_export:export --path=/path/to/directory/from/docroot
   *   Exports content types and fields to the specified directory.
   * @usage risley_export:export --no-readonly
   *   Exports content types and fields excluding readonly fields.
   */
  public function exportSheetData(array $options = ['path' => 'modules/custom/risley_export/files', 'no-readonly' => FALSE]): void {
    $this->options = $options;
    $this->path = $options['path'];

    $this->buildSpreadsheet('drupalsettings_data', ['Version', 'Fields', 'Content', 'Taxonomies', 'Media', 'Paragraphs']);
    $this->buildSpreadsheet('drupalsettings_modules', ['Version', 'CoreModules', 'ContribModules', 'CustomModules']);
    $this->buildSpreadsheet('drupalsettings_permission', ['Version', 'Roles', 'Permissions', 'Workflows']);
    $this->buildSpreadsheet('drupalsettings_content', ['Version', 'Menus', 'TaxonomiesForContent', 'ContentForContent', 'Redirects']);

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

      $this->filename = $absolutePath . DIRECTORY_SEPARATOR . "_$filename.xlsx";
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
    // Write file to the specified path.
    $writer = new Xlsx($this->spreadsheet);
    try {
      $writer->save($this->filename);
      if (!file_exists($this->filename)) {
        $this->logger()?->error(dt('Failed to create the file in the specified directory. Check permissions and path.'));
      }
      elseif (file_exists($this->filename)) {
        $this->logger()?->success(dt('Content types and fields have been exported to !filename', ['!filename' => $this->filename]));
      }
    }
    catch (\Exception $e) {
      $this->logger()?->error(dt('An error occurred while saving the file: !message', ['!message' => $e->getMessage()]));
    }
  }

}
