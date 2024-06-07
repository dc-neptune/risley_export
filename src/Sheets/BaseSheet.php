<?php

namespace Drupal\risley_export\Sheets;

use Consolidation\SiteAlias\SiteAliasManager;
use Drupal\Core\Config\ConfigFactoryInterface;
use Drupal\Core\Entity\EntityFieldManagerInterface;
use Drupal\Core\Entity\EntityInterface;
use Drupal\Core\Entity\EntityRepositoryInterface;
use Drupal\Core\Entity\EntityTypeBundleInfoInterface;
use Drupal\Core\Entity\EntityTypeManagerInterface;
use Drupal\Core\Extension\Extension;
use Drupal\Core\Extension\InfoParserInterface;
use Drupal\Core\Extension\ModuleExtensionList;
use Drupal\Core\Extension\ModuleHandlerInterface;
use Drupal\Core\Field\BaseFieldDefinition;
use Drupal\Core\Field\Entity\BaseFieldOverride;
use Drupal\Core\Field\FieldDefinitionInterface;
use Drupal\Core\Field\FieldTypePluginManagerInterface;
use Drupal\Core\Language\LanguageManagerInterface;
use Drupal\node\Entity\NodeType;
use Drupal\path_alias\AliasManagerInterface;
use Drupal\user\PermissionHandlerInterface;
use Drupal\webform\Entity\Webform;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Psr\Log\LoggerInterface;

/**
 * Constructs a blank sheet.
 */
class BaseSheet {

  /**
   * The spreadsheet being built.
   *
   * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
   */
  protected Spreadsheet $spreadsheet;

  /**
   * The current sheet being built.
   *
   * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
   */
  public Worksheet $sheet;

  /**
   * Tracks the last value in each column.
   *
   * @var array<mixed>
   */
  protected array $lastColumnValues;

  /**
   * Tracks whether the column has been skipped.
   *
   * @var array<string>
   */
  protected array $skippedColumns;

  /**
   * Tracks the cell where body starts and header ends.
   *
   * @var string
   */
  protected string $bodyCell;

  /**
   * The options passed through the drush command.
   *
   * @var array{
   *     path: string,
   *     file: string|null,
   *     filename: string
   *   }
   */
  protected array $options;

  /**
   * The entity field manager.
   *
   * @var \Drupal\Core\Entity\EntityFieldManagerInterface
   * @PHPStan-var EntityFieldManagerInterface
   */
  protected EntityFieldManagerInterface $entityFieldManager;

  /**
   * The entity type manager.
   *
   * @var \Drupal\Core\Entity\EntityTypeManagerInterface
   * @PHPStan-var EntityTypeManagerInterface
   */
  protected EntityTypeManagerInterface $entityTypeManager;

  /**
   * The field type plugin manager.
   *
   * @var \Drupal\Core\Field\FieldTypePluginManagerInterface
   * @PHPStan-var FieldTypePluginManagerInterface
   */
  protected FieldTypePluginManagerInterface $fieldTypePluginManager;

  /**
   * The config factory.
   *
   * @var \Drupal\Core\Config\ConfigFactoryInterface
   * @PHPStan-var ConfigFactoryInterface
   */
  protected ConfigFactoryInterface $configFactory;

  /**
   * The language manager.
   *
   * @var \Drupal\Core\Language\LanguageManagerInterface
   * @PHPStan-var LanguageManagerInterface
   */
  protected LanguageManagerInterface $languageManager;

  /**
   * The entity repository.
   *
   * @var \Drupal\Core\Entity\EntityRepositoryInterface
   * @PHPStan-var EntityRepositoryInterface
   */
  protected EntityRepositoryInterface $entityRepository;

  /**
   * The bundle info interface.
   *
   * @var \Drupal\Core\Entity\EntityTypeBundleInfoInterface
   * @PHPStan-var EntityTypeBundleInfoInterface
   */
  protected EntityTypeBundleInfoInterface $entityTypeBundleInfo;

  /**
   * The module extension list.
   *
   * @var \Drupal\Core\Extension\ModuleExtensionList
   * @PHPStan-var ModuleExtensionList
   */
  protected ModuleExtensionList $moduleExtensionList;

  /**
   * The user permissions interface.
   *
   * @var \Drupal\user\PermissionHandlerInterface
   * @PHPStan-var PermissionHandlerInterface;
   */
  protected PermissionHandlerInterface $userPermissions;

  /**
   * The logger service.
   *
   * @var \Psr\Log\LoggerInterface
   */
  protected LoggerInterface $logger;

  /**
   * The localization unique to this project.
   *
   * @var array<string>
   * @PHPStan-var array<string>;
   */
  protected array $localization;

  /**
   * The Info Parser service.
   *
   * @var \Drupal\Core\Extension\InfoParserInterface
   */
  protected InfoParserInterface $infoParser;

  /**
   * The module handler interface.
   *
   * @var \Drupal\Core\Extension\ModuleHandlerInterface
   */
  protected ModuleHandlerInterface $moduleHandler;

  /**
   * The settings unique to this project.
   *
   * @var array<mixed>
   */
  protected array $settings;

  /**
   * A list of sites on the filesystem.
   *
   * @var array<string>
   */
  protected array $sites;

  /**
   * The site alias manager.
   *
   * @var \Consolidation\SiteAlias\SiteAliasManager
   */
  protected SiteAliasManager $siteAliasManager;

  /**
   * The path alias manager.
   *
   * @var \Drupal\path_alias\AliasManagerInterface
   */
  protected AliasManagerInterface $pathAliasManager;

  /**
   * Constructs a new RisleyExportCommands object.
   *
   * @param \Drupal\Core\Entity\EntityFieldManagerInterface $entity_field_manager
   *   The entity field manager service.
   * @param \Drupal\Core\Entity\EntityTypeManagerInterface $entity_type_manager
   *   The entity type manager service.
   * @param \Drupal\Core\Field\FieldTypePluginManagerInterface $field_type_plugin_manager
   *   The field type plugin manager service.
   * @param \Drupal\Core\Language\LanguageManagerInterface $language_manager
   *   The language manager service.
   * @param \Drupal\Core\Config\ConfigFactoryInterface $config_factory
   *   The config factory.
   * @param \Drupal\Core\Entity\EntityTypeBundleInfoInterface $entity_type_bundle_info
   *   The bundle info interface.
   * @param \Drupal\Core\Entity\EntityRepositoryInterface $entity_repository
   *   The entity repository.
   * @param \Drupal\Core\Extension\ModuleExtensionList $module_extension_list
   *   The module extension list.
   * @param \Drupal\user\PermissionHandlerInterface $user_permissions
   *   The user permissions interface.
   * @param \Psr\Log\LoggerInterface $logger
   *   The logger service.
   * @param array $localization
   *   The localization.
   * @param \Drupal\Core\Extension\InfoParserInterface $info_parser
   *   The Info Parser service.
   * @param \Drupal\Core\Extension\ModuleHandlerInterface $module_handler
   *   The module handler interface.
   * @param array $settings
   *   The settings.
   * @param \Consolidation\SiteAlias\SiteAliasManager $site_alias_manager
   *   The site alias manager.
   * @param \Drupal\path_alias\AliasManagerInterface $path_alias_manager
   *   The path alias manager.
   */
  public function __construct(
    EntityFieldManagerInterface $entity_field_manager,
    EntityTypeManagerInterface $entity_type_manager,
    FieldTypePluginManagerInterface $field_type_plugin_manager,
    LanguageManagerInterface $language_manager,
    ConfigFactoryInterface $config_factory,
    EntityTypeBundleInfoInterface $entity_type_bundle_info,
    EntityRepositoryInterface $entity_repository,
    ModuleExtensionList $module_extension_list,
    PermissionHandlerInterface $user_permissions,
    LoggerInterface $logger,
    array $localization,
    InfoParserInterface $info_parser,
    ModuleHandlerInterface $module_handler,
    array $settings,
    SiteAliasManager $site_alias_manager,
    AliasManagerInterface $path_alias_manager
    ) {
    $this->entityFieldManager = $entity_field_manager;
    $this->entityTypeManager = $entity_type_manager;
    $this->fieldTypePluginManager = $field_type_plugin_manager;
    $this->lastColumnValues = [];
    $this->skippedColumns = [];
    $this->languageManager = $language_manager;
    $this->configFactory = $config_factory;
    $this->entityTypeBundleInfo = $entity_type_bundle_info;
    $this->entityRepository = $entity_repository;
    $this->moduleExtensionList = $module_extension_list;
    $this->userPermissions = $user_permissions;
    $this->logger = $logger;
    $this->localization = $localization;
    $this->infoParser = $info_parser;
    $this->moduleHandler = $module_handler;
    $this->settings = $settings;
    $this->siteAliasManager = $site_alias_manager;
    $this->pathAliasManager = $path_alias_manager;
    $this->sites = $this->getAllSites();
  }

  /**
   * Constructs the rest of the object.
   */
  public function setSpreadsheet(Spreadsheet $spreadsheet, array $options): void {
    $this->spreadsheet = $spreadsheet;
    $this->sheet = $spreadsheet->createSheet();
    if (!isset($options['path']) || !is_string($options['path']) ||
      (!is_string($options['file']) && $options['file'] !== NULL) ||
      !isset($options['filename']) || !is_string($options['filename'])) {
      throw new \InvalidArgumentException('Invalid options structure.');
    }
    $this->options = $options;
    $this->initialize();
  }

  /**
   * Initializes the sheet.
   */
  protected function initialize(): void {
    $sheet = $this->sheet;
    $sheet->setTitle('Worksheet');

    $headers = [
      "No.",
    ];
    $sheet->fromArray($headers);

    $this->setStyle();

    $this->setBorders();

    $sheet->getColumnDimension('A')->setWidth(5);
    $sheet->getColumnDimension('B')->setWidth(20);
    $sheet->getColumnDimension('C')->setWidth(20);

    $this->setStyle();
  }

  /**
   * Builds the headers of the sheet in a uniform way.
   *
   * @param array $headers
   *   An array of arrays that corresponds to cells starting at the origin.
   * @param string $startCell
   *   The first cell of the header array.
   *
   * @throws \Exception
   *   Warns developer that they messed up.
   */
  protected function setHeaders(array $headers, string $startCell = "A1"): int {
    if (preg_match('/([A-Z]+)(\d+)/', $startCell, $matches)) {
      $startCol = $this->colToInt($matches[1]);
      $row = (int) $matches[2];
    }
    else {
      throw new \Exception("$startCell is not a valid cell format");
    }
    $hasSettings = array_reduce($headers[0], function ($carry, $item) {
      return $carry && is_numeric($item);
    }, TRUE);

    if ($hasSettings) {
      $settings = array_shift($headers);
      foreach ($settings as $i => $width) {
        $this->sheet->getColumnDimension($this->intToCol($i))->setWidth($width);
      }
    }

    foreach ($headers as $rows) {
      foreach ($rows as $i => $cell) {
        $col = $this->intToCol($startCol + $i);
        $this->setCell($this->sheet, $this->intToCol($i), $row, is_string($cell) ? $cell : $cell['value'] ?? '');
        if (is_array($cell)) {
          unset($cell['value']);
          if (!empty($cell)) {
            $this->sheet->getStyle("$col$row:$col$row")->applyFromArray($cell);
          }
        }
        else {
          $this->sheet->getStyle("$col$row}:$col$row")->applyFromArray([
            'font' => [
              'name' => 'Meiryo UI',
              'size' => 10,
            ],
            'alignment' => [
              'wrapText' => TRUE,
              'horizontal' => Alignment::HORIZONTAL_CENTER,
              'vertical' => Alignment::VERTICAL_CENTER,
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
              'right' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
              ],
              'top' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
              ],
              'bottom' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
              ],
            ],
          ]);
        }
      }
      $row++;
    }

    $this->bodyCell = $this->intToCol($startCol) . $row;
    return $row;
  }

  /**
   * Sets base styles for the sheet.
   */
  protected function setStyle(): void {
    $styleArray = [
      'font' => [
        'name' => 'Meiryo UI',
        'size' => 10,
      ],
      'alignment' => [
        'wrapText' => TRUE,
        'vertical' => Alignment::VERTICAL_TOP,
      ],
    ];
    $this->sheet->getStyle(($this->bodyCell) . ':' . $this->sheet->getHighestColumn() . $this->sheet->getHighestRow())->applyFromArray($styleArray);
  }

  /**
   * Sets borders with optional start and end cell.
   *
   * @param string|null $startCell
   *   The first cell. If null, defaults to the first line.
   * @param string|null $endCell
   *   The last cell. If null, defaults to the last line.
   */
  protected function setBorders(string|null $startCell = NULL, string|null $endCell = NULL): void {
    $sheet = $this->sheet;

    $allBordersStyle = [
      'borders' => [
        'allBorders' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FF000000'],
        ],
      ],
    ];
    $blankCellBordersStyle = [
      'borders' => [
        'left' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FF000000'],
        ],
        'right' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FF000000'],
        ],
        'bottom' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FFFFFF'],
        ],
        'top' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FFFFFF'],
        ],
      ],
    ];
    $lastBlankCellBordersStyle = [
      'borders' => [
        'left' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FF000000'],
        ],
        'right' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FF000000'],
        ],
        'bottom' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FF000000'],
        ],
        'top' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FFFFFF'],
        ],
      ],
    ];
    $nextBlankCellBordersStyle = [
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
          'color' => ['argb' => 'FF000000'],
        ],
        'bottom' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FFFFFF'],
        ],
      ],
    ];
    $bottomBorderStyle = [
      'borders' => [
        'bottom' => [
          'borderStyle' => Border::BORDER_THIN,
          'color' => ['argb' => 'FF000000'],
        ],
      ],
    ];

    $range = $sheet->calculateWorksheetDimension();
    preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $range, $matches);

    if ($startCell === NULL && !empty($this->bodyCell)) {
      $startCell = $this->bodyCell;
    }

    $startColumn = $startCell ? $startCell[0] : $matches[1];
    $startRow = $startCell ? (int) substr($startCell, 1) : (int) $matches[2];
    $endColumn = $endCell ? $endCell[0] : $matches[3];
    $endRow = $endCell ? (int) substr($endCell, 1) : (int) $matches[4];

    for ($row = $startRow; $row <= $endRow; $row++) {
      for ($col = $startColumn; ord($col) <= ord($endColumn); $col = $this->incrementColumn($col)) {
        $cellValue = $sheet->getCell($col . $row)->getValue();
        $nextCellValue = $sheet->getCell($col . ($row + 1))->getValue();

        $styleArray = (!in_array($col, $this->skippedColumns)) ? $allBordersStyle : (($cellValue == '' && $nextCellValue == '') ? $blankCellBordersStyle : (($cellValue == '') ? $lastBlankCellBordersStyle : (($nextCellValue == '') ? $nextBlankCellBordersStyle : $allBordersStyle)));
        $sheet->getStyle($col . $row)->applyFromArray($styleArray);
      }
    }

    // Set last row.
    for ($col = $startColumn; ord($col) <= ord($endColumn); $col = $this->incrementColumn($col)) {
      $sheet->getStyle($col . $endRow)->applyFromArray($bottomBorderStyle);
    }

    // Now handle the merged cells specifically.
    $mergedCells = $sheet->getMergeCells();
    foreach ($mergedCells as $mergeRange) {
      $sheet->getStyle($mergeRange)->applyFromArray($allBordersStyle);
    }

  }

  /**
   * UTILITIES.
   */

  /**
   * Gets all cells in a range.
   */
  public function getCellsInRange(string $range): array {
    [$startCell, $endCell] = explode(':', $range);

    [$startColumn, $startRow] = Coordinate::coordinateFromString($startCell);
    [$endColumn, $endRow] = Coordinate::coordinateFromString($endCell);

    $cells = [];

    for ($row = $startRow; $row <= $endRow; $row++) {
      for ($col = Coordinate::columnIndexFromString($startColumn); $col <= Coordinate::columnIndexFromString($endColumn); $col++) {
        $cells[] = Coordinate::stringFromColumnIndex($col) . $row;
      }
    }

    return $cells;
  }

  /**
   * Gets the adjacent cell.
   */
  public function getAdjacentCell(string $cell, string $side): string {
    [$col, $row] = Coordinate::coordinateFromString($cell);

    switch ($side) {
      case 'top':
        $row--;
        break;

      case 'left':
        $col = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col) - 1);
        break;

      case 'right':
        $col = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col) + 1);
        break;

      case 'bottom':
        $row++;
        break;

      default:
        return $cell;
    }

    return $col . $row;

  }

  /**
   * Merges all cells.
   */
  protected function merge(array|string $ranges):void {

    if (is_string($ranges)) {
      $ranges = [$ranges];
    }

    if (!$this->settings['merge']) {
      foreach ($ranges as $range) {
        $cells = $this->getCellsInRange($range);
        $color = $this->sheet->getStyle($cells[0])->getFill()->getStartColor()->getRGB();
        $fillStyle = [
          'fillType' => Fill::FILL_SOLID,
          'startColor' => [
            'argb' => $color,
          ],
          'endColor' => [
            'argb' => $color,
          ],
        ];
        $borderStyle = [
          'borderStyle' => BORDER::BORDER_THIN,
          'color' => ['argb' => $color],
        ];
        $alignmentStyle = [
          'wrapText' => FALSE,
          'horizontal' => Alignment::HORIZONTAL_LEFT,
          'vertical' => Alignment::VERTICAL_TOP,
        ];
        $edgeStyle = $this->sheet->getStyle($cells[0])->getBorders()->getTop()->exportArray();
        foreach ($cells as $cell) {
          $style = $this->sheet->getStyle($cell);
          $style->getFill()->applyFromArray($fillStyle);
          $style->getAlignment()->applyFromArray($alignmentStyle);
          foreach (['top', 'bottom', 'left', 'right'] as $side) {
            $adjacentCell = $this->getAdjacentCell($cell, $side);
            $method = 'get' . ucfirst($side);
            if (!in_array($adjacentCell, $cells)) {
              $style->getBorders()->$method()->applyFromArray($edgeStyle);
            }
            else {
              $style->getBorders()->$method()->applyFromArray($borderStyle);
            }
          }
        }
      }

      return;
    }

    foreach ($ranges as $range) {
      $this->sheet->mergeCells($range);
    }
  }

  /**
   * Sets a single cell.
   *
   * If lastColumnValues is defined, it skips redundant rows.
   */
  protected function setCell(Worksheet $sheet, string $column, int $row, mixed $value, bool $skipRepeats = FALSE): void {
    $sheet->setCellValue($column . $row, ($skipRepeats && $value !== ' ' && isset($this->lastColumnValues[$column]) && $this->lastColumnValues[$column] == $value) ? "" : $value);
    $this->lastColumnValues[$column] = $value;

    if ($skipRepeats) {
      $this->skippedColumns[] = $column;
    }
  }

  /**
   * Converts an int to a column letter, such as 2 -> C.
   */
  protected function intToCol(int $int): string {
    $letter = '';
    while ($int >= 0) {
      $letter = chr($int % 26 + 65) . $letter;
      $int = intdiv($int, 26) - 1;
    }
    return $letter;
  }

  /**
   * Converts a column letter back to an int, such as C -> 2.
   */
  protected function colToInt(string $col): int {
    $int = 0;
    $length = strlen($col);
    for ($i = 0; $i < $length; $i++) {
      $int *= 26;
      $int += ord($col[$i]) - 64;
    }
    return $int - 1;
  }

  /**
   * Increments alphabetic columns in excel. A->B->AA, etc.
   */
  protected function incrementColumn(string $col, int $increment = 1): string {
    $length = strlen($col);
    $result = '';
    while ($length > 0) {
      $length--;
      $c = ord($col[$length]);
      // 'Z'
      if ($c == 90) {
        $result = 'A' . $result;
        if ($length == 0) {
          $col = 'A' . $result;
          break;
        }
      }
      else {
        $col[$length] = chr($c + $increment);
        $col = substr($col, 0, $length + 1) . $result;
        break;
      }
    }
    return $col;
  }

  /**
   * Sets true or false display.
   */
  protected function buildCheck(mixed $bool, string $trueLabel = "", string $falseLabel = ""): string {
    return $bool ? "○" . $trueLabel : "-" . $falseLabel;
  }

  /**
   * Gets the default value of a field definition.
   */
  protected function getDefaultValue(BaseFieldDefinition|BaseFieldOverride|FieldDefinitionInterface $object): mixed {
    if ($object instanceof BaseFieldOverride) {
      $defaultValue = $object->get('default_value');
      if (!is_array($defaultValue) || !is_array($defaultValue[0])) {
        return '';
      }

      return $defaultValue[0]['value'];
    }
    return isset($object['default_value']) && $object['default_value'][0]['value'];
  }

  /**
   * Checks whether a pathauto pattern for the given bundle exists.
   */
  protected function hasPathautoPattern(string $bundle): bool {
    $patterns = $this->entityTypeManager->getStorage('pathauto_pattern')->loadMultiple();

    if (!is_array($patterns)) {
      return FALSE;
    }

    foreach ($patterns as $pattern) {
      $criteria = $pattern->get('selection_criteria');

      if (!is_array($criteria)) {
        continue;
      }

      foreach ($criteria as $criterion) {
        if (!is_array($criterion) || !array_key_exists('bundles', $criterion)) {
          return FALSE;
        }

        $bundles = $criterion['bundles'];

        if (is_array($bundles) && in_array($bundle, $bundles)) {
          return TRUE;
        }
      }
    }

    return FALSE;
  }

  /**
   * Checks whether the given entity type is used in a workflow.
   */
  protected function hasWorkflow(string $entityTypeId, string $bundle): bool {
    $workflowConfigNames = $this->configFactory->listAll('workflows.workflow.');
    foreach ($workflowConfigNames as $configName) {
      $config = $this->configFactory->get($configName);
      $typeSettings = $config->get('type_settings');
      if (!is_array($typeSettings)) {
        continue;
      }

      $entityTypes = $typeSettings['entity_types'];
      if (!is_array($entityTypes)) {
        continue;
      }

      $entityTypes = $entityTypes[$entityTypeId];
      if (is_array($entityTypes) && in_array($bundle, $entityTypes)) {
        return TRUE;
      }
    }

    return FALSE;
  }

  /**
   * Checks whether the given entity type is translatable.
   */
  protected function canTranslate(string $entityTypeId, string $bundle): bool {
    $config = $this->configFactory->get("language.content_settings.$entityTypeId.$bundle");

    $thirdPartySettings = $config->get('third_party_settings');
    if (!is_array($thirdPartySettings)) {
      return FALSE;
    }

    $contentTranslation = $thirdPartySettings['content_translation'];
    if (!is_array($contentTranslation)) {
      return FALSE;
    }

    if (!isset($contentTranslation['enabled'])) {
      return FALSE;
    }

    return $contentTranslation['enabled'];
  }

  /**
   * Checks whether all the entities of a given entity type are translatable.
   */
  protected function canTranslateAll(string $entityTypeId): bool {
    $allBundles = $this->entityTypeBundleInfo->getBundleInfo($entityTypeId);

    foreach (array_keys($allBundles) as $bundle) {
      if ($bundle == "from_library") {
        continue;
      }

      if (!$this->canTranslate($entityTypeId, $bundle)) {
        // If any bundle is not translatable, return false.
        return FALSE;
      }
    }

    // If all bundles are translatable, return true.
    return TRUE;
  }

  /**
   * Checks whether the given entity type has any language settings enabled.
   */
  protected function hasLanguageSettings(string $entityTypeId, string $bundle): bool {
    return $this->getLanguageSettings($entityTypeId, $bundle) !== '';
  }

  /**
   * Gets any enabled language settings on the entity type.
   */
  protected function getLanguageSettings(string $entityTypeId, string $bundle): string {

    $config = $this->configFactory->get("language.content_settings.$entityTypeId.$bundle");
    $ret = [];

    // Returns true if default language has been declared.
    if (($value = $config->get('default_langcode')) !== 'und' && is_string($value)) {
      $ret[] = "default_langcode: $value";
    }

    // Returns true if users are able to alter the original language.
    $ret[] = "language_alterable: " . ($config->get('language_alterable') ? 'true' : 'false');

    // The hide flag is on by default, so returns true if it is missing.
    $thirdPartySettings = $config->get('third_party_settings');
    if (!is_array($thirdPartySettings) ||
          !isset($thirdPartySettings['content_translation']) ||
          !is_array($thirdPartySettings['content_translation']) ||
          !isset($thirdPartySettings['content_translation']['bundle_settings']) ||
          !is_array($thirdPartySettings['content_translation']['bundle_settings']) ||
          (!isset($thirdPartySettings['content_translation']['bundle_settings']['untranslatable_fields_hide']) ||
              $thirdPartySettings['content_translation']['bundle_settings']['untranslatable_fields_hide'] !== '1')) {
      $ret[] = "Show untranslatable fields on translation forms";
    }

    return implode("\n", $ret);
  }

  /**
   * Checks whether the given entity type is displayed on the simple sitemap.
   */
  protected function isShownOnSimpleSitemap(string $entityTypeId, string $bundle): bool {
    return (bool) $this->configFactory->get("simple_sitemap.bundle_settings.default.$entityTypeId.$bundle")->get('index');
  }

  /**
   * Checks if any menus are enabled.
   */
  protected function hasEnabledMenus(NodeType $entity): bool {
    return $this->getEnabledMenus($entity) !== '';
  }

  /**
   * Returns a list of menus for display.
   */
  protected function getEnabledMenus(NodeType $entity): string {
    $settings = $entity->get('third_party_settings');
    if (!is_array($settings)) {
      return '';
    }

    $menuUi = $settings['menu_ui'];
    if (!isset($menuUi)) {
      return '';
    }

    $availableMenus = $menuUi['available_menus'];
    if (!isset($availableMenus)) {
      return '';
    }

    return implode($availableMenus);
  }

  /**
   * Returns rabbit hole action setting for display.
   */
  protected function getRabbitholeSetting(string $entityCategory, string $bundle): mixed {
    return $this->configFactory->get("rabbit_hole.behavior_settings.{$entityCategory}_$bundle")->get('action');
  }

  /**
   * Gets a list of all core and contrib modules.
   *
   * See https://github.com/drush-ops/drush/blob/12.x/
   * src/Commands/pm/PmCommands.php for full implementation.
   *
   * @param string $origin
   *   Filter the list by origin: 'all', 'core', 'contrib', or 'custom'.
   * @param \Drupal\Core\Extension\Extension[]|array[]|null $modules
   *   A list of extensions. If not set, get from this site.
   *   If set, probably from other site.
   *
   * @return \Drupal\Core\Extension\Extension[]|array[]
   *   An array of modules filtered by the specified origin.
   */
  protected function getModules(string $origin = 'all', array $modules = NULL): array {
    if (!isset($modules)) {
      $modules = $this->moduleExtensionList->getList();
    }
    foreach ($modules as $key => $module) {
      if ($module instanceof Extension) {
        if (strpos($module->getPath(), 'tests')) {
          continue;
        }

        $moduleOrigin = property_exists($module, 'origin') ? $module->origin : '';

        if ($origin !== 'all' && (
            ($origin === 'core' && $moduleOrigin !== 'core') ||
            ($origin === 'contrib' && ($moduleOrigin === 'core' || str_contains($module->getPath(), 'contrib'))) ||
            ($origin === 'custom' && str_contains($module->getPath(), 'custom'))
          )) {
          unset($modules[$key]);
        }
      }
      else {
        if (strpos($module['subpath'], 'tests')) {
          continue;
        }

        $moduleOrigin = $module['origin'] ?? '';

        if ($origin !== 'all' && (
          ($origin === 'core' && $moduleOrigin !== 'core') ||
          ($origin === 'contrib' && ($moduleOrigin === 'core' || str_contains($module['subpath'], 'contrib'))) ||
          ($origin === 'custom' && str_contains($module['subpath'], 'custom'))
        )) {
          unset($modules[$key]);
        }

      }
    }

    return $modules;
  }

  /**
   * Gets an alphabetically sorted array of permission roles.
   */
  protected function getPermissionRoles(): array {
    $roles = $this->entityTypeManager->getStorage('user_role')->loadMultiple();

    if (isset($this->settings['hideAdminister']) && $this->settings['hideAdminister'] && isset($roles['administrator'])) {
      unset($roles['administrator']);
    }

    uasort($roles, function ($a, $b) {
      return $a->get('weight') <=> $b->get('weight');
    });

    return $roles;
  }

  /**
   * Translates the string first from db and then from module.
   */
  protected function translate(EntityInterface|string|NULL $mixed, string|NULL $default = NULL): string {
    if (!$mixed) {
      return $default ?? '';
    }
    elseif ($mixed instanceof EntityInterface) {
      $original = $mixed->label();
      $translation = $this->entityRepository->getTranslationFromContext($mixed, 'ja')?->label();
      if (isset($original) && isset($translation) && $translation !== $mixed->label()) {
        return (string) $translation;
      }
      $string = $original;
    }
    else {
      $string = (string) $mixed;
    }

    if (!isset($string)) {
      return $default ?? '';
    }

    if (isset($this->settings['localization']) && !$this->settings['localization']) {
      return $default ?? (string) $string;
    }

    if (array_key_exists($string, $this->localization)) {
      return $this->localization[$string] ?: $default ?? $string;
    }

    // Key does not exist, so add it to the sheet.
    $filePath = DRUPAL_ROOT . '/modules/custom/risley_export/src/Sheets/Localization/Localization.xlsx';

    try {
      $spreadsheet = IOFactory::load($filePath);
      $sheet = $spreadsheet->getActiveSheet();
      $newRow = $sheet->getHighestRow() + 1;
      $sheet->setCellValue('A' . $newRow, $string);
      $writer = new Xlsx($spreadsheet);
      $writer->save($filePath);
      return $default ?? $string;
    }
    catch (\Exception $e) {
      $this->logger->error("Error updating localization file: @error", ['@error' => $e->getMessage()]);
    }
    return $default ?? $string;
  }

  /**
   * Gets the translated module description.
   */
  protected function getModuleDescription(Extension|array $module):string {
    if (is_object($module) && method_exists($module, 'getPath')) {
      $infoFilePath = $module->getPath() . '/' . $module->getName() . '.info.yml';
    }
    elseif (is_array($module) && isset($module['subpath'])) {
      $infoFilePath = $module['subpath'] . '/' . $module['machine_name'] . '.info.yml';
    }
    else {
      return '';
    }

    if (file_exists($infoFilePath)) {
      $description = strip_tags($this->infoParser->parse($infoFilePath)['description'] ?: '');
      return $this->translate($description);
    }
    return '';
  }

  /**
   * Changes string to kebab case.
   */
  protected function toKebabCase(string $string):string {
    return trim(preg_replace('/[^a-z0-9]+/', '-', strtolower($string)) ?? '', '-');
  }

  /**
   * Centers stuff.
   */
  protected function setStyleCenter(array|string $range):void {
    if (is_array($range)) {
      foreach ($range as $_range) {
        $this->setStyleHeader($_range);
      }
      return;
    }

    $centerAlignmentStyle = [
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
    ];
    $this->sheet->getStyle($range)->applyFromArray($centerAlignmentStyle);
  }

  /**
   * Headers stuff.
   */
  protected function setStyleHeader(string|array $range):void {
    if (is_array($range)) {
      foreach ($range as $_range) {
        $this->setStyleHeader($_range);
      }
      return;
    }

    $style = [
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'startColor' => [
          'argb' => 'FFBFBFBF',
        ],
      ],
    ];
    $this->sheet->getStyle($range)->applyFromArray($style);
  }

  /**
   * Gets all sites in the /site directory and returns an array.
   *
   * If includeCurrent is false, does not include the site running this command.
   */
  protected function getAllSites():array {
    return array_filter(array_keys($this->siteAliasManager->getMultiple() ?: []), function ($site) {
      $siteObj = $this->siteAliasManager->get($site);
      return $siteObj && !empty($siteObj->uri());
    });
  }

  /**
   * Gets all modules from all sites.
   *
   * @param string $origin
   *   Filter the list by origin: 'all', 'core', 'contrib', or 'custom'.
   */
  protected function getModulesAcrossSites(string $origin = 'all'):array|NULL {
    return array_reduce($this->sites, function ($result, $site) use ($origin) {
      $siteObj = $this->siteAliasManager->get($site);
      if (!$siteObj) {
        return $result;
      }

      $uri = $siteObj->uri();
      if (empty($uri)) {
        return $result;
      }
      $command = "/opt/drupal/vendor/bin/drush --uri=\"$uri\" ev \"echo json_encode(\\Drupal::service('extension.list.module')->getList())\"";
      $jsonModules = shell_exec($command);

      if (!is_string($jsonModules)) {
        return $result;
      }

      $modules = json_decode($jsonModules, TRUE);

      if (!is_array($modules)) {
        return $result;
      }

      foreach ($modules as $machineName => &$module) {
        if (!is_array($module)) {

          continue;
        }
        // Add machine name to array for future indexing.
        $module['machine_name'] = $machineName;
      }

      $result[$site] = $this->getModules($origin, $modules);

      return $result;
    }, []);
  }

  /**
   * Gets the status of the module across all sites found.
   *
   * @param \Drupal\Core\Extension\Extension|array $module
   *   An extension if accessing site internally or an array
   *   if accessing all sites through hacky drush command.
   */
  protected function getModuleStatus(Extension|array $module): string {
    if ($module instanceof Extension) {
      // Simply check for this site.
      return $this->buildCheck($this->moduleHandler->moduleExists($module->getName()));
    }
    elseif (isset($this->modules)) {
      $enabledSites = [];
      $machineName = $module['machine_name'];
      foreach ($this->modules as $site => $modules) {
        if (isset($modules[$machineName]) && $modules[$machineName]['status'] > 0) {
          $enabledSites[] = $site;
        }
      }

      if (empty($enabledSites)) {
        return $this->buildCheck(FALSE);
      }
      elseif (count($this->sites) === count($enabledSites)) {
        return $this->buildCheck(TRUE);
      }
      else {
        $string = array_map(function ($site) {
          $siteWithoutAt = str_replace('@', '', $site);
          $parts = explode('.', $siteWithoutAt);
          $firstPart = $parts[0];
          return strtoupper($firstPart);
        }, $enabledSites);
        return implode(', ', $string);
      }
    }
    else {
      return 'ERR';
    }
  }

  /**
   * Gets the status of the webform across all sites found.
   *
   * @param \Drupal\webform\Entity\Webform|array $webform
   *   A webform if accessing site internally or an array
   *   if accessing all sites through hacky drush command.
   */
  protected function getWebformStatus(Webform|array $webform): string {
    if ($webform instanceof Webform) {
      // Simply check for this site.
      return $this->buildCheck($webform->isOpen());
    }
    elseif (isset($this->webforms)) {
      $enabledSites = [];
      $machineName = $webform['machine_name'];
      foreach ($this->webforms as $site => $webforms) {
        if (isset($webforms[$machineName]) && $webforms[$machineName]['status'] === 'open') {
          $enabledSites[] = $site;
        }
      }

      if (empty($enabledSites)) {
        return $this->buildCheck(FALSE);
      }
      elseif (count($this->sites) === count($enabledSites)) {
        return $this->buildCheck(TRUE);
      }
      else {
        $string = array_map(function ($site) {
          $siteWithoutAt = str_replace('@', '', $site);
          $parts = explode('.', $siteWithoutAt);
          $firstPart = $parts[0];
          return strtoupper($firstPart);
        }, $enabledSites);
        return implode(', ', $string);
      }
    }
    else {
      return 'ERR';
    }
  }

  /**
   * Gets a master array of modules across all found sites.
   */
  protected function getAllModules():array {
    if (!isset($this->modules) || !is_array($this->modules)) {
      return [];
    }

    $masterArray = [];
    foreach ($this->modules as $modules) {
      foreach ($modules as $module) {
        $machineName = $module['machine_name'];

        if (!isset($masterArray[$machineName])) {
          $masterArray[$machineName] = $module;
        }
      }
    }

    return $masterArray;
  }

  /**
   * Gets a master array of webforms across all found sites.
   */
  protected function getAllWebforms():array {
    if (!isset($this->webforms) || !is_array($this->webforms)) {
      return [];
    }

    $masterArray = [];
    foreach ($this->webforms as $webforms) {
      foreach ($webforms as $webform) {
        $machineName = $webform['machine_name'];

        if (!isset($masterArray[$machineName])) {
          $masterArray[$machineName] = $webform;
        }
      }
    }

    return $masterArray;
  }

  /**
   * Gets the field type such as 'Number' or 'Boolean'.
   */
  protected function getFieldTypeLabel(FieldDefinitionInterface $fieldDefinition): string {
    $fieldType = $fieldDefinition->getType();
    $fieldSettings = $fieldDefinition->getSettings();
    $label = is_array($_ = $this->fieldTypePluginManager->getDefinition($fieldType)) ? $_['label'] : '';

    if ($fieldType === 'entity_reference' || $fieldType === 'entity_reference_revisions') {
      $entityType = $fieldSettings['target_type'] ?: $fieldSettings['handler'];
      $entityType = explode(':', $entityType);
      $entityType = end($entityType);
      return (string) $this->entityTypeManager->getDefinition($entityType)?->getLabel();
    }

    if ($fieldType === 'file') {
      $uriScheme = $fieldSettings['uri_scheme'];
      $label = "$label ($uriScheme)";
    }

    return $label;
  }

  /**
   * Gets all webforms from all sites.
   */
  protected function getWebformsAcrossSites():array|NULL {
    return array_reduce($this->sites, function ($result, $site) {
      // $command = "/opt/drupal/vendor/bin/drush $site ev 'echo json_encode(array_map(function(\$webform) { return \$webform->toArray(); }, \\Drupal::service(\"entity_type.manager\")->getStorage(\"webform\")->loadMultiple()));'";
      $siteObj = $this->siteAliasManager->get($site);

      if (!$siteObj) {
        return $result;
      }

      $uri = $siteObj->uri();
      if (empty($uri)) {
        return $result;
      }

        $command = <<<EOT
          /opt/drupal/vendor/bin/drush --uri="$uri" ev '
          if(!\\Drupal::service("entity_type.manager")->hasDefinition("webform")) return [];
          \$webforms = \\Drupal::service("entity_type.manager")->getStorage("webform")->loadMultiple();
          \$webformsArray = array_map(function(\$webform) {
              \$array = \$webform->toArray();
              \$array["_url_alias"] = \\Drupal::service("path_alias.manager")->getAliasByPath("/webform/" . \$webform->id());
              return \$array;
          }, \$webforms);
          echo json_encode(\$webformsArray);
          '
          EOT;

        $jsonWebforms = shell_exec($command);

      if (!is_string($jsonWebforms)) {
        return $result;
      }

      $webforms = json_decode($jsonWebforms, TRUE);

      if (!is_array($webforms)) {
        return $result;
      }

      foreach ($webforms as $machineName => &$webform) {
        if (!is_array($webform)) {
          continue;
        }
        // Add machine name to array for future indexing.
        $webform['machine_name'] = $machineName;
      }

      $result[$site] = $webforms;
      return $result;
    }, []);
  }

}
