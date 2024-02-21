<?php

namespace Drupal\risley_export\Sheets;

use Drupal\Core\Config\ConfigFactoryInterface;
use Drupal\Core\Entity\EntityFieldManagerInterface;
use Drupal\Core\Entity\EntityInterface;
use Drupal\Core\Entity\EntityRepositoryInterface;
use Drupal\Core\Entity\EntityTypeBundleInfoInterface;
use Drupal\Core\Entity\EntityTypeManagerInterface;
use Drupal\Core\Extension\Extension;
use Drupal\Core\Extension\InfoParserInterface;
use Drupal\Core\Extension\ModuleExtensionList;
use Drupal\Core\Field\BaseFieldDefinition;
use Drupal\Core\Field\Entity\BaseFieldOverride;
use Drupal\Core\Field\FieldDefinitionInterface;
use Drupal\Core\Field\FieldTypePluginManagerInterface;
use Drupal\Core\Language\LanguageManagerInterface;
use Drupal\node\Entity\NodeType;
use Drupal\user\PermissionHandlerInterface;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
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
    protected $spreadsheet;

    /**
     * The current sheet being built.
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public $sheet;

    /**
     * Tracks the last value in each column.
     *
     * @var array<mixed>
     */
    protected $lastColumnValues;

    /**
     * Tracks whether the column has been skipped.
     *
     * @var array<string>
     */
    protected $skippedColumns;

    /**
     * The options passed through the drush command.
     *
     * @var array<mixed>
     */
    protected $options;

    /**
     * The entity field manager.
     *
     * @var \Drupal\Core\Entity\EntityFieldManagerInterface
     * @PHPStan-var EntityFieldManagerInterface
     */
    protected $entityFieldManager;

    /**
     * The entity type manager.
     *
     * @var \Drupal\Core\Entity\EntityTypeManagerInterface
     * @PHPStan-var EntityTypeManagerInterface
     */
    protected $entityTypeManager;

    /**
     * The field type plugin manager.
     *
     * @var \Drupal\Core\Field\FieldTypePluginManagerInterface
     * @PHPStan-var FieldTypePluginManagerInterface
     */
    protected $fieldTypePluginManager;


    /**
     * The config factory.
     *
     * @var \Drupal\Core\Config\ConfigFactoryInterface
     * @PHPStan-var ConfigFactoryInterface
     */
    protected $configFactory;

    /**
     * The language manager.
     *
     * @var \Drupal\Core\Language\LanguageManagerInterface
     * @PHPStan-var LanguageManagerInterface
     */
    protected $languageManager;

    /**
     * The entity repository.
     *
     * @var \Drupal\Core\Entity\EntityRepositoryInterface
     * @PHPStan-var EntityRepositoryInterface
     */
    protected $entityRepository;

    /**
     * The bundle info interface.
     *
     * @var \Drupal\Core\Entity\EntityTypeBundleInfoInterface
     * @PHPStan-var EntityTypeBundleInfoInterface
     */
    protected $entityTypeBundleInfo;

    /**
     * The module extension list.
     *
     * @var \Drupal\Core\Extension\ModuleExtensionList
     * @PHPStan-var ModuleExtensionList
     */
    protected $moduleExtensionList;

    /**
     * The user permissions interface.
     *
     * @var \Drupal\user\PermissionHandlerInterface
     * @PHPStan-var PermissionHandlerInterface;
     */
    protected $userPermissions;

    /**
     * The logger service.
     *
     * @var \Psr\Log\LoggerInterface
     */
    protected $logger;

    /**
     * The localization unique to this project.
     *
     * @var array<string>
     * @PHPStan-var array<string>;
     */
    protected $localization;

    /**
     * The Info Parser service.
     *
     * @var \Drupal\Core\Extension\InfoParserInterface
     */
    protected $infoParser;

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
        InfoParserInterface $info_parser
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
    }

    /**
     * Constructs the rest of the object.
     */
    public function setSpreadsheet(Spreadsheet $spreadsheet, array $options): void {
        $this->spreadsheet = $spreadsheet;
        $this->sheet = $spreadsheet->createSheet();
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
        $sheet->fromArray($headers, NULL, 'A1');

        $this->setStyle($sheet);

        $this->setBorders();

        $sheet->getColumnDimension('A')->setWidth(5);
        $sheet->getColumnDimension('B')->setWidth(20);
        $sheet->getColumnDimension('C')->setWidth(20);

        $this->setStyle($sheet);
    }

    /**
     * Sets base styles for the sheet.
     */
    protected function setStyle(Worksheet $sheet, int|string $headerRow = 1): void {
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
        $sheet->getStyle($sheet->calculateWorksheetDimension())->applyFromArray($styleArray);

        $headerStyleArray = [
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

        $range = NULL;
        if (is_int($headerRow)) {

            $lastColumn = $sheet->getHighestDataColumn($headerRow);
            $range = "A{$headerRow}:{$lastColumn}{$headerRow}";
        }
        else {
            $range = $headerRow;
        }

        $sheet->getStyle($range)->applyFromArray($headerStyleArray);
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

        if ($startCell === NULL || $endCell === NULL) {
            $range = $sheet->calculateWorksheetDimension();
            preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $range, $matches);
            $startColumn = $startCell ? $startCell[0] : $matches[1];
            $startRow = $startCell ? (int) substr($startCell, 1) : (int) $matches[2];
            $endColumn = $endCell ? $endCell[0] : $matches[3];
            $endRow = $endCell ? (int) substr($endCell, 1) : (int) $matches[4];
        }
        else {
            $startColumn = $startCell[0];
            $startRow = (int) substr($startCell, 1);
            $endColumn = $endCell[0];
            $endRow = (int) substr($endCell, 1);
        }

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

    }

    /*
     * UTILITIES
     */

    /**
     * Sets a single cell.
     *
     * If lastColumnValues is defined, it skips redundant rows.
     */
    protected function setCell(Worksheet $sheet, string $column, int $row, mixed $value, bool $skipRepeats = FALSE): void {
        $sheet->setCellValue($column . $row, ($skipRepeats && isset($this->lastColumnValues[$column]) && $this->lastColumnValues[$column] == $value) ? "" : $value);
        $this->lastColumnValues[$column] = $value;

        if ($skipRepeats) {
            $this->skippedColumns[] = $column;
        }
    }

    /**
     * Increments alphabetic columns in excel. A->B->AA, etc.
     */
    protected function incrementColumn(string $col): string {
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
                $col[$length] = chr($c + 1);
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
                if (!is_array($criterion)) {
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
     *
     * @return array
     *   An array of modules filtered by the specified origin.
     */
    protected function getModules(string $origin = 'all'): array {
        /** @var \Drupal\Core\Extension\Extension[] $modules */
        $modules = $this->moduleExtensionList->getList();

        foreach ($modules as $key => $module) {
            if (strpos($module->getPath(), 'tests')) {
                continue;
            }

            $moduleOrigin = property_exists($module, 'origin') ? $module->origin : '';

            if ($origin !== 'all' && (
                    ($origin === 'core' && $moduleOrigin !== 'core') ||
                    ($origin === 'contrib' && ($moduleOrigin === 'core' || strpos($module->getPath(), 'contrib') === FALSE)) ||
                    ($origin === 'custom' && strpos($module->getPath(), 'custom') === FALSE)
                )) {
                unset($modules[$key]);
            }
        }

        return $modules;
    }

    /**
     * Gets an alphabetically sorted array of permission roles.
     */
    protected function getPermissionRoles(): array {
        $roles = $this->entityTypeManager->getStorage('user_role')->loadMultiple();
        ksort($roles);
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
            $translation = $this->entityRepository->getTranslationFromContext($mixed, 'ja')->label();
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

        if (array_key_exists($string, $this->localization)) {
            return $this->localization[$string] ?: $default ?? $string;
        }

        // Key does not exist, so add it to the sheet.
        $filePath = DRUPAL_ROOT . '/modules/custom/risley_export/src/Sheets/Localization/Localization.xls';

        try {
            $spreadsheet = IOFactory::load($filePath);
            $sheet = $spreadsheet->getActiveSheet();
            $newRow = $sheet->getHighestRow() + 1;
            $sheet->setCellValue('A' . $newRow, $string);
            $writer = new Xls($spreadsheet);
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
    protected function getModuleDescription(Extension $module):string {
        $infoFilePath = $module->getPath() . '/' . $module->getName() . '.info.yml';
        if (file_exists($infoFilePath)) {
            $description = strip_tags($this->infoParser->parse($infoFilePath)['description'] ?: '');
            return $this->translate($description);
        }
        return '';
    }

}
