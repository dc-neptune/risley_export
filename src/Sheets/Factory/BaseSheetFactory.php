<?php

namespace Drupal\risley_export\Sheets\Factory;

use Consolidation\SiteAlias\SiteAliasManager;
use Drupal\Core\Config\ConfigFactoryInterface;
use Drupal\Core\Entity\EntityFieldManagerInterface;
use Drupal\Core\Entity\EntityRepositoryInterface;
use Drupal\Core\Entity\EntityTypeBundleInfoInterface;
use Drupal\Core\Entity\EntityTypeManagerInterface;
use Drupal\Core\Extension\InfoParserInterface;
use Drupal\Core\Extension\ModuleExtensionList;
use Drupal\Core\Extension\ModuleHandlerInterface;
use Drupal\Core\Field\FieldTypePluginManagerInterface;
use Drupal\Core\Language\LanguageManagerInterface;
use Drupal\Core\Logger\LoggerChannelFactoryInterface;
use Drupal\user\PermissionHandlerInterface;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * A factory for BaseSheet.
 */
class BaseSheetFactory {

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
   * The module handler interface.
   *
   * @var \Drupal\Core\Extension\ModuleHandlerInterface
   */
  protected $moduleHandler;

  /**
   * The settings unique to this project.
   *
   * @var array<mixed>
   */
  protected $settings;

  /**
   * The site alias manager.
   *
   * @var \Consolidation\SiteAlias\SiteAliasManager
   */
  protected $siteAliasManager;

  /**
   * Constructs a new BaseSheetFactory object.
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
   * @param \Drupal\Core\Logger\LoggerChannelFactoryInterface $logger_factory
   *   The logger factory.
   * @param \Drupal\Core\Extension\InfoParserInterface $info_parser
   *   The Info Parser service.
   * @param \Drupal\Core\Extension\ModuleHandlerInterface $module_handler
   *   The module handler interface.
   * @param \Consolidation\SiteAlias\SiteAliasManager $site_alias_manager
   *   The site alias manager.
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
    LoggerChannelFactoryInterface $logger_factory,
    InfoParserInterface $info_parser,
    ModuleHandlerInterface $module_handler,
    SiteAliasManager $site_alias_manager
  ) {
    $this->entityFieldManager = $entity_field_manager;
    $this->entityTypeManager = $entity_type_manager;
    $this->fieldTypePluginManager = $field_type_plugin_manager;
    $this->languageManager = $language_manager;
    $this->configFactory = $config_factory;
    $this->entityTypeBundleInfo = $entity_type_bundle_info;
    $this->entityRepository = $entity_repository;
    $this->moduleExtensionList = $module_extension_list;
    $this->userPermissions = $user_permissions;
    $this->logger = $logger_factory->get('risley_export');
    $this->localization = $this->buildLocalization();
    $this->infoParser = $info_parser;
    $this->moduleHandler = $module_handler;
    $this->settings = $this->buildSettings();
    $this->siteAliasManager = $site_alias_manager;
  }

  /**
   * Creates an extension of BaseSheet.
   */
  public function create(Spreadsheet $spreadsheet, string $key, array $options):Worksheet|NULL {
    $className = 'Drupal\\risley_export\\Sheets\\' . $key . 'Sheet';

    if (!class_exists($className)) {
      return NULL;
    }

    /** @var \Drupal\risley_export\Sheets\BaseSheet $customSheet */
    $customSheet = new $className(
      $this->entityFieldManager,
      $this->entityTypeManager,
      $this->fieldTypePluginManager,
      $this->languageManager,
      $this->configFactory,
      $this->entityTypeBundleInfo,
      $this->entityRepository,
      $this->moduleExtensionList,
      $this->userPermissions,
      $this->logger,
      $this->localization,
      $this->infoParser,
      $this->moduleHandler,
      $this->settings,
      $this->siteAliasManager
    );
    $customSheet->setSpreadsheet($spreadsheet, $options);
    return $customSheet->sheet;
  }

  /**
   * Converts the Localization.xlsx into a serviceable array.
   */
  public function buildLocalization(): array {
    $localization = [];
    $filePath = DRUPAL_ROOT . '/modules/custom/risley_export/src/Sheets/Localization/Localization.xlsx';

    // Check if file exists.
    if (!file_exists($filePath)) {
      // Create a new Spreadsheet object.
      $spreadsheet = new Spreadsheet();
      $spreadsheet->getProperties()
        ->setCreator("Digital Circus")
        ->setTitle("Localization");

      // Create a writer instance and save the file.
      $writer = new Xlsx($spreadsheet);
      try {
        $writer->save($filePath);
      }
      catch (Exception $e) {
        $this->logger->error("Error creating localization file: @error", ['@error' => $e->getMessage()]);
        // Return early if file cannot be created.
        return $localization;
      }
    }

    try {
      $spreadsheet = IOFactory::load($filePath);
      $sheet = $spreadsheet->getActiveSheet();

      $duplicateRows = [];
      foreach ($sheet->getRowIterator() as $row) {
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(FALSE);
        $cells = [];
        foreach ($cellIterator as $cell) {
          $cells[] = $cell->getValue();
        }

        if (isset($cells[0])) {
          if (isset($localization[$cells[0]])) {
            $duplicateRows[] = $row->getRowIndex();
          }
          else {
            $localization[$cells[0]] = $cells[1] ?? '';
          }
        }
      }
      // Validate file
      // Reverse sort duplicateRows to avoid messing up
      // row numbers while deleting.
      rsort($duplicateRows);
      foreach ($duplicateRows as $rowNum) {
        $sheet->removeRow($rowNum);
      }
      $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
      $writer->save($filePath);
    }
    catch (Exception $e) {
      // Handle exception.
      $this->logger->error("Error loading localization file: @error", ['@error' => $e->getMessage()]);
    }

    return $localization;
  }

  /**
   * Converts the settings.php.
   */
  public function buildSettings(): array {
    $settings = [];
    $filePath = DRUPAL_ROOT . '/modules/custom/risley_export/src/Sheets/Settings/settings.php';

    if (file_exists($filePath)) {
      include $filePath;
    }

    return $settings;
  }

}
