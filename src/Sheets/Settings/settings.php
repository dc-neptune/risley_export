<?php

/**
 * @file
 * A settings file.
 *
 * This file applies optional settings. Rename it settings.php to use it.
 *
 * Comment out whatever you don't want.
 *
 * This file will be overwritten within /modules when updated or rebuilt.
 * If you make any edits to it, it should be moved and saved at
 * /sites/settings/risley_export.settings.php
 *
 * todo: Choose a more appropriate permanent save location.
 *
 * If this no longer needs to be run from drush as well as installed,
 * then settings.php can be rebuilt as a typical configurable within
 * the GUI. But so long as it needs to be run from drush without
 * being installed on the site, this seems necessary.
 */

/*
 * Hides administer on the roles, permissions, and other sheets.
 */
$settings['hideAdminister'] = TRUE;

/*
 * Hides roles with no permissions enabled. Useful to avoid
 * sending unnecessary roles to clients.
 */
$settings['hideEmptyPermissions'] = TRUE;

/*
 * A list of permissions that will be shown regardless of
 * whether they are empty. Key value is whatever is output
 * in "Omitting empty permission".
 */
$settings['whiteListPermissions'] = [];

/*
 * Moves the listed groups to the top of the Permissions sheet
 * before sorting the rest alphabetically.
 *
 * Searches by label for no good reason.
 */
$settings['priorityPermissions'] = ['Node', 'Media'];

/*
 * Hides read-only fields in the Fields sheet.
 */
$settings['hideReadOnly'] = TRUE;

/*
 * Hides particular fields in the fields sheet.
 */
$settings['blackListFields'] = [];

/*
 * Orders fields in the Fields sheet according
 * to their weight in the form display
 * todo: Implement this later. It seems annoying to
 * deal with multiple content types and checks.
 */
$settings['weightedFields'] = TRUE;

/*
 * If set to true, add a localization file to Localization.
 */
$settings['localization'] = FALSE;


/*
 * If set to true, allows merging cells for a more
 * beautiful end state.
 */
$settings['merge'] = FALSE;


/*
 * If set to true, exports as a CSV
 *
 * Other tested types: 'xlsx', 'xls'
 */
$settings['filetype'] = 'csv';
