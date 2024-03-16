<?php

/**
 * @file
 * A settings file.
 *
 * This file applies optional settings. Rename it settings.php to use it.
 *
 * Comment out whatever you don't want.
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
