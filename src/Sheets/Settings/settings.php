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
 * Moves the listed groups to the top of the Permissions sheet
 * before sorting the rest alphabetically.
 *
 * Searches by label for no good reason.
 */
$settings['priorityPermissions'] = ['Node', 'Media'];
