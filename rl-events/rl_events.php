<?php
/**
 * @package RL Events
 */
/*
Plugin Name: RL Events
Plugin URI: https://resistencialudica.com/
Description: Plugin de Resistencia Lúdica para crear eventos fácilmente
Version: 0.0.1
Author: Roberto M. Oliva
Author URI:
License: GPLv2 or later
Text Domain: rl_events
*/

defined('ABSPATH') or die("Bye bye");

define('RL_EVENTS_PATH', plugin_dir_path(__FILE__));

include(plugin_dir_path(__FILE__).'/src/lib/defines.php');
include(plugin_dir_path(__FILE__).'/src/lib/assets.php');
include(plugin_dir_path(__FILE__).'/src/lib/menu.php');
include(plugin_dir_path(__FILE__).'/src/lib/permissions.php');
include(plugin_dir_path(__FILE__).'/src/lib/dates.php');
include(plugin_dir_path(__FILE__).'/src/lib/create_event.php');
include(plugin_dir_path(__FILE__).'/src/lib/classes/db.php');
include(plugin_dir_path(__FILE__).'/src/lib/classes/template.php');
include(plugin_dir_path(__FILE__).'/src/lib/classes/game_data/month_years.php');
include(plugin_dir_path(__FILE__).'/src/lib/classes/game_data/month_dates.php');
include(plugin_dir_path(__FILE__).'/src/lib/classes/game_data/games.php');
include(plugin_dir_path(__FILE__).'/src/lib/classes/game_data/game/capacity.php');
include(plugin_dir_path(__FILE__).'/src/lib/classes/game_data/game/attendees.php');
include(plugin_dir_path(__FILE__).'/src/lib/classes/game_data/game/organizers.php');
