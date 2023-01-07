<?php

defined('ABSPATH') or die("Bye bye");

function rl_events_template($post) {
  return new RLEvents_Template($post);
}

function rl_events_create_event_start_date($post) {
  return rl_events_datetime_start($post)->format('Y-m-d');
}

function rl_events_create_event_end_date($post) {
  return rl_events_datetime_end($post)->format('Y-m-d');
}
function rl_events_create_event_start_hour($post) {
  return array(
    'hour' => rl_events_datetime_start($post)->format('H'),
    'minute' => rl_events_datetime_start($post)->format('i'),
    'meridian' => rl_events_datetime_start($post)->format('a'),
  );
}

function rl_events_create_event_end_hour($post) {
  return array(
    'hour' => rl_events_datetime_end($post)->format('H'),
    'minute' => rl_events_datetime_end($post)->format('i'),
    'meridian' => rl_events_datetime_end($post)->format('a'),
  );
}

// array(13) {
// 	["date_publish"]=> string(19) "2022/06/06 20:00:00"
// 	["date_start"]=> string(11) "29-Jun-2022"
// 	["hour_start"]=> string(5) "22:00"
// 	["duration"]=> string(1) "1"
// 	["title"]=> string(6) "qwewqe"
// 	["game_system"]=> string(5) "wqewq"
// 	["setting"]=> string(5) "wqewq"
// 	["players_min"]=> string(1) "4"
// 	["players_max"]=> string(1) "4"
// 	["language"]=> array(1) {
// 		[0]=> string(2) "es"
// 	}
// 	["content_warning"]=> string(6) "qwewqe"
// 	["description"]=> string(0) ""
// 	["update_settings"]=> string(1) "Y"
// }
function rl_events_create_event_format_post($post) {
  return array(
    'post_title' => rl_events_template($post)->title(),
    'post_content' => rl_events_template($post)->content(),
    'post_status' => 'draft',
    'EventStartDate' => rl_events_create_event_start_date($post),
    'EventEndDate' => rl_events_create_event_end_date($post),
    'EventStartHour' => rl_events_create_event_start_hour($post)['hour'],
    'EventStartMinute' => rl_events_create_event_start_hour($post)['minute'],
    'EventStartMeridian' => rl_events_create_event_start_hour($post)['meridian'],
    'EventEndHour' => rl_events_create_event_end_hour($post)['hour'],
    'EventEndMinute' => rl_events_create_event_end_hour($post)['minute'],
    'EventEndMeridian' => rl_events_create_event_end_hour($post)['meridian'],
    'Organizer' => array(
      'Organizer' => rl_events_create_event_current_user_name(),
      'Email' => rl_events_create_event_current_user_email()
    ),
  );
  return $post;
}

function rl_events_create_event_current_user_name() {
  return wp_get_current_user()->get('display_name') ;
}

function rl_events_create_event_current_user_email() {
  return wp_get_current_user()->get('user_email') ;
}

function rl_events_create_event($data) {
  tribe_create_event($data);
}

