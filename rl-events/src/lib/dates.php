<?php

defined('ABSPATH') or die("Bye bye");

function rl_events_following_monday_publishing_datetime($date) {
  return $date->modify('next monday')->setTime(20, 00);
}

function rl_events_next_publishing_dates($size) {
  $dates = array_fill(0, $size, null);
  $date = new DateTimeImmutable('now');

  foreach($dates as $key => $ar_date) {
    $date = rl_events_following_monday_publishing_datetime($date);
    $dates[$key] = $date;
  }
  return $dates;
}

function rl_events_dates_format_value($date) {
  return $date->format('Y/m/d H:i:s');
}

function rl_events_dates_format_view($date) {
  return $date->format('l jS F Y');
}

function rl_events_init_hours() {
  $formatter = function ($time) {
    if ($time % 3600 == 0) {
      return date('H:i', $time);
    } else {
      return date('H:i', $time);
    }
  };
  $halfHourSteps = range(0, 47*1800, 1800);
  return array_map($formatter, $halfHourSteps);
}

function rl_events_init_hours_default() {
  return '22:00';
}

function rl_events_duration($limit) {
  return range(1, $limit, 1);
}

function rl_events_players_min($limit) {
  return range(1, $limit, 1);
}

function rl_events_players_min_default() {
  return 4;
}

function rl_events_players_max($limit) {
  return range(1, $limit, 1);
}

function rl_events_players_max_default() {
  return 4;
}

function rl_events_summary_editor($content = '') {
  $settings = array(
    'textarea_name' => 'description',
    'quicktags' => false,
    'media_buttons' => false,
    'teeny' => true,
    'editor_height' => 100,
    'tinymce' => array(
      'toolbar1' => 'bold,italic,underline,separator,undo,redo',
    )
  );

  return wp_editor($content, 'content_summary', $settings);
}

function rl_events_dropdown($name, $options, $default = null) {
	$html = "<select name='".$name."' autocomplete='off'>";
  foreach($options as $key => $option) {
    if ($option == $default) {
      $html = $html."<option value='".$option."' selected='selected'>".$option."</option>";
    } else {
      $html = $html."<option value='".$option."'>".$option."</option>";
    }
  }
	return $html."</select>";
}

// TODO: DRY this up
// TODO: Localize date strings
function rl_events_dates_dropdown($name, $options, $default = null) {
	$html = "<select name='".$name."' autocomplete='off'>";
  foreach($options as $key => $option) {
    if ($option == $default) {
      $html = $html."<option value='".rl_events_dates_format_value($option)."' selected='selected'>".rl_events_dates_format_view($option)."</option>";
    } else {
      $html = $html."<option value='".rl_events_dates_format_value($option)."'>".rl_events_dates_format_view($option)."</option>";
    }
  }
	return $html."</select>";
}


function rl_events_datetime_start($post) {
  $str_date = $post['date_start']." ".$post['hour_start'];
  return DateTime::createFromFormat('d-M-Y H:i', $str_date);
}

function rl_events_datetime_end($post) {
  return rl_events_datetime_start($post)->add(date_interval_create_from_date_string($post['duration'].' hours'));
}
