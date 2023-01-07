<?php

defined('ABSPATH') or die("Bye bye");

function rl_events_can_access_game_data() {
  return current_user_can ('manage_options');
}

function rl_events_can_access_event_list() {
  return current_user_can ('manage_options');
}

function rl_events_can_access_new_event_form() {
  return current_user_can ('manage_options');
}
