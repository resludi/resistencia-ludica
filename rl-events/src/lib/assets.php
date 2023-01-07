<?php

defined('ABSPATH') or die("Bye bye");

add_action('admin_enqueue_scripts', 'rl_events_styles');
add_action('admin_enqueue_scripts', 'rl_events_js');

function rl_events_styles() {
  wp_enqueue_style('rl_events', plugins_url('../../assets/stylesheets/rl_events.css', __FILE__), false, NULL, 'all' );
  wp_enqueue_style('datepicker', plugins_url('../../assets/stylesheets/mc-calendar.min.css', __FILE__), false, NULL, 'all' );
}

function rl_events_js() {
  wp_enqueue_script( 'rl_events', plugins_url('../../assets/javascripts/rl_events.js', __FILE__), false, NULL, false );
  wp_enqueue_script( 'datepicker', plugins_url('../../assets/javascripts/mc-calendar.min.js', __FILE__), false, NULL, false );
  wp_enqueue_script( 'jquery_validate', plugins_url('../../assets/javascripts/jquery.validate.min.js', __FILE__), false, NULL, false );
  wp_enqueue_script( 'jquery_validate_additional', plugins_url('../../assets/javascripts/additional-methods.min.js', __FILE__), false, NULL, false );
}
