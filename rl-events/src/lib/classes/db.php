<?php

defined('ABSPATH') or die( "Bye bye" );

class RLEvents_DB {
  public function getSqlResults ($statement) {
    global $wpdb;
    // echo("<pre>".$statement."</pre>");
    return $wpdb->get_results($statement);
  }
}
