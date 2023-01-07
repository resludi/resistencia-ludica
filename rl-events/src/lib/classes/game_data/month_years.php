<?php

defined('ABSPATH') or die( "Bye bye" );

class RLEvents_GameDataMonthYears {
  protected $db;

  public function __construct($db) {
      $this->db = $db;
  }

  public function data() {
    return $this->getUniqueStartDate();
  }

  public function first() {
    return $this->data()[0];
  }

  // Private

  private function getUniqueStartDate() {
    return array_unique($this->getToStartDate());
  }

  private function getToStartDate() {
    return array_map(array($this, 'recordStartDate'), $this->getResults());
  }

  private function getResults() {
    return $this->db->getSqlResults($this->getStatement());
  }

  private function getStatement () {
    return "SELECT
        DATE_FORMAT(pm.meta_value, '%Y-%m') as start_date
      FROM wp_posts wp, wp_postmeta pm
      WHERE wp.ID = pm.post_id
        AND pm.meta_key='_EventStartDate'
        AND post_type = 'tribe_events'
        AND wp.post_status='publish'
      ORDER BY start_date DESC
    ;";
  }

  private function recordStartDate ($record) {
    return $record->start_date;
  }
}
