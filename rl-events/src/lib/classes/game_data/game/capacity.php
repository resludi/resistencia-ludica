<?php

defined('ABSPATH') or die( "Bye bye" );

class RLEvents_GameDataGameCapacity {
  protected $db;
  protected $postId;

  public function __construct($db, $postId) {
    $this->db = $db;
    $this->postId = $postId;
  }

  public function data() {
    $results = $this->getResults();
    return isset($results[0]) ? $results[0]->capacity : null;
  }

  // Private

  private function getResults() {
    return $this->db->getSqlResults($this->getStatement());
  }

  private function getStatement () {
    return "SELECT
        pm_capacity.meta_value as capacity
      FROM wp_postmeta pm_event, wp_postmeta pm_capacity
      WHERE pm_event.post_id = pm_capacity.post_id AND
        pm_event.meta_key='_tribe_rsvp_for_event' AND
        pm_capacity.meta_key = '_tribe_ticket_capacity' AND
        pm_event.meta_value = $this->postId
      ;";
  }
}
