<?php

defined('ABSPATH') or die( "Bye bye" );

class RLEvents_GameDataGameOrganizers {
  protected $db;
  protected $postId;

  public function __construct($db, $postId) {
    $this->db = $db;
    $this->postId = $postId;
  }

  public function data() {
    return $this->getResults();
  }

  // Private

  private function getResults() {
    return $this->db->getSqlResults($this->getStatement());
  }

  private function getStatement () {
    return "SELECT pm.meta_value as organizer_id,wp_organizers.post_title organizer
      FROM wp_posts wp_events,
        wp_posts wp_organizers,
        wp_postmeta pm
      WHERE wp_events.ID = pm.post_id
        AND pm.meta_key = '_EventOrganizerID'
        AND wp_events.post_type = 'tribe_events'
        AND wp_organizers.post_type = 'tribe_organizer'
        AND pm.meta_value = wp_organizers.ID
        AND wp_events.ID = $this->postId;
    ;";
  }
}
