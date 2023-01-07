<?php

defined('ABSPATH') or die( "Bye bye" );

class RLEvents_GameDataGames {
  protected $db;
  protected $startDate;
  protected $endDate;

  public function __construct($db, $startDate, $endDate) {
    $this->db = $db;
    $this->startDate = $startDate;
    $this->endDate = $endDate;
  }

  public function data() {
    return array_map(array($this, 'gameData'), $this->getResults());
  }

  // Private

  private function getResults() {
    return $this->db->getSqlResults($this->getStatement());
  }

  private function getStatement () {
    return "SELECT
        wp.ID,wp.post_title,pm.meta_value as start_date
      FROM
        wp_posts wp, wp_postmeta pm
      WHERE wp.ID = pm.post_id
        AND pm.meta_key='_EventStartDate'
        AND DATE(pm.meta_value) >= DATE('$this->startDate')
        AND DATE(pm.meta_value) <= DATE('$this->endDate')
        AND post_type = 'tribe_events'
        AND wp.post_status='publish'
      ORDER BY start_date ASC;";
  }

  private function gameData($record) {
    return array(
      'game' => $this->gameGameData($record),
      'attendees' => array(
        'can_go' => $this->gameAttendees($record->ID, 'yes'),
        'cannot_go' => $this->gameAttendees($record->ID, 'no')
      )
    );
  }

  private function gameGameData($record) {
    return array(
      'capacity' => $this->gameCapacity($record->ID),
      'date' => $record->start_date,
      "title" => htmlspecialchars($record->post_title),
      "link" => get_permalink($record->ID),
      "gm" => $this->gameOrganizers($record->ID)
    );
  }

  private function gameCapacity($postId) {
    // TODO - We can cache all game capacities to do only one Query
    return (new RLEvents_GameDataGameCapacity($this->db, $postId))->data();
  }

  private function gameAttendees($postId, $canGo) {
    // TODO - We can cache all game capacities to do only one Query
    return (new RLEvents_GameDataGameAttendees($this->db, $postId, $canGo))->data();
  }

  private function gameOrganizers($postId) {
    // TODO - We can cache all game capacities to do only one Query
    return (new RLEvents_GameDataGameOrganizers($this->db, $postId))->data();
  }
}
