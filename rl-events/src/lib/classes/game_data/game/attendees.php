<?php

defined('ABSPATH') or die( "Bye bye" );

class RLEvents_GameDataGameAttendees {
  protected $db;
  protected $postId;
  protected $canGo; // 'yes' or 'no'

  public function __construct($db, $postId, $canGo) {
    $this->db = $db;
    $this->postId = $postId;
    $this->canGo = $canGo;
  }

  public function data() {
    return $this->filterAttendeesCanGo();
  }

  // Private

  private function getResults() {
    return $this->db->getSqlResults($this->getStatement());
  }

  private function getStatement () {
    return "SELECT pm_status.meta_value as can_go,
        pm_email.meta_value as email,
        pm_name.meta_value as name,
        wp.post_date as date from wp_posts wp,
        wp_postmeta pm_name,
        wp_postmeta pm_email,
        wp_postmeta pm_eventid,
        wp_postmeta pm_status
      WHERE wp.ID = pm_email.post_id
        AND wp.ID = pm_name.post_id
        AND wp.ID = pm_eventid.post_id
        AND wp.ID = pm_status.post_id
        AND wp.post_type = 'tribe_rsvp_attendees'
        AND pm_email.meta_key = '_tribe_rsvp_email'
        AND pm_name.meta_key = '_tribe_rsvp_full_name'
        AND pm_eventid.meta_key = '_tribe_rsvp_event'
        AND pm_status.meta_key = '_tribe_rsvp_status'
        AND pm_eventid.meta_value = $this->postId";
  }

  function attendeeCanGo($attendee) {
    return $attendee->can_go === $this->canGo;
  }

  function filterAttendeesCanGo() {
    return array_filter($this->getResults(), array($this, "attendeeCanGo"));
  }
}
