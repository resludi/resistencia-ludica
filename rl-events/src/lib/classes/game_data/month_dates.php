<?php

defined('ABSPATH') or die( "Bye bye" );

class RLEvents_GameDataMonthDates {
  protected $month_year;

  public function __construct($month_year) {
    $this->month_year = $month_year;
  }

  public function firstDay() {
    return date($this->month_year.'-01');
  }

  public function lastDay() {
    return date("Y-m-t", strtotime($this->firstDay()));
  }
}
