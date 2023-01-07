<?php

defined('ABSPATH') or die( "Bye bye" );

class RLEvents_Template {
  protected array $params;

  public function __construct(array $params) {
      $this->params = $params;
  }

  public function title() {
    return $this->params['title']." (".$this->game_system().")";
  }

  public function content() {
    return $this->template();
  }

  private function value($field) {
    return $this->params[$field];
  }

  private function description() {
    return $this->value('description');
  }

  private function game_system() {
    return $this->value('game_system');
  }

  private function setting() {
    return $this->value('setting');
  }

  private function players_min() {
    return $this->value('players_min');
  }

  private function players_max() {
    return $this->value('players_max');
  }

  private function languages() {
    return array_map("strtoupper", $this->value('language'));
  }

  private function content_warning() {
    return $this->value('content_warning');
  }

  private function start_datetime() {
    return rl_events_datetime_start($this->params);
  }

  private function end_datetime() {
    return rl_events_datetime_end($this->params);
  }

  private function html_p($text, $style = '') {
    return "<p style='".$style."'>".$text."</p>";
  }

  private function html_b($text) {
    return "<strong>".$text."</strong>";
  }

  private function intl_day($date) {
    $current = setlocale(LC_TIME, "es_ES.UTF8");
    $date_str = $date->format('l');
    setlocale(LC_TIME, $current);
    return $date_str;
  }

  private function template() {
    $str = $this->html_p($this->description());
    $str .= $this->html_p($this->html_b("Ambientación: ").$this->setting());
    $str .= $this->html_p($this->html_b("Sistema de juego: ").$this->game_system());
    $str .= $this->html_p($this->html_b("Jugadores: ")." mínimo ".$this->players_min().", máximo ".$this->players_max());
    $str .= $this->html_p($this->html_b("Idioma: ").implode(', ', $this->languages()));
    $str .= $this->html_p($this->html_b("Aviso de contenido: ").$this->content_warning());
    $str .= $this->html_p($this->html_b("Día: ").$this->intl_day($this->start_datetime())." ".$this->start_datetime()->format('d/n/Y \d\e H:i A \h')." a ".$this->end_datetime()->format('H:i A \h')." (CEST)");
    $str .= $this->html_p($this->html_b("Te puede interesar: ")."<a href='https://resistencialudica.com/#juega'>Juega con nosotros</a>, <a href='https://resistencialudica.com/#consideraciones'>Consideraciones para eventos online</a>, <a href='https://www.youtube.com/watch?v=DfGxTXRARV8&amp;ab_channel=GORREGR'>Primer encuentro: Discord</a>, <a href='https://resistencialudica.com/2022/01/23/cual-es-mi-discord-id/'>¿Cuál es mi Discord ID?</a>, <a href='https://resistencialudica.com/2022/01/23/como-me-doy-de-baja/'>¿Cómo aviso de que no puedo ir a la partida?</a>");
    $str .= $this->html_p("<img class='wp-image-4063 aligncenter' src='https://resistencialudica.com/wp-content/uploads/2021/10/image20.png' alt='' width='200' height='200' />", 'text-align: center;');
    $str .= $this->html_p("<a href='http://discord.gg/gJAtAwb'><strong>¡Jugamos en Discord!</strong></a></p>", 'text-align: center;');
    return $str;
  }
}
