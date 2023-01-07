<?php

defined('ABSPATH') or die( "Bye bye" );

if (isset($_POST["update_settings"])) {
	rl_events_create_event(rl_events_create_event_format_post($_POST));
}

//Comprueba que tienes permisos para acceder a esta pagina
if (! rl_events_can_access_new_event_form()) wp_die (__ ('No tienes suficientes permisos para acceder a esta página.'));
?>
	<h2><?php _e( 'Resistencia Lúdica - Events', 'rl_events_long' ) ?></h2>
	<h3>Da de alta la partida que quieres dirigir próximamente:</h3>
	<div class="wrap">
		<form id="create_event" method="post" action="">
			<table>
				<tr>
					<td>
						<label for="date_publish">Fecha de publicación:</label>
						<? echo(rl_events_dates_dropdown('date_publish', rl_events_next_publishing_dates(5))) ?>
					</td>
					<td>
						<label for='date_start'>Día de inicio la partida:</label>
						<input type='text' name='date_start' id='date_start' required>
					</td>
					<td>
						<label for="hour_start">Hora de Inicio:</label>
						<? echo(rl_events_dropdown('hour_start', rl_events_init_hours(8), rl_events_init_hours_default())) ?>
					</td>
					<td>
						<label for="duration">Duración:</label>
						<? echo(rl_events_dropdown('duration', rl_events_duration(8))) ?>
					</td>
				</tr>
				<tr>
					<td colspan='4'>
						<label for="title">Título de la partida:</label>
						<input type="title" id="title" name="title" size="180" required value='Huyendo del Huracan'>
					</td>
				</tr>
				<tr>
					<td colspan='2'>
						<label for="game_system">Sistema de juego:</label>
						<input type="game_system" id="game_system" name="game_system" size="90" required value='Pulp Cthulhu'>
					</td>
					<td colspan='2'>
						<label for="setting">Ambientación:</label>
						<input type="setting" id="setting" name="setting" size="90" required value='Acción, Horror, Suspense'>
					</td>
				</tr>
				<tr>
					<td>
						<label for="players_min">Jugadores Mínimos:</label>
						<? echo(rl_events_dropdown('players_min', rl_events_players_min(8), rl_events_players_min_default())) ?>
					</td>
					<td>
						<label for="players_max">Jugadores Máximos:</label>
						<? echo(rl_events_dropdown('players_max', rl_events_players_max(8), rl_events_players_max_default())) ?>
					</td>
					<td colspan='2'>
						<label for="language">Idioma:</label><br />
						<label class="">Español</label>
						<input value="es" type="checkbox" name="language[]" class="language_group selectit" checked>
						<label class="">Catalán</label>
						<input value="cat" type="checkbox" name="language[]" class="language_group selectit">
					</td>
				</tr>
				<tr>
					<td colspan='4'>
						<label for="content_warning">Aviso de contenido:</label>
						<input type="title" id="content_warning" name="content_warning" size="180" required value='Violencia, horror, investigación, sangre, vísceras, lenguaje mal sonante, mutilación, consumo de drogas'>
					</td>
				</tr>
				<tr>
					<td colspan='4'>
						<label for="content_summary">Sipnosis:</label>
						<?
							echo(rl_events_summary_editor(
								"Los héroes quedan atrapados en una isla a causa de un temporal. ¿Conseguirán liberarla del mal que se cierne sobre ella? Su destino está en manos de los héroes."
							));
						?>
					</td>
				</tr>
			</table>

			<input type="hidden" name="update_settings" value="Y" />
			<input type="submit" value="Crear" class="button-primary action" />
		</form>
	</div>
<?php
?>
