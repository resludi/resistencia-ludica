<?php

defined('ABSPATH') or die( "Bye bye" );

//Comprueba que tienes permisos para acceder a esta pagina
if (!rl_events_can_access_event_list()) wp_die (__ ('No tienes suficientes permisos para acceder a esta página.'));
?>
	<div class="wrap">
		<h2><?php _e( 'Resistencia Lúdica - Lista Eventos Próximo Mes', 'rl_events_next_month_list' ) ?></h2>
		Esta es la lista de eventos para el próximo mes:
	</div>
<?php
 ?>
