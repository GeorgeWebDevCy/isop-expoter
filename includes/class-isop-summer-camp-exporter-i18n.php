<?php

/**
 * Define the internationalization functionality
 *
 * Loads and defines the internationalization files for this plugin
 * so that it is ready for translation.
 *
 * @link       https://www.georgenicolaou.me/
 * @since      1.0.0
 *
 * @package    Isop_Summer_Camp_Exporter
 * @subpackage Isop_Summer_Camp_Exporter/includes
 */

/**
 * Define the internationalization functionality.
 *
 * Loads and defines the internationalization files for this plugin
 * so that it is ready for translation.
 *
 * @since      1.0.0
 * @package    Isop_Summer_Camp_Exporter
 * @subpackage Isop_Summer_Camp_Exporter/includes
 * @author     George Nicolaou <info@georgenicolaou.me>
 */
class Isop_Summer_Camp_Exporter_i18n {


	/**
	 * Load the plugin text domain for translation.
	 *
	 * @since    1.0.0
	 */
	public function load_plugin_textdomain() {

		load_plugin_textdomain(
			'isop-summer-camp-exporter',
			false,
			dirname( dirname( plugin_basename( __FILE__ ) ) ) . '/languages/'
		);

	}



}
