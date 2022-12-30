<?php

/**
 * The plugin bootstrap file
 *
 * This file is read by WordPress to generate the plugin information in the plugin
 * admin area. This file also includes all of the dependencies used by the plugin,
 * registers the activation and deactivation functions, and defines a function
 * that starts the plugin.
 *
 * @link              https://www.georgenicolaou.me/
 * @since             1.0.0
 * @package           Isop_Summer_Camp_Exporter
 *
 * @wordpress-plugin
 * Plugin Name:       ISOP Summer Camp Exporter
 * Plugin URI:        https://georgenicolaou.me/plugins/isop-summer-school-exporter
 * Description:       This plugin will export all the information regarding the summer camp orders from WooCommerce to an Excel sheet in a human readable format
 * Version:           1.0.0
 * Author:            George Nicolaou
 * Author URI:        https://www.georgenicolaou.me/
 * License:           GPL-2.0+
 * License URI:       http://www.gnu.org/licenses/gpl-2.0.txt
 * Text Domain:       isop-summer-camp-exporter
 * Domain Path:       /languages
 */

// If this file is called directly, abort.
if ( ! defined( 'WPINC' ) ) {
	die;
}

/**
 * Currently plugin version.
 * Start at version 1.0.0 and use SemVer - https://semver.org
 * Rename this for your plugin and update it as you release new versions.
 */
define( 'ISOP_SUMMER_CAMP_EXPORTER_VERSION', '1.0.0' );


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * The code that runs during plugin activation.
 * This action is documented in includes/class-isop-summer-camp-exporter-activator.php
 */
function activate_isop_summer_camp_exporter() {
	require_once plugin_dir_path( __FILE__ ) . 'includes/class-isop-summer-camp-exporter-activator.php';
	Isop_Summer_Camp_Exporter_Activator::activate();
}

/**
 * The code that runs during plugin deactivation.
 * This action is documented in includes/class-isop-summer-camp-exporter-deactivator.php
 */
function deactivate_isop_summer_camp_exporter() {
	require_once plugin_dir_path( __FILE__ ) . 'includes/class-isop-summer-camp-exporter-deactivator.php';
	Isop_Summer_Camp_Exporter_Deactivator::deactivate();
}

register_activation_hook( __FILE__, 'activate_isop_summer_camp_exporter' );
register_deactivation_hook( __FILE__, 'deactivate_isop_summer_camp_exporter' );

/**
 * The core plugin class that is used to define internationalization,
 * admin-specific hooks, and public-facing site hooks.
 */
require plugin_dir_path( __FILE__ ) . 'includes/class-isop-summer-camp-exporter.php';

/**
 * Begins execution of the plugin.
 *
 * Since everything within the plugin is registered via hooks,
 * then kicking off the plugin from this point in the file does
 * not affect the page life cycle.
 *
 * @since    1.0.0
 */
function run_isop_summer_camp_exporter() {

	$plugin = new Isop_Summer_Camp_Exporter();
	$plugin->run();
}
run_isop_summer_camp_exporter();

add_action( 'admin_menu', 'isop_summer_camp_menu' );

function isop_summer_camp_menu() {
  add_menu_page(
    'Isop Summer Camp Exporter Page',
    'Isop Summer Camp Exporter',
    'manage_options',
    'isop-summer-camp-Exporter',
    'isop_summer_camp_callback'
  );
}

function isop_summer_camp_callback() {
	// Check if the user has clicked the export button
	if ( isset( $_POST['export_orders'] ) ) {
	  // Load the WooCommerce plugin functions
	  require_once( ABSPATH . 'wp-admin/includes/plugin.php' );
	  if ( is_plugin_active( 'woocommerce/woocommerce.php' ) ) {
		// Get the orders to export
		$orders = wc_get_orders( array(
		  'status' => array( 'completed', 'processing' ),
		) );
		
		// Load the PhpSpreadsheet library
		require_once( dirname( __FILE__ ) . '/vendor/autoload.php' );
        //echo dirname( __FILE__ ) . '/vendor/autoload.php';

		
		// Create a new Spreadsheet object
		$spreadsheet = new Spreadsheet();
		
		// Set the document properties
		$spreadsheet->getProperties()->setCreator( 'Isop Summer Camp Exporter' )
		  ->setLastModifiedBy( 'Isop Summer Camp Exporter' )
		  ->setTitle( 'Isop Summer Camp Orders' )
		  ->setSubject( 'Isop Summer Camp Orders' )
		  ->setDescription( 'Exported orders from Isop Summer Camp' )
		  ->setKeywords( 'isop summer camp orders' )
		  ->setCategory( 'Orders' );
		
		// Add the orders to the spreadsheet
		$spreadsheet->setActiveSheetIndex( 0 );
		$sheet = $spreadsheet->getActiveSheet();
		$sheet->setTitle( 'Orders' );
		$sheet->setCellValue( 'A1', 'Order ID' );
		$sheet->setCellValue( 'B1', 'Date' );
		$sheet->setCellValue( 'C1', 'Status' );
		$sheet->setCellValue( 'D1', 'Customer' );
		$sheet->setCellValue( 'E1', 'Total' );
		$row = 2;
		foreach ( $orders as $order ) {
		  $sheet->setCellValue( 'A' . $row, $order->get_id() );
		  $sheet->setCellValue( 'B' . $row, $order->get_date_created()->format( 'Y-m-d H:i:s' ) );
		  $sheet->setCellValue( 'C' . $row, $order->get_status() );
		  $customer_name = $order->get_formatted_billing_full_name();
		  if ( ! $customer_name ) {
			$customer_name = 'Guest';
		  }
		  $sheet->setCellValue( 'D' . $row, $customer_name );
		  $sheet->setCellValue( 'E' . $row, $order->get_total() );
		  $row++;
		}
		
		// Set the column widths
		$sheet->getColumnDimension( 'A' )->setWidth( 10 );
		$sheet->getColumnDimension( 'B' )->setWidth( 20 );
		$sheet->getColumnDimension( 'C' )->setWidth( 15 );
		$sheet->getColumnDimension( 'D' )->setWidth( 30 );
		$sheet->getColumnDimension( 'E' )->setWidth( 15 );
		
		// Set the styles for the header row
		$header_style = array(
		  'font' => array(
			'bold' => true,
		  ),
		  'alignment' => array(
			'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
		  ),
		);
		$sheet->getStyle( 'A1:E1' )->applyFromArray( $header_style );
		
		// Set the page setup
		$sheet->getPageSetup()->setOrientation( \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE );
		$sheet->getPageSetup()->setPaperSize( \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4 );
		$sheet->getPageSetup()->setFitToWidth( 1 );
		$sheet->getPageSetup()->setFitToHeight( 0 );
		
		// Set the page margins
		$sheet->getPageMargins()->setTop( 0.75 );
		$sheet->getPageMargins()->setRight( 0.75 );
		$sheet->getPageMargins()->setLeft( 0.75 );
		$sheet->getPageMargins()->setBottom( 0.75 );
		
		// Set the page breaks
		//$sheet->setBreak( 'A2', \PhpOffice\PhpSpreadsheet\Worksheet::BREAK_ROW );
		
		// Redirect output to a client’s web browser (Excel5)
		header( 'Content-Type: application/vnd.ms-excel' );
		header( 'Content-Disposition: attachment;filename="isop-summer-camp-orders.xls"' );
		header( 'Cache-Control: max-age=0' );
		
		$excel_writer = new Xlsx($spreadsheet);
		ob_end_clean();
		$excel_writer->save( 'php://output' );
		exit;
	  }
	}
	
	// Display the export form
	?>
	<form method="post">
	<input type="submit" name="export_orders" value="Export Orders">
  </form>
  <?php
}
	
  
	   
  