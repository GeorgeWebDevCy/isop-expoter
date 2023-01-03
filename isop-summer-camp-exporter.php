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
/*
Constant I need for the custom exporter
*/

define ('KINDERGARTEN','KINDERGARTEN PROGRAMME Ages: 2.5 - 3.5 (Only Non-Issp. If you child is in the ISOP Kindergarden contact the school)');
define ('PROGRAMME','Select the Programme the child will be attending (Registration fee €20 non-refundable)');
define ('ISISOP','Is the child a student at The International School of Paphos 2022 - 2023?');
define ('YEARGROUP','Which year group are they in?');
define ('WEEKS','Please choose the week/s that you would like to register your child for');
define ('NAME','Name');
define ('SURNAME','Surname');
define ('DOB','Date of birth');
define ('NATIONALITY','Nationality');
define ('SPOKEN_LANGS','Please list the language/s that your child speaks');
define ('ALLERGIES','Does your child have any health problems / allergies?');
define ('ALLOW_SWIMMING','Allow child to take part in swimming activity');
define ('PARENTAL_CONSENT','As a parent/guardian of the applicant and with our doctor\'s agreement, I declare that my child is healthy and can take part in the athletic activities of the Summer Camp.');
define ('ADD_CHILD','Add Another Child');
define ('WEEK1','Week 1: Monday 26th June - Friday 30th June');
define ('WEEK2','Week 2: Monday 3rd July - Friday 7th July');
define ('WEEK3','Week 3: Monday 10th July - Friday 14th July');
define ('WEEK4','Week 4: Monday 17th July - Friday 21nd July');
define ('WEEK5','Week 5: Monday 24th July - Friday 28th July');
define ('ALL_WEEKS','All 5 weeks (If you selected this, please do not select the weeks below)');
define ('PARENT_NAME','Name of Parent / Guardian');
define ('PARENT_PHONE','Telephone / Contact number');
define ('PARENT_EMAIL','Parent\'s e-mail address');
define ('PARENT_ADDRESS','Parent\'s address (and address residing in Paphos if different):');
define ('PARENT_SIG','E-Signature of parent / guardian:');
define ('SET_YES','Yes');
define ('SET_NO','No');

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

function get_epo_data($orderid,$elementid)
{
	$options = THEMECOMPLETE_EPO_API()->get_option( $orderid ,$elementid );
		  foreach ($options as $item_id => $epos){
				foreach ($epos as $epo){
					  return $epo['value'];
				}
			}
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
			'order' => 'ASC',
			'orderby' => 'ID'
			) );
		
		// Load the PhpSpreadsheet library
		require_once( dirname( __FILE__ ) . '/vendor/autoload.php' );
		
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
		$sheet->setCellValue( 'F1', 'Programme' );
		$sheet->setCellValue( 'G1', 'ISOP Student' );
		$sheet->setCellValue( 'H1', 'Year Group' );
		$sheet->setCellValue( 'I1', 'Name' );
		$sheet->setCellValue( 'J1', 'Surname' );
		$sheet->setCellValue( 'K1', 'DOB' );
		$sheet->setCellValue( 'L1', 'Nationallity' );
		$sheet->setCellValue( 'M1', 'Languages' );
		$sheet->setCellValue( 'N1', 'Allergies' );
		$sheet->setCellValue( 'O1', 'Swimming allowed' );
		$sheet->setCellValue( 'P1', 'Athletic activities allowed' );
		$sheet->setCellValue( 'Q1', 'Week 1' );
		$sheet->setCellValue( 'R1', 'Week 2' );
		$sheet->setCellValue( 'S1', 'Week 3' );
		$sheet->setCellValue( 'T1', 'Week 4' );
		$sheet->setCellValue( 'U1', 'Week 5' );
		$sheet->setCellValue( 'V1', 'Parent Name' );
		$sheet->setCellValue( 'W1', 'Parent Phone' );
		$sheet->setCellValue( 'X1', 'Parent Email' );
		$sheet->setCellValue( 'Y1', 'Parent Address' );
		$sheet->setCellValue( 'Z1', 'Parent Signature' );

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
		  //get EPO data start
		  
		  //$ch1_weeks_non_isop = get_epo_data(5093,'63af5966bf6823.33977599');
		  //$ch1_weeks_is_isop =  get_epo_data(5093,'63af5966bf6833.29671303');
		  //var_dump($ch1_weeks_is_isop);
		  //exit;

		  //child1
		  $ch1_programme = get_epo_data($order->get_id(),'63af5966bf65f2.83538784');
		  $ch1_is_isop = get_epo_data($order->get_id(),'63af5966bf66b8.29518554');
		  $ch1_year_group = get_epo_data($order->get_id(),'63af5966bf6609.77729253');
		  $ch1_weeks_non_isop = get_epo_data($order->get_id(),'63af5966bf6823.33977599');
		  $ch1_weeks_is_isop = get_epo_data($order->get_id(),'63af5966bf6833.29671303');
		  $ch1_name = get_epo_data($order->get_id(),'63af5966bf68e0.45359226');
		  $ch1_surname = get_epo_data($order->get_id(),'63af5966bf68f9.57822234');
		  $ch1_dob = get_epo_data($order->get_id(),'63af5966bf6ab7.76311202');
		  $ch1_nationality = get_epo_data($order->get_id(),'63af5966bf6902.11779081');
		  $ch1_langs_spoken = get_epo_data($order->get_id(),'63af5966bf6914.39580309');
		  $ch1_health = get_epo_data($order->get_id(),'63af5966bf6b17.39160668');
		  $ch1_swimming = get_epo_data($order->get_id(),'63af5966bf66c0.60214178');
		  $ch1_consent = get_epo_data($order->get_id(),'63af5966bf66d7.73256181');
		  $ch1_swimming = get_epo_data($order->get_id(),'63af5966bf66e9.68809535');

		  var_dump($ch1_weeks_is_isop);
		  var_dump($ch1_weeks_non_isop);

		  if($ch1_programme==KINDERGARTEN)
		  {

			$ch1_week1="Yes";
			$ch1_week2="Yes";
			$ch1_week3="Yes";
			$ch1_week4="Yes";
			$ch1_week5="Yes";
			$ch1_year_group="N/A";
			$ch1_is_isop="No";

		  }
		  if($ch1_programme!=KINDERGARTEN && $ch1_weeks_is_isop!=NULL)
		  {
			//if()
		  }
		  $parent_name = get_epo_data($order->get_id(),'63af5966bf6a63.63266197');
		  $parent_phone = get_epo_data($order->get_id(),'63af5966bf6a78.97056490');
		  $parent_email = get_epo_data($order->get_id(),'63af5966bf6a88.53940357');
		  $parent_address = get_epo_data($order->get_id(),'63af5966bf6a93.98073229');
		  $parent_sig = get_epo_data($order->get_id(),'63af5966bf6aa3.21857197');

		  
		  $sheet->setCellValue( 'F' . $row, $ch1_programme );
		  $sheet->setCellValue( 'G' . $row, $ch1_is_isop );
		  $sheet->setCellValue( 'H' . $row, $ch1_year_group );
		  $sheet->setCellValue( 'I' . $row, $ch1_name );
		  $sheet->setCellValue( 'J' . $row, $ch1_surname );
		  $sheet->setCellValue( 'K' . $row, $ch1_dob );
		  $sheet->setCellValue( 'L' . $row, $ch1_nationality );
		  $sheet->setCellValue( 'M' . $row, $ch1_langs_spoken );
		  $sheet->setCellValue( 'N' . $row, $ch1_health );
		  $sheet->setCellValue( 'O' . $row, $ch1_swimming );
		  $sheet->setCellValue( 'P' . $row, $ch1_consent );
		  $sheet->setCellValue( 'Q' . $row, $ch1_week1 );
		  $sheet->setCellValue( 'R' . $row, $ch1_week2 );
		  $sheet->setCellValue( 'S' . $row, $ch1_week3 );
		  $sheet->setCellValue( 'T' . $row, $ch1_week4 );
		  $sheet->setCellValue( 'U' . $row, $ch1_week5 );
		  $sheet->setCellValue( 'V' . $row, $parent_name );
		  $sheet->setCellValue( 'W' . $row, $parent_phone );
		  $sheet->setCellValue( 'X' . $row, $parent_email );
		  $sheet->setCellValue( 'Y' . $row, $parent_address );
		  $sheet->setCellValue( 'Z' . $row, $parent_sig );

		  
		  
		  
		  
		  
		  
		  //child2
		  //child3
		  //child4
		  //child5
		  //child6



		  	$row++;
		}
		
		// Set the column widths
		$sheet->getColumnDimension( 'A' )->setWidth( 10 );
		$sheet->getColumnDimension( 'B' )->setWidth( 20 );
		$sheet->getColumnDimension( 'C' )->setWidth( 15 );
		$sheet->getColumnDimension( 'D' )->setWidth( 30 );
		$sheet->getColumnDimension( 'E' )->setWidth( 15 );
		$sheet->getColumnDimension( 'F' )->setWidth( 30 );
		$sheet->getColumnDimension( 'G' )->setWidth( 80 );
		$sheet->getColumnDimension( 'H' )->setWidth( 30 );
		$sheet->getColumnDimension( 'I' )->setWidth( 30 );
		$sheet->getColumnDimension( 'J' )->setWidth( 30 );
		$sheet->getColumnDimension( 'K' )->setWidth( 30 );
		$sheet->getColumnDimension( 'L' )->setWidth( 30 );
		$sheet->getColumnDimension( 'M' )->setWidth( 30 );
		$sheet->getColumnDimension( 'N' )->setWidth( 30 );
		$sheet->getColumnDimension( 'O' )->setWidth( 30 );
		$sheet->getColumnDimension( 'P' )->setWidth( 30 );
		$sheet->getColumnDimension( 'Q' )->setWidth( 30 );
		$sheet->getColumnDimension( 'R' )->setWidth( 30 );
		$sheet->getColumnDimension( 'S' )->setWidth( 30 );
		$sheet->getColumnDimension( 'T' )->setWidth( 30 );
		$sheet->getColumnDimension( 'U' )->setWidth( 30 );
		$sheet->getColumnDimension( 'V' )->setWidth( 30 );
		$sheet->getColumnDimension( 'W' )->setWidth( 30 );
		$sheet->getColumnDimension( 'X' )->setWidth( 30 );
		$sheet->getColumnDimension( 'Y' )->setWidth( 30 );
		$sheet->getColumnDimension( 'Z' )->setWidth( 30 );		
		// Set the styles for the header row
		$header_style = array(
		  'font' => array(
			'bold' => true,
		  ),
		  'alignment' => array(
			'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
		  ),
		);
		$sheet->getStyle( 'A1:Z1' )->applyFromArray( $header_style );
		
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
		//header( 'Content-Type: application/vnd.ms-excel' );
		//header( 'Content-Disposition: attachment;filename="isop-summer-camp-orders.xls"' );
		//header( 'Cache-Control: max-age=0' );
		
		//$excel_writer = new Xlsx($spreadsheet);
		//ob_end_clean();
		//$excel_writer->save( 'php://output' );
		//exit;
	  }
	}
	
	// Display the export form
	?>
	<form method="post">
	<input type="submit" name="export_orders" value="Export Orders">
  </form>
  <?php
}
	
  
	   
  