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
 * Version:           4.0.5
 * Author:            George Nicolaou
 * Author URI:        https://www.georgenicolaou.me/
 * License:           GPL-2.0+
 * License URI:       http://www.gnu.org/licenses/gpl-2.0.txt
 * Text Domain:       isop-summer-camp-exporter
 * Domain Path:       /languages
 */

// If this file is called directly, abort.
if (!defined('WPINC')) {
    die;
}

/**
 * Currently plugin version.
 * Start at version 1.0.0 and use SemVer - https://semver.org
 * Rename this for your plugin and update it as you release new versions.
 */
define('ISOP_SUMMER_CAMP_EXPORTER_VERSION', '4.0.5');
/*
Constant I need for the custom exporter
*/

/*define('KINDERGARTEN', 'KINDERGARTEN PROGRAMME Ages: 2.5 - 3.5 (Only Non-ISOP. If your child is in the ISOP Kindergarten, please see their teacher)');*/
define('PROGRAMME', 'Select the Programme the child will be attending (Registration fee €20 non-refundable)');
define('ISISOP', 'Was the child a student at The International School of Paphos in 2024-2025 and/or will the child be a student in  The International School of Paphos in 2025-2026?');
define('YEARGROUP', 'Which year group are they in?');
define('WEEKS', 'Please choose the week/s that you would like to register your child for');
define('NAME', 'Name');
define('SURNAME', 'Surname');
define('DOB', 'Date of birth');
define('NATIONALITY', 'Nationality');
define('SPOKEN_LANGS', 'Please list the language/s that your child speaks');
define('ALLERGIES', 'Does your child have any health problems / allergies?');
define('ALLOW_SWIMMING', 'I give permission for my child to take part in swimming');
define('PARENTAL_CONSENT', 'As a parent/guardian of the applicant and with our doctor\'s agreement, I declare that my child is healthy and can take part in the athletic activities of the Summer Camp.');
define('ADD_CHILD', 'Add Another Child');
define('WEEK1', 'Week 1: Monday 23rd June – Friday 27th June');
define('WEEK2', 'Week 2: Monday 30th July - Friday 4th July');
define('WEEK3', 'Week 3: Monday 7th July - Friday 11th July');
define('WEEK4', 'Week 4: Monday 14th July - Friday 18th July');
define('WEEK5', 'Week 5: Monday 21st July - Friday 25th July');
define('WEEK6', 'Week 6: Monday 28th July - Thursday 31st July');
define('ALL_WEEKS', 'All 6 weeks (If you selected this, please do not select the weeks below)');
define('PARENT_NAME', 'Name of Parent / Guardian');
define('PARENT_PHONE', 'Telephone / Contact number');
define('PARENT_EMAIL', 'Parent\'s e-mail address');
define('PARENT_ADDRESS', 'Parent\'s address (and address residing in Paphos if different):');
define('PARENT_SIG', 'E-Signature of parent / guardian:');
define('SET_YES', 'Yes');
define('SET_NO', 'No');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


/** 
 * 
 * Add git autoupdate to help release fixes
*/


require 'plugin-update-checker/plugin-update-checker.php';
use YahnisElsts\PluginUpdateChecker\v5\PucFactory;

$myUpdateChecker = PucFactory::buildUpdateChecker(
	'https://github.com/GeorgeWebDevCy/isop-expoter',
	__FILE__,
	'isop-summer-camp-exporter'
);

//Set the branch that contains the stable release.
$myUpdateChecker->setBranch('main');

//Optional: If you're using a private repository, specify the access token like this:
//$myUpdateChecker->setAuthentication('your-token-here');

/**
 * The code that runs during plugin activation.
 * This action is documented in includes/class-isop-summer-camp-exporter-activator.php
 */
function activate_isop_summer_camp_exporter()
{
    require_once plugin_dir_path(__FILE__) . 'includes/class-isop-summer-camp-exporter-activator.php';
    Isop_Summer_Camp_Exporter_Activator::activate();
}

/**
 * The code that runs during plugin deactivation.
 * This action is documented in includes/class-isop-summer-camp-exporter-deactivator.php
 */
function deactivate_isop_summer_camp_exporter()
{
    require_once plugin_dir_path(__FILE__) . 'includes/class-isop-summer-camp-exporter-deactivator.php';
    Isop_Summer_Camp_Exporter_Deactivator::deactivate();
}

register_activation_hook(__FILE__, 'activate_isop_summer_camp_exporter');
register_deactivation_hook(__FILE__, 'deactivate_isop_summer_camp_exporter');

/**
 * The core plugin class that is used to define internationalization,
 * admin-specific hooks, and public-facing site hooks.
 */
require plugin_dir_path(__FILE__) . 'includes/class-isop-summer-camp-exporter.php';

/**
 * Begins execution of the plugin.
 *
 * Since everything within the plugin is registered via hooks,
 * then kicking off the plugin from this point in the file does
 * not affect the page life cycle.
 *
 * @since    1.0.0
 */
function run_isop_summer_camp_exporter()
{

    $plugin = new Isop_Summer_Camp_Exporter();
    $plugin->run();
}
run_isop_summer_camp_exporter();

add_action('admin_menu', 'isop_summer_camp_menu');

function isop_summer_camp_menu()
{
    add_menu_page(
        'Isop Summer Camp Exporter Page',
        'Isop Summer Camp Exporter',
        'edit_posts',
        'isop-summer-camp-Exporter',
        'isop_summer_camp_callback'
    );
}
function get_epo_data($orderid, $elementid)
{
    $options = THEMECOMPLETE_EPO_API()->get_option($orderid, $elementid);
    foreach ($options as $item_id => $epos) {
        foreach ($epos as $epo) {
            return $epo['value'];
        }
    }
}

function get_epo_checkbox($orderid, $elementid)
{
	$myitems = array();
    $options = THEMECOMPLETE_EPO_API()->get_option($orderid, $elementid);
    foreach ($options as $item_id => $epos) {
        foreach ($epos as $epo) {
             $myitems[] = $epo['value'];
        }
		return $myitems;
    }
}

/*function get_current_child_data($ch_programme, $ch_is_isop, $ch_year_group, $ch_weeks_is_isop, $ch_weeks_non_isop, $ch_name, $ch_surname, $ch_dob, $ch_nationality, $ch_langs_spoken, $ch_health, $ch_swimming, $ch_consent, $ch_add, $ch_parent_name, $ch_parent_phone, $ch_parent_address, $ch_parent_email, $ch_parent_sig, $ch_photo)
{
    $ch_data = array(
        'programme' => $ch_programme,
        'is_isop' => $ch_is_isop,
        'year_group' => $ch_year_group,
        'weeks_non_isop' => array($ch_weeks_non_isop),
        'weeks_is_isop' => array($ch_weeks_is_isop),
        'name' => $ch_name,
        'surname' => $ch_surname,
        'dob' => $ch_dob,
        'nationality' => $ch_nationality,
        'langs_spoken' => $ch_langs_spoken,
        'health' => $ch_health,
        'swimming' => $ch_swimming,
        'consent' => $ch_consent,
        'add' => $ch_add,
        'parent_name' => $ch_parent_name,
        'parent_phone' => $ch_parent_phone,
        'parent_address' => $ch_parent_address,
        'parent_email' => $ch_parent_email,
        'parent_sig' => $ch_parent_sig,
        'photo' => $ch_photo,
    );
    //var_dump($ch_data);
    return $ch_data;
}
*/
function strip_euro_recursive($value) {
    if (is_array($value)) {
        // If value is an array, recursively apply strip_euro_recursive to each element
        return array_map('strip_euro_recursive', $value);
    } else {
        // If value is not an array, strip euro sign and anything after it and then trim whitespace
        return trim(preg_replace('/€.*$/', '', $value));
    }
}

function get_current_child_data($ch_programme, $ch_is_isop, $ch_year_group, $ch_weeks_is_isop, $ch_weeks_non_isop, $ch_name, $ch_surname, $ch_dob, $ch_nationality, $ch_langs_spoken, $ch_health, $ch_swimming, $ch_consent, $ch_add, $ch_parent_name, $ch_parent_phone, $ch_parent_address, $ch_parent_email, $ch_parent_sig, $ch_photo)
{
    // Apply strip_euro_recursive to all values
    $ch_programme = strip_euro_recursive($ch_programme);
    $ch_is_isop = strip_euro_recursive($ch_is_isop);
    $ch_year_group = strip_euro_recursive($ch_year_group);
    $ch_weeks_is_isop = strip_euro_recursive($ch_weeks_is_isop);
    $ch_weeks_non_isop = strip_euro_recursive($ch_weeks_non_isop);
    $ch_name = strip_euro_recursive($ch_name);
    $ch_surname = strip_euro_recursive($ch_surname);
    $ch_dob = strip_euro_recursive($ch_dob);
    $ch_nationality = strip_euro_recursive($ch_nationality);
    $ch_langs_spoken = strip_euro_recursive($ch_langs_spoken);
    $ch_health = strip_euro_recursive($ch_health);
    $ch_swimming = strip_euro_recursive($ch_swimming);
    $ch_consent = strip_euro_recursive($ch_consent);
    $ch_add = strip_euro_recursive($ch_add);
    $ch_parent_name = strip_euro_recursive($ch_parent_name);
    $ch_parent_phone = strip_euro_recursive($ch_parent_phone);
    $ch_parent_address = strip_euro_recursive($ch_parent_address);
    $ch_parent_email = strip_euro_recursive($ch_parent_email);
    $ch_parent_sig = strip_euro_recursive($ch_parent_sig);
    $ch_photo = strip_euro_recursive($ch_photo);

    // Create the data array
    $ch_data = array(
        'programme' => $ch_programme,
        'is_isop' => $ch_is_isop,
        'year_group' => $ch_year_group,
        'weeks_non_isop' => $ch_weeks_non_isop,
        'weeks_is_isop' => $ch_weeks_is_isop,
        'name' => $ch_name,
        'surname' => $ch_surname,
        'dob' => $ch_dob,
        'nationality' => $ch_nationality,
        'langs_spoken' => $ch_langs_spoken,
        'health' => $ch_health,
        'swimming' => $ch_swimming,
        'consent' => $ch_consent,
        'add' => $ch_add,
        'parent_name' => $ch_parent_name,
        'parent_phone' => $ch_parent_phone,
        'parent_address' => $ch_parent_address,
        'parent_email' => $ch_parent_email,
        'parent_sig' => $ch_parent_sig,
        'photo' => $ch_photo,
    );

    // Output the data array
    foreach ($ch_data as $key => $value) {
        //echo "Before stripping: $key<br>";
        //echo "After stripping: ";
        if (is_array($value)) {
            //var_dump($value);
        } else {
            //echo $value . "<br>";
        }
    }

    return $ch_data;
}

function in_array_multi(
    mixed $needle, 
    array $haystack, 
    bool $strict = false) 
{
    foreach ($haystack as $i) {
        if (
            ($strict ? $i === $needle : $i == $needle) ||
            (is_array($i) && in_array_multi($needle, $i, $strict))
        ) {
            return true;
        }
    }
    return false;
}

function insert_child_into_sheet($sheet, $row, $order, $child_data, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig)
{
	
    if ($child_data['programme'] == NULL) {
        return $row; //just send me the same row back since nothing was affected
    } else { //start else
	echo 'order id = '.$order->get_id();
	echo '<pre>';
	var_dump($child_data);
	echo '</pre>';
        $sheet->setCellValue('A' . $row, $order->get_id()); //order id
        $sheet->setCellValue('B' . $row, $order->get_date_created()->format('Y-m-d H:i:s')); //order date
        $sheet->setCellValue('C' . $row, $order->get_status()); //order status
        $sheet->setCellValue('Q' . $row, SET_NO); //week 1
        $sheet->setCellValue('R' . $row, SET_NO); //week 2
        $sheet->setCellValue('S' . $row, SET_NO); //week 3
        $sheet->setCellValue('T' . $row, SET_NO); //week 4 
        $sheet->setCellValue('U' . $row, SET_NO); //week 5
        $sheet->setCellValue('V' . $row, SET_NO); //week 6
        $sheet->setCellValue('G' . $row, SET_NO); //isop student
        $sheet->setCellValue('O' . $row, SET_NO); //swimming
        $sheet->setCellValue('AB' . $row, SET_NO); //photo consent
        $sheet->setCellValue('H' . $row, "N/A"); //yeargroup for kindergarden does apply in 2025

        $customer_name = $order->get_formatted_billing_full_name();
        if (!$customer_name) {
            $customer_name = 'Guest';
        }
        $sheet->setCellValue('D' . $row, $customer_name);
        $sheet->setCellValue('E' . $row, $order->get_total());
        $sheet->setCellValue('F' . $row, $child_data['programme']);

        if (isset($child_data['photo'])) {
            $sheet->setCellValue('AB' . $row, $child_data['photo']);
        }
        /*if ($child_data['programme'] == KINDERGARTEN) {
            //echo "Row in child 1 row = " . $row. " Programme is if" . $child1['programme']; 

            $sheet->setCellValue('Q' . $row, SET_YES);
            $sheet->setCellValue('R' . $row, SET_YES);
            $sheet->setCellValue('S' . $row, SET_YES);
            $sheet->setCellValue('T' . $row, SET_YES);
            $sheet->setCellValue('U' . $row, SET_YES);
            $sheet->setCellValue('O' . $row, SET_NO);

            //kindergarten is non isop 100%
            $sheet->setCellValue('G' . $row, SET_NO);
            $sheet->setCellValue('O' . $row, SET_NO); //no swimming for sure
        }*/

        /*if ($child_data['programme'] != KINDERGARTEN) {
            $sheet->setCellValue('F' . $row, $child_data['programme']);
        }
            */

        if ($child_data['is_isop'] == SET_YES) {
            $sheet->setCellValue('G' . $row, $child_data['is_isop']);
        }



        if (isset($child_data['year_group'])) {
            $sheet->setCellValue('H' . $row, $child_data['year_group']);
        }

        if (isset($child_data['name'])) {
            $sheet->setCellValue('I' . $row, $child_data['name']);
        }

        if (isset($child_data['surname'])) {
            $sheet->setCellValue('J' . $row, $child_data['surname']);
        }

        if (isset($child_data['dob'])) {
            $sheet->setCellValue('K' . $row, $child_data['dob']);
        }
        if (isset($child_data['nationality'])) {
            $sheet->setCellValue('L' . $row, $child_data['nationality']);
        }

        if (isset($child_data['langs_spoken'])) {
            $sheet->setCellValue('M' . $row, $child_data['langs_spoken']);
        }

        if (isset($child_data['health'])) {
            $sheet->setCellValue('N' . $row, $child_data['health']);
        }

        if (isset($child_data['swimming'])) {
            $sheet->setCellValue('O' . $row, $child_data['swimming']);
        }

        if (isset($child_data['consent'])) {
            $sheet->setCellValue('P' . $row, $child_data['consent']);
        }

        if (isset($child_data['photo'])) {
            $sheet->setCellValue('AB' . $row, $child_data['photo']);
        }

        if ($child_data['weeks_non_isop'] != NULL) {
            if (in_array_multi(WEEK1, $child_data['weeks_non_isop'])) {
                $sheet->setCellValue('Q' . $row, SET_YES);
            }

            if (in_array_multi(WEEK2, $child_data['weeks_non_isop'])) {
                $sheet->setCellValue('R' . $row, SET_YES);
            }


            if (in_array_multi(WEEK3, $child_data['weeks_non_isop'])) {
                $sheet->setCellValue('S' . $row, SET_YES);
            }


            if (in_array_multi(WEEK4, $child_data['weeks_non_isop'])) {
                $sheet->setCellValue('T' . $row, SET_YES);
            }


            if (in_array_multi(WEEK5, $child_data['weeks_non_isop'])) {
                $sheet->setCellValue('U' . $row, SET_YES);
            }

            if (in_array_multi(WEEK6, $child_data['weeks_non_isop'])) {
                $sheet->setCellValue('V' . $row, SET_YES);
            }

            

            if (in_array_multi(ALL_WEEKS, $child_data['weeks_non_isop'])) {
                $sheet->setCellValue('Q' . $row, SET_YES);
                $sheet->setCellValue('R' . $row, SET_YES);
                $sheet->setCellValue('S' . $row, SET_YES);
                $sheet->setCellValue('T' . $row, SET_YES);
                $sheet->setCellValue('U' . $row, SET_YES);
                $sheet->setCellValue('V' . $row, SET_YES);

            }
        }

        //is isop
        
        if ($child_data['weeks_is_isop'] != NULL) {
            //debug
            echo 'in if !null';
            echo '<pre>';
            var_dump($child_data['weeks_is_isop']);
            var_dump(in_array_multi(WEEK1, $child_data['weeks_is_isop'])); // true
            echo '</pre>';
            if (in_array_multi(WEEK1, $child_data['weeks_is_isop'])) {
                //echo 'in if week1';
                $sheet->setCellValue('Q' . $row, SET_YES);
            }

            if (in_array_multi(WEEK2, $child_data['weeks_is_isop'])) {
                //echo 'in if week2';
                $sheet->setCellValue('R' . $row, SET_YES);
            }


            if (in_array_multi(WEEK3, $child_data['weeks_is_isop'])) {
                //echo 'in if week3';
                $sheet->setCellValue('S' . $row, SET_YES);
            }


            if (in_array_multi(WEEK4, $child_data['weeks_is_isop'])) {
                //echo 'in if week4';
                $sheet->setCellValue('T' . $row, SET_YES);
            }


            if (in_array_multi(WEEK5, $child_data['weeks_is_isop'])) {
                //echo 'in if week5';
                $sheet->setCellValue('U' . $row, SET_YES);
            }

            if (in_array_multi(WEEK6, $child_data['weeks_is_isop'])) {
                //echo 'in if week6';
                $sheet->setCellValue('V' . $row, SET_YES);
            }


            if (in_array_multi(ALL_WEEKS, $child_data['weeks_is_isop'])) {
                $sheet->setCellValue('Q' . $row, SET_YES);
                $sheet->setCellValue('R' . $row, SET_YES);
                $sheet->setCellValue('S' . $row, SET_YES);
                $sheet->setCellValue('T' . $row, SET_YES);
                $sheet->setCellValue('U' . $row, SET_YES);
                $sheet->setCellValue('V' . $row, SET_YES);

            }
        }

        //parent data set start
        $sheet->setCellValue('W' . $row, $parent_name);
        $sheet->setCellValue('X' . $row, $parent_phone);
        $sheet->setCellValue('Y' . $row, $parent_email);
        $sheet->setCellValue('Z' . $row, $parent_address);
        $sheet->setCellValue('AA' . $row, $parent_sig);





        $row++;

    } //end else

    return $row;
}

function isop_summer_camp_callback()
{


    // Check if the user has clicked the export button

    if (isset($_POST['export_orders']) && isset($_POST['start-year']) && isset($_POST['end-year'])) {

        //print_r($_POST);
        // Load the WooCommerce plugin functions
        require_once(ABSPATH . 'wp-admin/includes/plugin.php');
        if (is_plugin_active('woocommerce/woocommerce.php')) {
            // Get the orders to export
            $startYear = $_POST['start-year'];
            $endYear = $_POST['end-year'];
            // Construct the date range
            $startDate = $startYear . '-01-01'; // Set the start date to the first day of the year
            $endDate = $endYear . '-12-31'; // Set the end date to the last day of the year
            //echo 'Start date' . $startDate;
            //echo 'End date' . $endDate;

            $orders = wc_get_orders(
                array(
                     'status' => array('completed', 'processing', 'on-hold'),
                    'order' => 'ASC',
                    'orderby' => 'ID',
                    'date_created' => $startDate . '...' . $endDate,
					'posts_per_page' => -1

                )
            );


            // Load the PhpSpreadsheet library
            require_once(dirname(__FILE__) . '/vendor/autoload.php');

            // Create a new Spreadsheet object
            $spreadsheet = new Spreadsheet();
            // Set the document properties
            $spreadsheet->getProperties()->setCreator('Isop Summer Camp Exporter')
                ->setLastModifiedBy('Isop Summer Camp Exporter')
                ->setTitle('Isop Summer Camp Orders')
                ->setSubject('Isop Summer Camp Orders')
                ->setDescription('Exported orders from Isop Summer Camp')
                ->setKeywords('isop summer camp orders')
                ->setCategory('Orders');

            // Add the orders to the spreadsheet
            $spreadsheet->setActiveSheetIndex(0);
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->setTitle('Orders');
            $sheet->setCellValue('A1', 'Order ID');
            $sheet->setCellValue('B1', 'Date');
            $sheet->setCellValue('C1', 'Status');
            $sheet->setCellValue('D1', 'Customer');
            $sheet->setCellValue('E1', 'Total');
            $sheet->setCellValue('F1', 'Programme');
            $sheet->setCellValue('G1', 'ISOP Student');
            $sheet->setCellValue('H1', 'Year Group');
            $sheet->setCellValue('I1', 'Name');
            $sheet->setCellValue('J1', 'Surname');
            $sheet->setCellValue('K1', 'DOB');
            $sheet->setCellValue('L1', 'Nationallity');
            $sheet->setCellValue('M1', 'Languages');
            $sheet->setCellValue('N1', 'Allergies');
            $sheet->setCellValue('O1', 'Swimming allowed');
            $sheet->setCellValue('P1', 'Athletic activities allowed');
            $sheet->setCellValue('Q1', 'Week 1');
            $sheet->setCellValue('R1', 'Week 2');
            $sheet->setCellValue('S1', 'Week 3');
            $sheet->setCellValue('T1', 'Week 4');
            $sheet->setCellValue('U1', 'Week 5');
            $sheet->setCellValue('V1', 'Week 6');
            $sheet->setCellValue('W1', 'Parent Name');
            $sheet->setCellValue('X1', 'Parent Phone');
            $sheet->setCellValue('Y1', 'Parent Email');
            $sheet->setCellValue('Z1', 'Parent Address');
            $sheet->setCellValue('AA1', 'Parent Signature');
            $sheet->setCellValue('AB1', 'Photo Consent');

            $row = 2;
            foreach ($orders as $order) {
                unset($ch1_weeks_is_isop);
                unset($ch1_weeks_non_isop);
                unset($ch2_weeks_is_isop);
                unset($ch2_weeks_non_isop);
                unset($ch3_weeks_is_isop);
                unset($ch3_weeks_non_isop);
                unset($ch4_weeks_is_isop);
                unset($ch4_weeks_non_isop);
                unset($ch5_weeks_is_isop);
                unset($ch5_weeks_non_isop);
                unset($ch6_weeks_is_isop);
                unset($ch6_weeks_non_isop);
                $parent_name = get_epo_data($order->get_id(), '63c796ae351489.63307542');
                $parent_phone = get_epo_data($order->get_id(), '63c796ae351491.52493353');
                $parent_email = get_epo_data($order->get_id(), '63c796ae3514a5.64335462');
                $parent_address = get_epo_data($order->get_id(), '63c796ae3514b1.17358693');
                $parent_sig = get_epo_data($order->get_id(), '63c796ae3514c9.80400881');

                $ch1_programme = get_epo_data($order->get_id(), '63c796ae350fd4.47899329');
                $ch1_is_isop = get_epo_data($order->get_id(), '63c796ae351098.29269019');
                $ch1_year_group = get_epo_data($order->get_id(), '63c796ae350fe9.67195496');
                $ch1_weeks_non_isop[] = get_epo_checkbox($order->get_id(), '63c796ae351207.25731871');
                $ch1_weeks_is_isop[] = get_epo_checkbox($order->get_id(), '63c796ae351213.50136665');
                $ch1_name = get_epo_data($order->get_id(), '63c796ae351307.70122204');
                $ch1_surname = get_epo_data($order->get_id(), '63c796ae351316.45003654');
                $ch1_dob = get_epo_data($order->get_id(), '63c796ae3514d2.17093747');
                $ch1_nationality = get_epo_data($order->get_id(), '63c796ae351320.19034332');
                $ch1_langs_spoken = get_epo_data($order->get_id(), '63c796ae351333.14471075');
                $ch1_health = get_epo_data($order->get_id(), '63c796ae351538.87430147');
                $ch1_swimming = get_epo_data($order->get_id(), '63c796ae3510a2.56869018');
                $ch1_consent = get_epo_data($order->get_id(), '63c796ae3510b4.32887095');
                $ch1_add = get_epo_data($order->get_id(), '63c796ae3510c2.48147773');
                $ch1_photo = get_epo_data($order->get_id(), '63cd127a292407.37927849');

                $ch2_programme = get_epo_data($order->get_id(), '63c796ae350ff4.86528396');
                $ch2_is_isop = get_epo_data($order->get_id(), '63c796ae3510d6.49644554');
                $ch2_year_group = get_epo_data($order->get_id(), '63c796ae351002.93863775');
                $ch2_weeks_non_isop = get_epo_checkbox($order->get_id(), '63c796ae351226.05609817');
                $ch2_weeks_is_isop = get_epo_checkbox($order->get_id(), '63c796ae351233.01383284');
                $ch2_name = get_epo_data($order->get_id(), '63c796ae351341.89993507');
                $ch2_surname = get_epo_data($order->get_id(), '63c796ae351356.44394816');
                $ch2_dob = get_epo_data($order->get_id(), '63c796ae3514e6.76479703');
                $ch2_nationality = get_epo_data($order->get_id(), '63c796ae351363.69267898');
                $ch2_langs_spoken = get_epo_data($order->get_id(), '63c796ae351375.65384919');
                $ch2_health = get_epo_data($order->get_id(), '63c796ae351544.70768613');
                $ch2_swimming = get_epo_data($order->get_id(), '63c796ae3510e4.01393417');
                $ch2_consent = get_epo_data($order->get_id(), '63c796ae3510f5.73100731');
                $ch2_add = get_epo_data($order->get_id(), '63c796ae351105.12345352');
                $ch2_photo = get_epo_data($order->get_id(), '63cd1657292438.67347362');

                $ch3_programme = get_epo_data($order->get_id(), '63c796ae351014.63152225');
                $ch3_is_isop = get_epo_data($order->get_id(), '63c796ae351112.16466643');
                $ch3_year_group = get_epo_data($order->get_id(), '63c796ae351025.83062078');
                $ch3_weeks_non_isop = get_epo_checkbox($order->get_id(), '63c796ae351244.09025481');
                $ch3_weeks_is_isop = get_epo_checkbox($order->get_id(), '63c796ae351255.96533613');
                $ch3_name = get_epo_data($order->get_id(), '63c796ae351384.61975114');
                $ch3_surname = get_epo_data($order->get_id(), '63c796ae351394.09162182');
                $ch3_dob = get_epo_data($order->get_id(), '63c796ae3514f3.26171042');
                $ch3_nationality = get_epo_data($order->get_id(), '63c796ae3513a8.83702350');
                $ch3_langs_spoken = get_epo_data($order->get_id(), '63c796ae3513b2.97668616');
                $ch3_health = get_epo_data($order->get_id(), '63c796ae351557.74051345');
                $ch3_swimming = get_epo_data($order->get_id(), '63c796ae351129.76215516');
                $ch3_consent = get_epo_data($order->get_id(), '63c796ae351138.74682524');
                $ch3_add = get_epo_data($order->get_id(), '63c796ae351149.05024047');
                $ch3_photo = get_epo_data($order->get_id(), '63cd1669292448.40299089');

                $ch4_programme = get_epo_data($order->get_id(), '63c796ae351037.35378954');
                $ch4_is_isop = get_epo_data($order->get_id(), '63c796ae351159.16308749');
                $ch4_year_group = get_epo_data($order->get_id(), '63c796ae351042.98702374');
                $ch4_weeks_non_isop = get_epo_checkbox($order->get_id(), '63c796ae351269.16023819');
                $ch4_weeks_is_isop = get_epo_checkbox($order->get_id(), '63c796ae351272.45449040');
                $ch4_name = get_epo_data($order->get_id(), '63c796ae3513c2.35501410');
                $ch4_surname = get_epo_data($order->get_id(), '63c796ae3513d7.04441495');
                $ch4_dob = get_epo_data($order->get_id(), '63c796ae351509.09606580');
                $ch4_nationality = get_epo_data($order->get_id(), '63c796ae3513e6.31276199');
                $ch4_langs_spoken = get_epo_data($order->get_id(), '63c796ae3513f8.02759799');
                $ch4_health = get_epo_data($order->get_id(), '63c796ae351561.12618878');
                $ch4_swimming = get_epo_data($order->get_id(), '63c796ae351162.38108754');
                $ch4_consent = get_epo_data($order->get_id(), '63c796ae351175.59176972');
                $ch4_add = get_epo_data($order->get_id(), '63c796ae351188.63717837');
                $ch4_photo = get_epo_data($order->get_id(), '63cd1681292452.72686890');

                $ch5_programme = get_epo_data($order->get_id(), '63c796ae351056.35349346');
                $ch5_is_isop = get_epo_data($order->get_id(), '63c796ae351191.05095302');
                $ch5_year_group = get_epo_data($order->get_id(), '63c796ae351068.59422766');
                $ch5_weeks_non_isop = get_epo_checkbox($order->get_id(), '63c796ae3512c5.48807658');
                $ch5_weeks_is_isop = get_epo_checkbox($order->get_id(), '63c796ae3512d6.33127856');
                $ch5_name = get_epo_data($order->get_id(), '63c796ae351403.28834192');
                $ch5_surname = get_epo_data($order->get_id(), '63c796ae351418.19452819');
                $ch5_dob = get_epo_data($order->get_id(), '63c796ae351513.50958059');
                $ch5_nationality = get_epo_data($order->get_id(), '63c796ae351427.49487978');
                $ch5_langs_spoken = get_epo_data($order->get_id(), '63c796ae351435.87962631');
                $ch5_health = get_epo_data($order->get_id(), '63c796ae351572.36626393');
                $ch5_swimming = get_epo_data($order->get_id(), '63c796ae3511a1.96267578');
                $ch5_consent = get_epo_data($order->get_id(), '63c796ae3511b3.92669862');
                $ch5_add = get_epo_data($order->get_id(), '63c796ae3511c2.34129641');
                $ch5_photo = get_epo_data($order->get_id(), '63cd1698292467.84694245');

                $ch6_programme = get_epo_data($order->get_id(), '63c796ae351073.56974688');
                $ch6_is_isop = get_epo_data($order->get_id(), '63c796ae3511d9.66370336');
                $ch6_year_group = get_epo_data($order->get_id(), '63c796ae351081.10354514');
                $ch6_weeks_non_isop = get_epo_checkbox($order->get_id(), '63c796ae3512e7.89146672');
                $ch6_weeks_is_isop = get_epo_checkbox($order->get_id(), '63c796ae3512f6.58501165');
                $ch6_name = get_epo_data($order->get_id(), '63c796ae351446.60946329');
                $ch6_surname = get_epo_data($order->get_id(), '63c796ae351452.03336171');
                $ch6_dob = get_epo_data($order->get_id(), '63c796ae351524.28460941');
                $ch6_nationality = get_epo_data($order->get_id(), '63c796ae351467.65291553');
                $ch6_langs_spoken = get_epo_data($order->get_id(), '63c796ae351479.57300118');
                $ch6_health = get_epo_data($order->get_id(), '63c796ae351581.38636241');
                $ch6_swimming = get_epo_data($order->get_id(), '63c796ae3511e3.43617115');
                $ch6_consent = get_epo_data($order->get_id(), '63c796ae3511f6.41303666');
                $ch6_photo = get_epo_data($order->get_id(), '63cd16a9292474.15688992');

                $child1 = get_current_child_data($ch1_programme, $ch1_is_isop, $ch1_year_group, $ch1_weeks_is_isop, $ch1_weeks_non_isop, $ch1_name, $ch1_surname, $ch1_dob, $ch1_nationality, $ch1_langs_spoken, $ch1_health, $ch1_swimming, $ch1_consent, $ch1_add, $parent_name, $parent_phone, $parent_address, $parent_email, $parent_sig, $ch1_photo);
                $child2 = get_current_child_data($ch2_programme, $ch2_is_isop, $ch2_year_group, $ch2_weeks_is_isop, $ch2_weeks_non_isop, $ch2_name, $ch2_surname, $ch2_dob, $ch2_nationality, $ch2_langs_spoken, $ch2_health, $ch2_swimming, $ch2_consent, $ch2_add, $parent_name, $parent_phone, $parent_address, $parent_email, $parent_sig, $ch2_photo);
                $child3 = get_current_child_data($ch3_programme, $ch3_is_isop, $ch3_year_group, $ch3_weeks_is_isop, $ch3_weeks_non_isop, $ch3_name, $ch3_surname, $ch3_dob, $ch3_nationality, $ch3_langs_spoken, $ch3_health, $ch3_swimming, $ch3_consent, $ch3_add, $parent_name, $parent_phone, $parent_address, $parent_email, $parent_sig, $ch3_photo);
                $child4 = get_current_child_data($ch4_programme, $ch4_is_isop, $ch4_year_group, $ch4_weeks_is_isop, $ch4_weeks_non_isop, $ch4_name, $ch4_surname, $ch4_dob, $ch4_nationality, $ch4_langs_spoken, $ch4_health, $ch4_swimming, $ch4_consent, $ch4_add, $parent_name, $parent_phone, $parent_address, $parent_email, $parent_sig, $ch4_photo);
                $child5 = get_current_child_data($ch5_programme, $ch5_is_isop, $ch5_year_group, $ch5_weeks_is_isop, $ch5_weeks_non_isop, $ch5_name, $ch5_surname, $ch5_dob, $ch5_nationality, $ch5_langs_spoken, $ch5_health, $ch5_swimming, $ch5_consent, $ch5_add, $parent_name, $parent_phone, $parent_address, $parent_email, $parent_sig, $ch5_photo);
                $child6 = get_current_child_data($ch6_programme, $ch6_is_isop, $ch6_year_group, $ch6_weeks_is_isop, $ch6_weeks_non_isop, $ch6_name, $ch6_surname, $ch6_dob, $ch6_nationality, $ch6_langs_spoken, $ch6_health, $ch6_swimming, $ch6_consent, $ch6_add, $parent_name, $parent_phone, $parent_address, $parent_email, $parent_sig, $ch6_photo);
                //$isoparray = get_epo_checkbox(5303, '63c796ae351213.50136665');
				//var_dump($isoparray);
			//	echo '<pre>';
                
                //var_dump($ch1_weeks_is_isop);
              //  var_dump($child1);
               // echo '</pre>';
				
                $row = insert_child_into_sheet($sheet, $row, $order, $child1, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig);
                $row = insert_child_into_sheet($sheet, $row, $order, $child2, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig);
                $row = insert_child_into_sheet($sheet, $row, $order, $child3, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig);
                $row = insert_child_into_sheet($sheet, $row, $order, $child4, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig);
                $row = insert_child_into_sheet($sheet, $row, $order, $child5, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig);
                $row = insert_child_into_sheet($sheet, $row, $order, $child6, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig);
//debug

//var_dump($child1);

            }
            // Set the column widths
            $sheet->getColumnDimension('A')->setWidth(10);
            $sheet->getColumnDimension('B')->setWidth(20);
            $sheet->getColumnDimension('C')->setWidth(15);
            $sheet->getColumnDimension('D')->setWidth(30);
            $sheet->getColumnDimension('E')->setWidth(15);
            $sheet->getColumnDimension('F')->setWidth(30);
            $sheet->getColumnDimension('G')->setWidth(80);
            $sheet->getColumnDimension('H')->setWidth(30);
            $sheet->getColumnDimension('I')->setWidth(30);
            $sheet->getColumnDimension('J')->setWidth(30);
            $sheet->getColumnDimension('K')->setWidth(30);
            $sheet->getColumnDimension('L')->setWidth(30);
            $sheet->getColumnDimension('M')->setWidth(30);
            $sheet->getColumnDimension('N')->setWidth(30);
            $sheet->getColumnDimension('O')->setWidth(30);
            $sheet->getColumnDimension('P')->setWidth(30);
            $sheet->getColumnDimension('Q')->setWidth(30);
            $sheet->getColumnDimension('R')->setWidth(30);
            $sheet->getColumnDimension('S')->setWidth(30);
            $sheet->getColumnDimension('T')->setWidth(30);
            $sheet->getColumnDimension('U')->setWidth(30);
            $sheet->getColumnDimension('V')->setWidth(30);
            $sheet->getColumnDimension('W')->setWidth(30);
            $sheet->getColumnDimension('X')->setWidth(30);
            $sheet->getColumnDimension('Y')->setWidth(30);
            $sheet->getColumnDimension('Z')->setWidth(30);
            $sheet->getColumnDimension('AA')->setWidth(30);
            $sheet->getColumnDimension('AB')->setWidth(30);
            // Set the styles for the header row
            $header_style = array(
                'font' => array(
                    'bold' => true,
                ),
                'alignment' => array(
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                ),
            );
            $sheet->getStyle('A1:AA1')->applyFromArray($header_style);

            // Set the page setup
            $sheet->getPageSetup()->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
            $sheet->getPageSetup()->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
            $sheet->getPageSetup()->setFitToWidth(1);
            $sheet->getPageSetup()->setFitToHeight(0);

            // Set the page margins
            $sheet->getPageMargins()->setTop(0.75);
            $sheet->getPageMargins()->setRight(0.75);
            $sheet->getPageMargins()->setLeft(0.75);
            $sheet->getPageMargins()->setBottom(0.75);


            // Redirect output to a client’s web browser (Excel5)
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="isop-summer-camp-orders.xlsx"');
            header('Cache-Control: max-age=0');

            $excel_writer = new Xlsx($spreadsheet);
            ob_end_clean();
            $excel_writer->save('php://output');
            exit;
        }
    }

    // Display the export form
    ?>
    <form method="post" onsubmit="return validateForm()">
        Start Year: <input type="text" id="start-year" name="start-year" pattern="[0-9]{4}"
            value="<?php echo $_POST['start-year']; ?>"><br>
        End Year: <input type="text" id="end-year" name="end-year" pattern="[0-9]{4}"
            value="<?php echo $_POST['end-year']; ?>"><br>
        <input type="submit" name="export_orders" value="Export Orders" id="export-orders-button" disabled>
    </form>


    <script>
        // Get the export orders button
        var exportOrdersButton = document.getElementById('export-orders-button');

        function validateForm() {
            var startYear = document.getElementById('start-year').value;
            var endYear = document.getElementById('end-year').value;

            // Check if the start and end years have been filled out
            if (startYear == '' || endYear == '') {
                exportOrdersButton.disabled = true;  // Disable the button
            } else {
                exportOrdersButton.disabled = false;  // Enable the button
            }
        }

        // Add event listeners to the start-year and end-year inputs
        document.getElementById('start-year').addEventListener('input', validateForm);
        document.getElementById('end-year').addEventListener('input', validateForm);
    </script>

<?php
}