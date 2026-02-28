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
 * Version:           4.0.8
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
define('ISOP_SUMMER_CAMP_EXPORTER_VERSION', '4.0.8');
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
        'manage_woocommerce',
        'isop-summer-camp-Exporter',
        'isop_summer_camp_callback'
    );
}

function get_epo_options($orderid, $elementid)
{
    if (!function_exists('THEMECOMPLETE_EPO_API')) {
        return array();
    }

    $epo_api = THEMECOMPLETE_EPO_API();
    if (method_exists($epo_api, 'get_saved_addons_from_order')) {
        $options = $epo_api->get_saved_addons_from_order($orderid, $elementid);
    } else {
        // Backward compatibility with older versions of the dependency plugin.
        $options = $epo_api->get_option($orderid, $elementid);
    }

    return is_array($options) ? $options : array();
}

function get_epo_data($orderid, $elementid)
{
    $options = get_epo_options($orderid, $elementid);
    foreach ($options as $item_id => $epos) {
        foreach ($epos as $epo) {
            if (isset($epo['value'])) {
                return $epo['value'];
            }
        }
    }

    return null;
}

function get_epo_checkbox($orderid, $elementid)
{
    $myitems = array();
    $options = get_epo_options($orderid, $elementid);
    foreach ($options as $item_id => $epos) {
        foreach ($epos as $epo) {
            if (isset($epo['value'])) {
                $myitems[] = $epo['value'];
            }
        }
    }

    return $myitems;
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

function insert_child_into_sheet($sheet, $row, $order, $child_data, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig,$marketing_source,$marketing_source_other)
{
	
    if ($child_data['programme'] == NULL) {
        return $row; //just send me the same row back since nothing was affected
    } else { //start else
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
        $sheet->setCellValue('AC' . $row, $marketing_source);
        $sheet->setCellValue('AD' . $row, $marketing_source_other);






        $row++;

    } //end else

    return $row;
}

function isop_summer_camp_callback()
{
    $startYearInput = isset($_POST['start-year']) ? sanitize_text_field(wp_unslash($_POST['start-year'])) : '';
    $endYearInput = isset($_POST['end-year']) ? sanitize_text_field(wp_unslash($_POST['end-year'])) : '';


    // Check if the user has clicked the export button

    if (isset($_POST['export_orders']) && isset($_POST['start-year']) && isset($_POST['end-year'])) {
        if (!current_user_can('manage_woocommerce')) {
            wp_die('You do not have permission to export orders.');
        }

        if (
            !isset($_POST['isop_summer_camp_export_nonce']) ||
            !wp_verify_nonce(
                sanitize_text_field(wp_unslash($_POST['isop_summer_camp_export_nonce'])),
                'isop_summer_camp_export'
            )
        ) {
            wp_die('Security check failed.');
        }

        //print_r($_POST);
        // Load the WooCommerce plugin functions
        require_once(ABSPATH . 'wp-admin/includes/plugin.php');
        if (is_plugin_active('woocommerce/woocommerce.php')) {
            if (!function_exists('THEMECOMPLETE_EPO_API')) {
                wp_die('ThemeComplete Extra Product Options API is not available.');
            }

            // Get the orders to export
            if (!preg_match('/^\d{4}$/', $startYearInput) || !preg_match('/^\d{4}$/', $endYearInput)) {
                wp_die('Invalid year format. Please use YYYY.');
            }

            $startYear = (int) $startYearInput;
            $endYear = (int) $endYearInput;
            $maxAllowedYear = (int) gmdate('Y') + 2;

            if ($startYear > $endYear) {
                wp_die('Start year must be less than or equal to end year.');
            }

            if ($startYear < 2000 || $endYear > $maxAllowedYear) {
                wp_die('Year range is out of allowed bounds.');
            }

            // Construct the date range
            $startDate = sprintf('%04d-01-01', $startYear); // Set the start date to the first day of the year
            $endDate = sprintf('%04d-12-31', $endYear); // Set the end date to the last day of the year
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
            $sheet->setCellValue('AC1', 'Marketing Source');
            $sheet->setCellValue('AD1', 'Marketing Source Set to Other');


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
                $parent_name = get_epo_data($order->get_id(), '69a1695a7c14d6.23089076');
                $parent_phone = get_epo_data($order->get_id(), '69a1695a7c14e4.02544636');
                $parent_email = get_epo_data($order->get_id(), '69a1695a7c14f5.02100431');
                $parent_address = get_epo_data($order->get_id(), '69a1695a7c1503.88503410');
                $parent_sig = get_epo_data($order->get_id(), '69a1695a7c1511.62189457');
                $marketing_source = get_epo_data($order->get_id(), '69a1695a7c0f64.05526525');
                $marketing_source_other = get_epo_data($order->get_id(), '69a1695a7c1527.20035702');

                $ch1_programme = get_epo_data($order->get_id(), '69a1695a7c0ea9.02082713');
                $ch1_is_isop = get_epo_data($order->get_id(), '69a1695a7c0f74.09504718');
                $ch1_year_group = get_epo_data($order->get_id(), '69a1695a7c0eb2.21953864');
                $ch1_weeks_non_isop[] = get_epo_checkbox($order->get_id(), '69a1695a7c1140.48717701');
                $ch1_weeks_is_isop[] = get_epo_checkbox($order->get_id(), '69a1695a7c1152.89440043');
                $ch1_name = get_epo_data($order->get_id(), '69a1695a7c1308.02829879');
                $ch1_surname = get_epo_data($order->get_id(), '69a1695a7c1313.60288492');
                $ch1_dob = get_epo_data($order->get_id(), '69a1695a7c1562.12839538');
                $ch1_nationality = get_epo_data($order->get_id(), '69a1695a7c1326.64112164');
                $ch1_langs_spoken = get_epo_data($order->get_id(), '69a1695a7c1332.20935890');
                $ch1_health = get_epo_data($order->get_id(), '69a1695a7c15c2.52819341');
                $ch1_swimming = get_epo_data($order->get_id(), '69a1695a7c0f89.88311589');
                $ch1_consent = get_epo_data($order->get_id(), '69a1695a7c0f98.23082060');
                $ch1_add = get_epo_data($order->get_id(), '69a1695a7c0fb3.44723221');
                $ch1_photo = get_epo_data($order->get_id(), '69a1695a7c0fa6.37872831');

                $ch2_programme = get_epo_data($order->get_id(), '69a1695a7c0ec4.78903872');
                $ch2_is_isop = get_epo_data($order->get_id(), '69a1695a7c0fc5.02663835');
                $ch2_year_group = get_epo_data($order->get_id(), '69a1695a7c0ed1.17217141');
                $ch2_weeks_non_isop = get_epo_checkbox($order->get_id(), '69a1695a7c1172.05618072');
                $ch2_weeks_is_isop = get_epo_checkbox($order->get_id(), '69a1695a7c1183.12477560');
                $ch2_name = get_epo_data($order->get_id(), '69a1695a7c1344.61271865');
                $ch2_surname = get_epo_data($order->get_id(), '69a1695a7c1351.91965440');
                $ch2_dob = get_epo_data($order->get_id(), '69a1695a7c1579.01169817');
                $ch2_nationality = get_epo_data($order->get_id(), '69a1695a7c1369.37220497');
                $ch2_langs_spoken = get_epo_data($order->get_id(), '69a1695a7c1370.75713671');
                $ch2_health = get_epo_data($order->get_id(), '69a1695a7c15d1.39282728');
                $ch2_swimming = get_epo_data($order->get_id(), '69a1695a7c0fd6.82927620');
                $ch2_consent = get_epo_data($order->get_id(), '69a1695a7c0fe0.86844026');
                $ch2_add = get_epo_data($order->get_id(), '69a1695a7c1009.11213078');
                $ch2_photo = get_epo_data($order->get_id(), '69a1695a7c0ff4.85291424');

                $ch3_programme = get_epo_data($order->get_id(), '69a1695a7c0ee4.83165602');
                $ch3_is_isop = get_epo_data($order->get_id(), '69a1695a7c1018.73562668');
                $ch3_year_group = get_epo_data($order->get_id(), '69a1695a7c0ef3.65612751');
                $ch3_weeks_non_isop = get_epo_checkbox($order->get_id(), '69a1695a7c11a9.13800066');
                $ch3_weeks_is_isop = get_epo_checkbox($order->get_id(), '69a1695a7c11b0.81646841');
                $ch3_name = get_epo_data($order->get_id(), '69a1695a7c1385.92716432');
                $ch3_surname = get_epo_data($order->get_id(), '69a1695a7c1399.66233074');
                $ch3_dob = get_epo_data($order->get_id(), '69a1695a7c1583.88760547');
                $ch3_nationality = get_epo_data($order->get_id(), '69a1695a7c13a6.45871572');
                $ch3_langs_spoken = get_epo_data($order->get_id(), '69a1695a7c13b0.07881027');
                $ch3_health = get_epo_data($order->get_id(), '69a1695a7c15e7.91607207');
                $ch3_swimming = get_epo_data($order->get_id(), '69a1695a7c1025.39379772');
                $ch3_consent = get_epo_data($order->get_id(), '69a1695a7c1039.44641730');
                $ch3_add = get_epo_data($order->get_id(), '69a1695a7c1055.75282022');
                $ch3_photo = get_epo_data($order->get_id(), '69a1695a7c1049.19243106');

                $ch4_programme = get_epo_data($order->get_id(), '69a1695a7c0f06.33469928');
                $ch4_is_isop = get_epo_data($order->get_id(), '69a1695a7c1066.17656952');
                $ch4_year_group = get_epo_data($order->get_id(), '69a1695a7c0f17.97629543');
                $ch4_weeks_non_isop = get_epo_checkbox($order->get_id(), '69a1695a7c1242.51721986');
                $ch4_weeks_is_isop = get_epo_checkbox($order->get_id(), '69a1695a7c1250.73158171');
                $ch4_name = get_epo_data($order->get_id(), '69a1695a7c13c6.55032705');
                $ch4_surname = get_epo_data($order->get_id(), '69a1695a7c13d4.53869052');
                $ch4_dob = get_epo_data($order->get_id(), '69a1695a7c1598.89701528');
                $ch4_nationality = get_epo_data($order->get_id(), '69a1695a7c1416.46492015');
                $ch4_langs_spoken = get_epo_data($order->get_id(), '69a1695a7c1426.13625078');
                $ch4_health = get_epo_data($order->get_id(), '69a1695a7c15f6.30984382');
                $ch4_swimming = get_epo_data($order->get_id(), '69a1695a7c1078.87025028');
                $ch4_consent = get_epo_data($order->get_id(), '69a1695a7c1088.25742100');
                $ch4_add = get_epo_data($order->get_id(), '69a1695a7c10a9.74429260');
                $ch4_photo = get_epo_data($order->get_id(), '69a1695a7c1097.60866743');

                $ch5_programme = get_epo_data($order->get_id(), '69a1695a7c0f24.07286904');
                $ch5_is_isop = get_epo_data($order->get_id(), '69a1695a7c10b6.93091538');
                $ch5_year_group = get_epo_data($order->get_id(), '69a1695a7c0f36.81403704');
                $ch5_weeks_non_isop = get_epo_checkbox($order->get_id(), '69a1695a7c1275.63528494');
                $ch5_weeks_is_isop = get_epo_checkbox($order->get_id(), '69a1695a7c1282.56556619');
                $ch5_name = get_epo_data($order->get_id(), '69a1695a7c1434.04451151');
                $ch5_surname = get_epo_data($order->get_id(), '69a1695a7c1441.68392177');
                $ch5_dob = get_epo_data($order->get_id(), '69a1695a7c15a7.11185236');
                $ch5_nationality = get_epo_data($order->get_id(), '69a1695a7c1455.71124343');
                $ch5_langs_spoken = get_epo_data($order->get_id(), '69a1695a7c1466.28004610');
                $ch5_health = get_epo_data($order->get_id(), '69a1695a7c1603.23082201');
                $ch5_swimming = get_epo_data($order->get_id(), '69a1695a7c10c6.92621074');
                $ch5_consent = get_epo_data($order->get_id(), '69a1695a7c10d0.21965373');
                $ch5_add = get_epo_data($order->get_id(), '69a1695a7c10f6.23402097');
                $ch5_photo = get_epo_data($order->get_id(), '69a1695a7c10e0.10325022');

                $ch6_programme = get_epo_data($order->get_id(), '69a1695a7c0f42.44973218');
                $ch6_is_isop = get_epo_data($order->get_id(), '69a1695a7c1106.86287243');
                $ch6_year_group = get_epo_data($order->get_id(), '69a1695a7c0f57.90637022');
                $ch6_weeks_non_isop = get_epo_checkbox($order->get_id(), '69a1695a7c12a8.83399846');
                $ch6_weeks_is_isop = get_epo_checkbox($order->get_id(), '69a1695a7c12e8.78249298');
                $ch6_name = get_epo_data($order->get_id(), '69a1695a7c1478.76545373');
                $ch6_surname = get_epo_data($order->get_id(), '69a1695a7c1489.91474275');
                $ch6_dob = get_epo_data($order->get_id(), '69a1695a7c15b1.51032037');
                $ch6_nationality = get_epo_data($order->get_id(), '69a1695a7c14b4.93312802');
                $ch6_langs_spoken = get_epo_data($order->get_id(), '69a1695a7c14c0.71052850');
                $ch6_health = get_epo_data($order->get_id(), '69a1695a7c1611.87468054');
                $ch6_swimming = get_epo_data($order->get_id(), '69a1695a7c1119.17109144');
                $ch6_consent = get_epo_data($order->get_id(), '69a1695a7c1121.51828903');
                $ch6_add = null;
                $ch6_photo = get_epo_data($order->get_id(), '69a1695a7c1130.56580189');

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
				
                $row = insert_child_into_sheet($sheet, $row, $order, $child1, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig,$marketing_source,$marketing_source_other);
                $row = insert_child_into_sheet($sheet, $row, $order, $child2, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig,$marketing_source,$marketing_source_other);
                $row = insert_child_into_sheet($sheet, $row, $order, $child3, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig,$marketing_source,$marketing_source_other);
                $row = insert_child_into_sheet($sheet, $row, $order, $child4, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig,$marketing_source,$marketing_source_other);
                $row = insert_child_into_sheet($sheet, $row, $order, $child5, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig,$marketing_source,$marketing_source_other);
                $row = insert_child_into_sheet($sheet, $row, $order, $child6, $parent_name, $parent_phone, $parent_email, $parent_address, $parent_sig,$marketing_source,$marketing_source_other);
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
            $sheet->getColumnDimension('AC')->setWidth(90);
            $sheet->getColumnDimension('AD')->setWidth(90);
            // Set the styles for the header row
            $header_style = array(
                'font' => array(
                    'bold' => true,
                ),
                'alignment' => array(
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                ),
            );
            $sheet->getStyle('A1:AD1')->applyFromArray($header_style);

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
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="isop-summer-camp-orders.xlsx"');
            header('Cache-Control: max-age=0');

            $excel_writer = new Xlsx($spreadsheet);
            while (ob_get_level() > 0) {
                ob_end_clean();
            }
            $excel_writer->save('php://output');
            exit;
        }
    }

    // Display the export form
    ?>
    <form method="post" onsubmit="return validateForm()">
        <?php wp_nonce_field('isop_summer_camp_export', 'isop_summer_camp_export_nonce'); ?>
        Start Year: <input type="text" id="start-year" name="start-year" pattern="[0-9]{4}"
            value="<?php echo esc_attr($startYearInput); ?>"><br>
        End Year: <input type="text" id="end-year" name="end-year" pattern="[0-9]{4}"
            value="<?php echo esc_attr($endYearInput); ?>"><br>
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
