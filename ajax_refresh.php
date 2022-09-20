<?php
require_once '../db.php';
require_once '../views/StatisticData.php';
require_once "../vendor/autoload.php";

if (session_status() == PHP_SESSION_NONE) {
    session_start();
}

if (!(isset ($_SESSION['logged_user']) && ($_SESSION['logged_user']->role == 3 || $_SESSION['logged_user']->role == 2))){
    header('Location: \login.php');
    exit();
}

use MDword\WordProcessor;

$TemplateProcessor = new WordProcessor();
$template = 'docx'.DIRECTORY_SEPARATOR.'template.docx';
$TemplateProcessor->load($template);

$t_one = StatisticData::get_upo_rows_arr();
$TemplateProcessor->clones('t1_row', count($t_one));
$i = 0;
foreach ($t_one as $val){
    $TemplateProcessor->setValue('t1_upo#'.$i, $val['upo']);
    $TemplateProcessor->setValue('t1_level#'.$i, $val['type']);
    $TemplateProcessor->setValue('t1_num#'.$i, $val['count']);
    $i++;
}

$t_two = StatisticData::get_otrasl_rows_arr_asp();
$TemplateProcessor->clones('t2_row', count($t_two));
$i = 0;
foreach ($t_two as $val){
    $TemplateProcessor->setValue('t2_otrasl#'.$i, $val['description']);
    $TemplateProcessor->setValue('t2_num#'.$i, $val['count']);
    $i++;
}

$t_three = StatisticData::get_otrasl_rows_arr_doc();
$TemplateProcessor->clones('t3_row', count($t_three));
$i = 0;
foreach ($t_three as $val){
    $TemplateProcessor->setValue('t3_otrasl#'.$i, $val['description']);
    $TemplateProcessor->setValue('t3_num#'.$i, $val['count']);
    $i++;
}

$t_four = StatisticData::get_spec_rows_arr_asp();
$TemplateProcessor->clones('t4_row', count($t_four));
$i = 0;
foreach ($t_four as $val){
    $TemplateProcessor->setValue('t4_index#'.$i, $i + 1);
    $TemplateProcessor->setValue('t4_code#'.$i, $val['code']);
    $TemplateProcessor->setValue('t4_spec#'.$i, $val['description']);
    $TemplateProcessor->setValue('t4_num#'.$i, $val['count']);
    $i++;
}

$t_five = StatisticData::get_spec_rows_arr_doc();
$TemplateProcessor->clones('t5_row', count($t_five));
$i = 0;
foreach ($t_five as $val){
    $TemplateProcessor->setValue('t5_index#'.$i, $i + 1);
    $TemplateProcessor->setValue('t5_code#'.$i, $val['code']);
    $TemplateProcessor->setValue('t5_spec#'.$i, $val['description']);
    $TemplateProcessor->setValue('t5_num#'.$i, $val['count']);
    $i++;
}

$t_six = StatisticData::get_diss_stats_rows_arr()[0];

$TemplateProcessor->setValue('t6_all_a', $t_six['all_a']);
$TemplateProcessor->setValue('t6_all_d', $t_six['all_d']);
$TemplateProcessor->setValue('t6_n_a', $t_six['n_a']);
$TemplateProcessor->setValue('t6_n_d', $t_six['n_d']);
$TemplateProcessor->setValue('updates', $t_six['updates']);

$t_seven = StatisticData::get_per_month_stats_users_arr();

foreach ($t_seven as $val){
    $TemplateProcessor->setValue('m_'.$val['month'], $val['unique_ip'].'/'.$val['unique_sess']);
}

$rtemplate = 'docx'.DIRECTORY_SEPARATOR.'Количество тем в разрезе УПО.docx';
$TemplateProcessor->saveAs($rtemplate);

echo json_encode(array('success' => 1));