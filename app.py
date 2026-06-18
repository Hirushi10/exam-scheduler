<?php
if (session_status() === PHP_SESSION_NONE) {
    session_start();
}

require_once 'config.php';
require_once 'SimpleXLSX.php';
use Shuchkin\SimpleXLSX;

$all_staff_names = []; 
$all_halls = []; 
$all_subjects = []; 
$hall_capacities = [];

// 📊 1. PARSE EXCEL RECORDS
if ($xlsx = SimpleXLSX::parse('staff_list.xlsx')) {
    $rows = $xlsx->rows();
    for ($i = 1; $i < count($rows); $i++) { 
        if (!empty($rows[$i][0])) $all_staff_names[] = trim($rows[$i][0]); 
    }
}
if ($xlsx = SimpleXLSX::parse('halls_list.xlsx')) {
    $rows = $xlsx->rows();
    for ($i = 1; $i < count($rows); $i++) { 
        if (!empty($rows[$i][0])) {
            $h_name = trim($rows[$i][0]);
            $all_halls[] = $h_name; 
            $hall_capacities[$h_name] = !empty($rows[$i][1]) ? (int)$rows[$i][1] : 'N/A';
        }
    }
}
if ($xlsx = SimpleXLSX::parse('subjects_list.xlsx')) {
    $rows = $xlsx->rows();
    for ($i = 1; $i < count($rows); $i++) { 
        if (!empty($rows[$i][0])) $all_subjects[] = trim($rows[$i][0]); 
    }
}

// 🗑️ 2. DELETE ENGINE ROUTE
if (isset($_GET['action']) && $_GET['action'] === 'delete' && isset($_GET['id'])) {
    try {
        $del_stmt = $pdo->prepare("DELETE FROM exam_assignments WHERE id = :id");
        $del_stmt->execute([':id' => (int)$_GET['id']]);
        header("Location: index.php?status=deleted"); 
        exit();
    } catch (PDOException $e) { 
        die("Error deleting entry: " . $e->getMessage()); 
    }
}

// ✏️ 3. EDIT ENGINE MODE ROUTE
$edit_mode = false; 
$edit_id = null; 
$edit_data = [];
if (isset($_GET['action']) && $_GET['action'] === 'edit' && isset($_GET['id'])) {
    $edit_id = (int)$_GET['id'];
    $edit_stmt = $pdo->prepare("SELECT * FROM exam_assignments WHERE id = :id");
    $edit_stmt->execute([':id' => $edit_id]);
    $edit_data = $edit_stmt->fetch();
    if ($edit_data) { 
        $edit_mode = true; 
    }
}

// 📑 4. FETCH CURRENT ENTRIES
try {
    $stmt = $pdo->query("SELECT * FROM exam_assignments ORDER BY id DESC");
    $schedules = $stmt->fetchAll();
} catch (PDOException $e) { 
    die("Database Query Failure: " . $e->getMessage()); 
}

// 📊 5. CALCULATE WORKLOAD SUMMARIES
$duty_counts = array_fill_keys($all_staff_names, 0);
foreach ($schedules as $row) {
    $sup = $row['supervisor'];
    if (isset($duty_counts[$sup])) $duty_counts[$sup]++;
    if (!empty($row['invigilators'])) {
        $invs = array_map('trim', explode(',', $row['invigilators']));
        foreach ($invs as $inv) { 
            if (isset($duty_counts[$inv])) $duty_counts[$inv]++; 
        }
    }
}
arsort($duty_counts);

// 🕒 6. STATE HANDLING MATRICES
if ($edit_mode) {
    $current_date = $edit_data['date'];
    $parts = explode(' - ', $edit_data['time']);
    $current_start = isset($parts[0]) ? date('H:i', strtotime($parts[0])) : '09:00';
    $current_end = isset($parts[1]) ? date('H:i', strtotime($parts[1])) : '12:00';
} else {
    $current_date = $_SESSION['last_date'] ?? date('Y-m-d');
    $current_start = $_SESSION['last_start'] ?? '09:00';
    $current_end = $_SESSION['last_end'] ?? '12:00';
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exams Scheduler</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #F8FAFC; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .main-title { font-size:26px; font-weight:bold; color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 5px; margin-bottom: 15px; }
        .hall-box { border: 1px solid #CBD5E1; padding: 15px; background-color: #FFFFFF; border-radius: 5px; margin-bottom: 15px; min-height: 240px; box-shadow: 1px 1px 5px rgba(0,0,0,0.05); transition: all 0.2s; }
        .hall-title { font-size: 16px; font-weight: bold; color: #1E40AF; border-bottom: 1px solid #E2E8F0; padding-bottom: 3px; margin-bottom: 8px
