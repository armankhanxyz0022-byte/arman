<?php
// meesho.php - Premium Meesho Business Report
// Sales → RTO → Net Sales → Ads → Final Cash → Return → Fees → PROFIT/LOSS

error_reporting(E_ALL);
ini_set('display_errors', 0);
ini_set('memory_limit', '512M');
set_time_limit(300);

$uploadDir  = __DIR__ . '/uploads/';
$extractDir = __DIR__ . '/uploads/extracted/';
if (!is_dir($uploadDir))  mkdir($uploadDir, 0777, true);
if (!is_dir($extractDir)) mkdir($extractDir, 0777, true);

$report = null;
$error  = '';

// ================================================================
//  NATIVE XLSX READER
// ================================================================
class NativeXlsxReader
{
    private $zip;
    private $sharedStrings = [];
    private $sheetMap = [];

    public function load($path)
    {
        if (!class_exists('ZipArchive')) throw new Exception('PHP Zip extension not enabled.');
        $this->zip = new ZipArchive();
        if ($this->zip->open($path) !== true) throw new Exception('Cannot open XLSX file.');
        $this->sharedStrings = [];
        $ssXml = $this->zip->getFromName('xl/sharedStrings.xml');
        if ($ssXml !== false) $this->parseSharedStrings($ssXml);
        $this->sheetMap = $this->buildSheetMap();
    }

    private function parseSharedStrings($xml)
    {
        $reader = new XMLReader();
        $reader->XML($xml);
        $currentString = ''; $inSi = false;
        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::ELEMENT) {
                if ($reader->localName === 'si') { $inSi = true; $currentString = ''; }
                if ($reader->localName === 't' && $inSi) $currentString .= $reader->readString();
            }
            if ($reader->nodeType === XMLReader::END_ELEMENT) {
                if ($reader->localName === 'si') {
                    $this->sharedStrings[] = $currentString;
                    $inSi = false;
                }
            }
        }
        $reader->close();
    }

    private function buildSheetMap()
    {
        $map = [];
        $wbXml = $this->zip->getFromName('xl/workbook.xml');
        if ($wbXml === false) return $map;

        $relsXml = $this->zip->getFromName('xl/_rels/workbook.xml.rels');
        $rels = [];
        if ($relsXml !== false) {
            if (preg_match_all('/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/i', $relsXml, $m)) {
                foreach ($m[1] as $i => $rid) {
                    $target = ltrim($m[2][$i], '/');
                    if (strpos($target, 'xl/') !== 0) $target = 'xl/' . $target;
                    $rels[$rid] = $target;
                }
            }
        }
        if (preg_match_all('/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/i', $wbXml, $sm)) {
            foreach ($sm[1] as $i => $name) {
                $rid = $sm[2][$i];
                if (isset($rels[$rid])) $map[$name] = $rels[$rid];
            }
        }
        return $map;
    }

    public function getSheetNames() { return array_keys($this->sheetMap); }

    public function getAllSheets()
    {
        $out = [];
        foreach ($this->sheetMap as $name => $path) $out[$name] = $this->parseSheet($path);
        return $out;
    }

    private function parseSheet($sheetPath)
    {
        $xml = $this->zip->getFromName($sheetPath);
        if ($xml === false) return [];

        $rows = [];
        $reader = new XMLReader();
        $reader->XML($xml);

        $currentRow = []; $colIndex = -1; $inCell = false;
        $cellType = ''; $cellValue = '';

        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::ELEMENT) {
                if ($reader->localName === 'row') {
                    $currentRow = [];
                } elseif ($reader->localName === 'c') {
                    $inCell = true;
                    $cellType = $reader->getAttribute('t') ?? '';
                    $ref = $reader->getAttribute('r') ?? '';
                    $colIndex = $this->colLetterToIndex($ref);
                    $cellValue = '';
                } elseif ($inCell && $reader->localName === 'v') {
                    $cellValue = $reader->readString();
                } elseif ($inCell && $reader->localName === 't') {
                    $cellValue = $reader->readString();
                }
            } elseif ($reader->nodeType === XMLReader::END_ELEMENT) {
                if ($reader->localName === 'c' && $inCell) {
                    $value = ($cellType === 's') ? ($this->sharedStrings[(int)$cellValue] ?? '') : $cellValue;
                    if ($colIndex >= 0) $currentRow[$colIndex] = $value;
                    $inCell = false; $colIndex = -1;
                } elseif ($reader->localName === 'row') {
                    if (!empty($currentRow)) {
                        $maxCol = max(array_keys($currentRow));
                        $norm = [];
                        for ($c = 0; $c <= $maxCol; $c++) $norm[$c] = $currentRow[$c] ?? '';
                        $rows[] = $norm;
                    }
                }
            }
        }
        $reader->close();
        return $rows;
    }

    private function colLetterToIndex($ref)
    {
        preg_match('/^([A-Z]+)/', $ref, $m);
        if (empty($m[1])) return -1;
        $letters = $m[1]; $num = 0;
        for ($i = 0; $i < strlen($letters); $i++) $num = $num * 26 + (ord($letters[$i]) - ord('A') + 1);
        return $num - 1;
    }

    public function close() { if ($this->zip) $this->zip->close(); }
}

// ================================================================
//  HELPERS
// ================================================================
function cleanNumber($val) {
    if (is_numeric($val)) return (float)$val;
    $val = str_replace([',', '₹', '$', ' ', '"', "\xC2\xA0"], '', (string)$val);
    return is_numeric($val) ? (float)$val : 0;
}

function detectStatus($status) {
    $s = strtolower(trim($status));
    if ($s === '') return 'unknown';
    if (strpos($s, 'rto') !== false) return 'rto';
    if (strpos($s, 'return') !== false || strpos($s, 'refund') !== false) return 'return';
    if (strpos($s, 'cancel') !== false) return 'cancelled';
    if (strpos($s, 'delivered') !== false) return 'delivered';
    if (strpos($s, 'shipped') !== false || strpos($s, 'dispatch') !== false) return 'shipped';
    if (strpos($s, 'pending') !== false || strpos($s, 'processing') !== false) return 'pending';
    return 'other';
}

function findMeeshoHeaderRow($rows) {
    foreach ($rows as $i => $row) {
        foreach ($row as $cell) {
            $c = strtolower(trim((string)$cell));
            if ($c === 'sub order no' || $c === 'sub order') return $i;
        }
    }
    foreach ($rows as $i => $row) {
        if (count(array_filter($row, fn($v) => $v !== null && $v !== '')) > 5) return $i;
    }
    return 0;
}

function meeshoCols() {
    return [
        'sub_order_no'=>0,'order_date'=>1,'dispatch_date'=>2,'product'=>3,
        'sku'=>4,'catalog_id'=>5,'order_source'=>6,'status'=>7,
        'product_gst'=>8,'listing_price'=>9,'qty'=>10,'transaction_id'=>11,
        'payment_date'=>12,'final_settlement'=>13,'price_type'=>14,
        'total_sale_amount'=>15,'total_sale_return'=>16,'fixed_fee_1'=>17,
        'warehousing_fee_1'=>18,'return_premium'=>19,'return_premium_return'=>20,
        'commission_pct'=>21,'meesho_commission'=>22,'meesho_gold_fee'=>23,
        'meesho_mall_fee'=>24,'fixed_fee_2'=>25,'warehousing_fee_2'=>26,
        'return_shipping'=>27,'gst_compensation'=>28,'shipping_charge'=>29,
        'other_support'=>30,'waivers'=>31,'net_other_support'=>32,
        'gst_on_other'=>33,'tcs'=>34,'tds_rate'=>35,'tds'=>36,
        'compensation'=>37,'claims'=>38,'recovery'=>39,
    ];
}

// ================================================================
//  PROCESS MEESHO FILE
// ================================================================
function processMeeshoFile($path)
{
    $reader = new NativeXlsxReader();
    $reader->load($path);
    $sheets = $reader->getAllSheets();
    $reader->close();

    if (empty($sheets)) throw new Exception('No sheets found in file.');

    $r = [
        'filename'    => basename($path),
        'sheet_names' => array_keys($sheets),

        'total_rows' => 0, 'total_orders' => 0,
        'sold' => 0, 'sold_qty' => 0,
        'delivered' => 0, 'delivered_qty' => 0,
        'shipped' => 0, 'shipped_qty' => 0,
        'pending' => 0, 'pending_qty' => 0,
        'return' => 0, 'return_qty' => 0,
        'rto' => 0, 'rto_qty' => 0,
        'cancelled' => 0, 'cancelled_qty' => 0,
        'other' => 0,

        'listing_price'         => 0,
        'total_sale_amount'     => 0,
        'total_sale_return'     => 0,

        'fixed_fee'             => 0,
        'warehousing_fee'       => 0,
        'return_premium'        => 0,
        'meesho_commission'     => 0,
        'meesho_gold_fee'       => 0,
        'meesho_mall_fee'       => 0,
        'return_shipping'       => 0,
        'gst_compensation'      => 0,
        'shipping_charge'       => 0,
        'other_support'         => 0,
        'waivers'               => 0,
        'net_other_support'     => 0,
        'gst_on_other'          => 0,

        'tcs' => 0, 'tds' => 0,

        'compensation' => 0, 'claims' => 0, 'recovery' => 0,

        'final_settlement'      => 0,

        'ads_cost' => 0, 'ads_cost_incl' => 0, 'ads_gst' => 0, 'ads_total' => 0, 'ads_rows' => 0,

        'referral_net' => 0, 'referral_tax' => 0, 'referral_rows' => 0,

        'platform_compensation' => 0, 'platform_comp_rows' => 0,

        'profit_calc' => [
            'delivered_orders'      => 0,
            'delivered_qty'         => 0,
            'my_selling_total'      => 0,
            'meesho_selling_total'  => 0,
            'gross_profit_margin'   => 0,
            'meesho_commission_cut' => 0,
            'meesho_other_fee_cut'  => 0,
            'meesho_total_earned'   => 0,
            'shipping_charged'      => 0,
            'tax_deducted'          => 0,
            'net_you_earned'        => 0,
            'return_orders'         => 0,
            'return_qty'            => 0,
            'return_loss'           => 0,
            'return_shipping_cut'   => 0,
            'rto_orders'            => 0,
            'rto_qty'               => 0,
            'rto_loss'              => 0,
            'rto_shipping_cut'      => 0,
            'ads_cost'              => 0,
            'ads_gst'               => 0,
            'total_ads'             => 0,
            'final_in_hand'         => 0,
            'net_profit_loss'       => 0,
            'per_unit_profit'       => 0,
        ],

        // ⭐ REVENUE BREAKDOWN (Step by Step) ⭐
        'revenue_breakdown' => [
            'gross_sales_pieces'      => 0,
            'gross_sales_value'       => 0,
            'rto_pieces'              => 0,
            'rto_value'               => 0,
            'customer_return_pieces'  => 0,
            'customer_return_value'   => 0,
            'net_sold_pieces'         => 0,
            'net_product_revenue'     => 0,
            'returned_stock_pieces'   => 0,
            'returned_stock_value'    => 0,
            'rto_courier_charges'     => 0,
            'customer_return_shipping'=> 0,
            'platform_fees_total'     => 0,
            'final_cash_in_hand'      => 0,
            'cost_price_total'        => 0,
            'total_profit_loss'       => 0,
        ],

        'status_breakdown' => [],
        'top_products' => [],
    ];

    // ========== 1. ORDER PAYMENTS ==========
    $orderSheet = null;
    foreach ($sheets as $name => $rows) {
        if (stripos($name, 'order') !== false && stripos($name, 'payment') !== false) {
            $orderSheet = $rows; break;
        }
    }

    if ($orderSheet) {
        $headerIdx = findMeeshoHeaderRow($orderSheet);
        $dataRows  = array_slice($orderSheet, $headerIdx + 1);

        if (!empty($dataRows)) {
            $first = $dataRows[0];
            $isFormula = false;
            foreach ($first as $cell) {
                if (preg_match('/^[A-Z]\s*[=+\-]/', (string)$cell)) { $isFormula = true; break; }
            }
            if ($isFormula) array_shift($dataRows);
        }

        $C = meeshoCols();
        $uniqueOrders = [];

        foreach ($dataRows as $row) {
            $nonEmpty = array_filter($row, fn($v) => $v !== null && $v !== '');
            if (count($nonEmpty) < 3) continue;

            $orderNo = trim((string)($row[$C['sub_order_no']] ?? ''));
            if ($orderNo === '') continue;

            $r['total_rows']++;
            $uniqueOrders[$orderNo] = true;

            $statusRaw  = trim((string)($row[$C['status']] ?? ''));
            $statusType = detectStatus($statusRaw);

            $product = trim((string)($row[$C['product']] ?? '')) ?: 'Unknown';
            $qty     = max(1, (int)($row[$C['qty']] ?? 1));

            $myPrice      = cleanNumber($row[$C['listing_price']] ?? 0);
            $meeshoPrice  = cleanNumber($row[$C['total_sale_amount']] ?? 0);
            $commission   = cleanNumber($row[$C['meesho_commission']] ?? 0);
            $goldFee      = cleanNumber($row[$C['meesho_gold_fee']] ?? 0);
            $mallFee      = cleanNumber($row[$C['meesho_mall_fee']] ?? 0);
            $fixedFee     = cleanNumber($row[$C['fixed_fee_1']] ?? 0) + cleanNumber($row[$C['fixed_fee_2']] ?? 0);
            $warehouse    = cleanNumber($row[$C['warehousing_fee_1']] ?? 0) + cleanNumber($row[$C['warehousing_fee_2']] ?? 0);
            $shipping     = cleanNumber($row[$C['shipping_charge']] ?? 0);
            $retShipping  = cleanNumber($row[$C['return_shipping']] ?? 0);
            $retPremium   = cleanNumber($row[$C['return_premium']] ?? 0) + cleanNumber($row[$C['return_premium_return']] ?? 0);
            $otherSupport = cleanNumber($row[$C['other_support']] ?? 0);
            $netOther     = cleanNumber($row[$C['net_other_support']] ?? 0);
            $gstOnOther   = cleanNumber($row[$C['gst_on_other']] ?? 0);
            $tcs          = cleanNumber($row[$C['tcs']] ?? 0);
            $tds          = cleanNumber($row[$C['tds']] ?? 0);
            $finalSettle  = cleanNumber($row[$C['final_settlement']] ?? 0);
            $saleReturn   = cleanNumber($row[$C['total_sale_return']] ?? 0);

            // Status counters
            switch ($statusType) {
                case 'delivered': $r['delivered']++; $r['delivered_qty'] += $qty; $r['sold']++; $r['sold_qty'] += $qty; break;
                case 'shipped':   $r['shipped']++;   $r['shipped_qty'] += $qty;   $r['sold']++; $r['sold_qty'] += $qty; break;
                case 'pending':   $r['pending']++;   $r['pending_qty'] += $qty;   break;
                case 'cancelled': $r['cancelled']++; $r['cancelled_qty'] += $qty; break;
                case 'return':    $r['return']++;    $r['return_qty'] += $qty;    break;
                case 'rto':       $r['rto']++;       $r['rto_qty'] += $qty;       break;
                default:          $r['other']++;
            }

            // File-wide accumulators
            $r['listing_price']     += $myPrice;
            $r['total_sale_amount'] += $meeshoPrice;
            $r['total_sale_return'] += $saleReturn;
            $r['fixed_fee']         += $fixedFee;
            $r['warehousing_fee']   += $warehouse;
            $r['return_premium']    += $retPremium;
            $r['meesho_commission'] += $commission;
            $r['meesho_gold_fee']   += $goldFee;
            $r['meesho_mall_fee']   += $mallFee;
            $r['return_shipping']   += $retShipping;
            $r['gst_compensation']  += cleanNumber($row[$C['gst_compensation']] ?? 0);
            $r['shipping_charge']   += $shipping;
            $r['other_support']     += $otherSupport;
            $r['net_other_support'] += $netOther;
            $r['gst_on_other']      += $gstOnOther;
            $r['tcs']               += $tcs;
            $r['tds']               += $tds;
            $r['compensation']      += cleanNumber($row[$C['compensation']] ?? 0);
            $r['claims']            += cleanNumber($row[$C['claims']] ?? 0);
            $r['recovery']          += cleanNumber($row[$C['recovery']] ?? 0);
            $r['final_settlement']  += $finalSettle;

            // ⭐ REVENUE BREAKDOWN accumulators ⭐
            $rb = &$r['revenue_breakdown'];
            $rb['gross_sales_pieces'] += $qty;
            $rb['gross_sales_value']  += $myPrice * $qty;

            if ($statusType === 'rto') {
                $rb['rto_pieces'] += $qty;
                $rb['rto_value']  += $myPrice * $qty;
                $rb['rto_courier_charges'] += abs($retShipping);
                $rb['returned_stock_pieces'] += $qty;
                $rb['returned_stock_value']  += $myPrice * $qty;
            } elseif ($statusType === 'return') {
                $rb['customer_return_pieces'] += $qty;
                $rb['customer_return_value']  += $myPrice * $qty;
                $rb['customer_return_shipping'] += abs($retShipping);
                $rb['returned_stock_pieces'] += $qty;
                $rb['returned_stock_value']  += $myPrice * $qty;
            } elseif ($statusType === 'delivered') {
                $rb['net_sold_pieces'] += $qty;
                $rb['net_product_revenue'] += $myPrice * $qty;
                $rb['platform_fees_total'] += abs($commission) + abs($goldFee) + abs($mallFee)
                                           + abs($fixedFee) + abs($warehouse)
                                           + abs($shipping) + abs($otherSupport) + abs($netOther)
                                           + abs($tcs) + abs($tds) + abs($gstOnOther);
                $rb['final_cash_in_hand'] += $finalSettle;
            }

            // Profit Calculator
            $pc = &$r['profit_calc'];

            if ($statusType === 'delivered') {
                $pc['delivered_orders']++;
                $pc['delivered_qty'] += $qty;
                $pc['my_selling_total']     += $myPrice * $qty;
                $pc['meesho_selling_total'] += $meeshoPrice * $qty;
                $pc['gross_profit_margin']  += ($meeshoPrice - $myPrice) * $qty;

                $meeshoCut = abs($commission) + abs($goldFee) + abs($mallFee);
                $otherFee  = abs($fixedFee) + abs($warehouse) + abs($otherSupport) + abs($netOther);
                $pc['meesho_commission_cut'] += $meeshoCut;
                $pc['meesho_other_fee_cut']  += $otherFee;
                $pc['meesho_total_earned']   += $meeshoCut + $otherFee;

                $pc['shipping_charged'] += abs($shipping);
                $pc['tax_deducted']     += abs($tcs) + abs($tds) + abs($gstOnOther);
                $pc['net_you_earned']   += $finalSettle;
            }
            elseif ($statusType === 'return') {
                $pc['return_orders']++;
                $pc['return_qty'] += $qty;
                $pc['return_loss'] += abs($saleReturn);
                $pc['return_shipping_cut'] += abs($retShipping);
            }
            elseif ($statusType === 'rto') {
                $pc['rto_orders']++;
                $pc['rto_qty'] += $qty;
                $pc['rto_loss'] += abs($saleReturn);
                $pc['rto_shipping_cut'] += abs($retShipping);
            }

            $sk = $statusRaw ?: 'Unknown';
            $r['status_breakdown'][$sk] = ($r['status_breakdown'][$sk] ?? 0) + 1;
            $r['top_products'][$product] = ($r['top_products'][$product] ?? 0) + $qty;
        }

        $r['total_orders'] = count($uniqueOrders);
    }

    // ========== 2. ADS COST ==========
    foreach ($sheets as $name => $srows) {
        if (stripos($name, 'ads') !== false) {
            $hIdx = 0;
            foreach ($srows as $i => $row) {
                foreach ($row as $c) {
                    $cs = strtolower(trim((string)$c));
                    if ($cs === 'deduction date' || $cs === 'ad cost') { $hIdx = $i; break 2; }
                }
            }
            foreach (array_slice($srows, $hIdx + 1) as $row) {
                $ne = array_filter($row, fn($v) => $v !== null && $v !== '');
                if (count($ne) < 2) continue;
                $first = trim((string)($row[0] ?? ''));
                if ($first === '' || preg_match('/^[A-Z]\s*[=+\-]/', $first)) continue;

                $r['ads_rows']++;
                $r['ads_cost']      += abs(cleanNumber($row[3] ?? 0));
                $r['ads_cost_incl'] += abs(cleanNumber($row[5] ?? 0));
                $r['ads_gst']       += abs(cleanNumber($row[6] ?? 0));
                $r['ads_total']     += abs(cleanNumber($row[7] ?? 0));
            }
            break;
        }
    }

    // ========== 3. REFERRAL ==========
    foreach ($sheets as $name => $rows) {
        if (stripos($name, 'referral') !== false) {
            foreach ($rows as $row) {
                if (empty($row)) continue;
                $first = strtolower((string)($row[0] ?? ''));
                if (strpos($first, 'no data') !== false) break;
                if (strpos($first, 'reward id') !== false) continue;
                if (strpos($first, 'referral payment') !== false) continue;

                $net = cleanNumber($row[4] ?? 0);
                $tax = cleanNumber($row[5] ?? 0);
                if ($net != 0 || $tax != 0) {
                    $r['referral_rows']++;
                    $r['referral_net'] += $net;
                    $r['referral_tax'] += $tax;
                }
            }
            break;
        }
    }

    // ========== 4. COMPENSATION ==========
    foreach ($sheets as $name => $rows) {
        if (stripos($name, 'compensation') !== false) {
            foreach ($rows as $row) {
                if (empty($row)) continue;
                $first = strtolower((string)($row[0] ?? ''));
                if (strpos($first, 'no data') !== false) break;
                if (strpos($first, 'date') !== false) continue;
                if (strpos($first, 'platform recovery') !== false) continue;

                $amt = cleanNumber($row[3] ?? 0);
                if ($amt != 0) {
                    $r['platform_comp_rows']++;
                    $r['platform_compensation'] += $amt;
                }
            }
            break;
        }
    }

    arsort($r['top_products']);
    $r['top_products'] = array_slice($r['top_products'], 0, 10, true);

    // ⭐ FINAL CALCULATION (STEP BY STEP) ⭐
    $rb = &$r['revenue_breakdown'];
    $pc = &$r['profit_calc'];

    $netSales = $rb['net_product_revenue'];
    $adsCost = abs($r['ads_total']);
    $finalCash = $netSales - $adsCost;
    $returnCharges = $rb['rto_courier_charges'] + $rb['customer_return_shipping'];
    $platformFees = $rb['platform_fees_total'];
    $costPrice = $rb['cost_price_total'];

    $rb['final_cash_in_hand'] = $finalCash;
    $rb['total_profit_loss'] = $finalCash - $returnCharges - $platformFees - $costPrice;

    $pc['ads_cost'] = $r['ads_cost'];
    $pc['ads_gst']  = $r['ads_gst'];
    $pc['total_ads'] = $r['ads_total'];
    $pc['final_in_hand'] = $finalCash;
    $pc['net_profit_loss'] = $rb['total_profit_loss'];

    if ($pc['delivered_qty'] > 0) {
        $pc['per_unit_profit'] = $pc['net_profit_loss'] / $pc['delivered_qty'];
    }

    $r['total_commission'] = $r['meesho_commission'] + $r['meesho_gold_fee'] + $r['meesho_mall_fee'];
    $r['total_deductions'] = abs($r['meesho_commission']) + abs($r['meesho_gold_fee']) + abs($r['meesho_mall_fee'])
                           + abs($r['fixed_fee']) + abs($r['warehousing_fee'])
                           + abs($r['shipping_charge']) + abs($r['return_shipping'])
                           + abs($r['return_premium']) + abs($r['other_support'])
                           + abs($r['net_other_support']);
    $r['total_taxes'] = abs($r['tcs']) + abs($r['tds']) + abs($r['gst_on_other']);

    return $r;
}

// ================================================================
//  PROCESS UPLOAD
// ================================================================
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['meesho_file'])) {
    $file = $_FILES['meesho_file'];

    if ($file['error'] !== UPLOAD_ERR_OK) {
        $error = 'Upload failed (code ' . $file['error'] . ')';
    } else {
        $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
        if (!in_array($ext, ['xlsx', 'xls', 'csv'])) {
            $error = 'Invalid format. Please upload .xlsx, .xls, or .csv';
        } else {
            $safeName = 'meesho_' . date('Ymd_His') . '_' . uniqid() . '.' . $ext;
            $target   = $uploadDir . $safeName;

            if (move_uploaded_file($file['tmp_name'], $target)) {
                try {
                    if ($ext === 'xlsx') {
                        $report = processMeeshoFile($target);
                    } else {
                        throw new Exception('Please upload .xlsx file (Meesho multi-sheet format required).');
                    }
                    @unlink($target);
                } catch (Exception $e) {
                    $error = 'Processing error: ' . $e->getMessage();
                }
            } else {
                $error = 'Failed to save uploaded file.';
            }
        }
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Meesho Premium Report</title>
<style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        background: #0f1419;
        background-image:
            radial-gradient(at 0% 0%, rgba(196, 69, 105, 0.18) 0px, transparent 50%),
            radial-gradient(at 100% 100%, rgba(255, 107, 107, 0.15) 0px, transparent 50%);
        min-height: 100vh;
        padding: 24px;
        color: #e2e8f0;
        -webkit-font-smoothing: antialiased;
    }
    .container { max-width: 1400px; margin: 0 auto; }

    /* HEADER */
    .header {
        background: linear-gradient(135deg, rgba(45, 55, 72, 0.6), rgba(26, 32, 44, 0.6));
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 20px;
        padding: 32px;
        text-align: center;
        margin-bottom: 24px;
    }
    .header h1 {
        font-size: 28px;
        font-weight: 800;
        background: linear-gradient(135deg, #fbb6ce, #ff6b6b, #c44569);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 8px;
        letter-spacing: -0.5px;
    }
    .header p { color: #94a3b8; font-size: 14px; }
    .meesho-badge {
        display: inline-block;
        background: linear-gradient(135deg, #c44569, #ff6b6b);
        color: #fff;
        padding: 6px 18px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 800;
        margin-top: 12px;
        letter-spacing: 2px;
    }

    /* CARD */
    .card {
        background: rgba(30, 41, 59, 0.6);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 20px;
        padding: 28px;
        margin-bottom: 24px;
    }

    /* UPLOAD AREA */
    .upload-area {
        border: 2px dashed rgba(255, 107, 107, 0.4);
        border-radius: 16px;
        padding: 44px 24px;
        text-align: center;
        cursor: pointer;
        background: rgba(196, 69, 105, 0.05);
        transition: all 0.3s ease;
    }
    .upload-area:hover {
        border-color: #ff6b6b;
        background: rgba(196, 69, 105, 0.12);
        transform: translateY(-2px);
    }
    .upload-area input[type="file"] { display: none; }
    .upload-icon { font-size: 48px; margin-bottom: 12px; display: block; }
    .upload-title { color: #e2e8f0; font-weight: 700; font-size: 16px; margin-bottom: 6px; }
    .upload-sub { color: #64748b; font-size: 13px; }

    .btn {
        background: linear-gradient(135deg, #ff6b6b, #c44569);
        color: #fff;
        border: none;
        padding: 14px 48px;
        border-radius: 12px;
        font-size: 15px;
        font-weight: 700;
        cursor: pointer;
        transition: all 0.25s ease;
        box-shadow: 0 8px 24px rgba(196, 69, 105, 0.35);
    }
    .btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 32px rgba(196, 69, 105, 0.55);
    }

    .error {
        background: rgba(245, 101, 101, 0.15);
        color: #feb2b2;
        padding: 16px 20px;
        border-radius: 12px;
        margin-bottom: 24px;
        border-left: 4px solid #f56565;
        font-family: 'SF Mono', monospace;
        font-size: 13px;
        white-space: pre-wrap;
    }

    /* HERO */
    .hero {
        background: linear-gradient(135deg, #1a202c, #2d3748);
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 24px;
        padding: 48px 32px;
        text-align: center;
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
    }
    .hero::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        background: linear-gradient(90deg, #ff6b6b, #c44569, #ed8936, #48bb78);
    }
    .hero .hero-icon { font-size: 56px; margin-bottom: 12px; display: block; }
    .hero .hero-label {
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: 4px;
        color: #94a3b8;
        font-weight: 700;
        margin-bottom: 16px;
    }
    .hero .hero-value {
        font-size: 72px;
        font-weight: 900;
        letter-spacing: -3px;
        font-variant-numeric: tabular-nums;
        line-height: 1;
        margin: 16px 0;
    }
    .hero .hero-value.positive {
        background: linear-gradient(135deg, #48bb78, #38a169);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .hero .hero-value.negative {
        background: linear-gradient(135deg, #f56565, #e53e3e);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .hero .hero-sub {
        color: #94a3b8;
        font-size: 14px;
        margin-top: 16px;
        font-weight: 500;
    }

    /* SECTION TITLE */
    .section-title {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 20px;
        padding-bottom: 16px;
        border-bottom: 1px solid rgba(255,255,255,0.08);
    }
    .section-title .st-icon {
        width: 42px; height: 42px;
        border-radius: 12px;
        background: linear-gradient(135deg, #ff6b6b, #c44569);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 20px;
    }
    .section-title h2 {
        color: #e2e8f0;
        font-size: 20px;
        font-weight: 800;
        letter-spacing: -0.3px;
    }

    /* FLOW STEPS */
    .flow-step {
        display: flex;
        align-items: center;
        padding: 20px 24px;
        border-radius: 14px;
        margin-bottom: 12px;
        gap: 20px;
        transition: all 0.2s ease;
    }
    .flow-step:hover { transform: translateX(4px); }
    .flow-step .fs-badge {
        flex-shrink: 0;
        width: 48px; height: 48px;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 20px;
        font-weight: 900;
    }
    .flow-step .fs-content { flex: 1; min-width: 0; }
    .flow-step .fs-title {
        font-size: 15px;
        font-weight: 700;
        margin-bottom: 4px;
        color: #f1f5f9;
    }
    .flow-step .fs-desc {
        font-size: 13px;
        color: #94a3b8;
        line-height: 1.5;
    }
    .flow-step .fs-amount {
        flex-shrink: 0;
        text-align: right;
        font-family: 'SF Mono', 'Courier New', monospace;
        font-weight: 800;
        font-size: 22px;
        font-variant-numeric: tabular-nums;
    }
    .flow-step .fs-sub {
        font-size: 11px;
        color: #64748b;
        margin-top: 4px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .flow-step.income {
        background: linear-gradient(135deg, rgba(72, 187, 120, 0.1), rgba(72, 187, 120, 0.05));
        border: 1px solid rgba(72, 187, 120, 0.2);
    }
    .flow-step.income .fs-badge { background: rgba(72, 187, 120, 0.2); color: #68d391; }
    .flow-step.income .fs-amount { color: #68d391; }

    .flow-step.expense {
        background: linear-gradient(135deg, rgba(245, 101, 101, 0.08), rgba(245, 101, 101, 0.03));
        border: 1px solid rgba(245, 101, 101, 0.15);
    }
    .flow-step.expense .fs-badge { background: rgba(245, 101, 101, 0.2); color: #fc8181; }
    .flow-step.expense .fs-amount { color: #fc8181; }

    .flow-step.neutral {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.06);
    }
    .flow-step.neutral .fs-badge { background: rgba(251, 182, 206, 0.15); color: #fbb6ce; }
    .flow-step.neutral .fs-amount { color: #cbd5e0; }

    .flow-step.highlight {
        background: linear-gradient(135deg, rgba(196, 69, 105, 0.15), rgba(196, 69, 105, 0.08));
        border: 1px solid rgba(196, 69, 105, 0.3);
        box-shadow: 0 4px 20px rgba(196, 69, 105, 0.1);
    }
    .flow-step.highlight .fs-badge { background: rgba(196, 69, 105, 0.3); color: #fbb6ce; }
    .flow-step.highlight .fs-amount { color: #fbb6ce; }
    .flow-step.highlight .fs-title { color: #fbb6ce; }

    .flow-step.final {
        background: linear-gradient(135deg, #1a202c, #2d3748);
        border: 2px solid rgba(255,255,255,0.15);
        padding: 32px 28px;
        margin-top: 20px;
        box-shadow: 0 12px 40px rgba(0,0,0,0.4);
    }
    .flow-step.final .fs-badge { width: 60px; height: 60px; font-size: 28px; }
    .flow-step.final .fs-title {
        font-size: 18px;
        font-weight: 900;
        letter-spacing: 1px;
        text-transform: uppercase;
    }
    .flow-step.final .fs-amount { font-size: 40px; letter-spacing: -1.5px; }
    .flow-step.final.profit { border-color: rgba(72, 187, 120, 0.5); }
    .flow-step.final.profit .fs-badge { background: rgba(72, 187, 120, 0.25); }
    .flow-step.final.profit .fs-amount { color: #68d391; }
    .flow-step.final.loss { border-color: rgba(245, 101, 101, 0.5); }
    .flow-step.final.loss .fs-badge { background: rgba(245, 101, 101, 0.25); }
    .flow-step.final.loss .fs-amount { color: #fc8181; }

    /* STATS GRID */
    .stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
        gap: 16px;
        margin-bottom: 16px;
    }
    .stat-card {
        padding: 20px;
        border-radius: 14px;
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.06);
    }
    .stat-card.green { border-left: 4px solid #48bb78; }
    .stat-card.red { border-left: 4px solid #f56565; }
    .stat-card.orange { border-left: 4px solid #ed8936; }
    .stat-card.blue { border-left: 4px solid #4299e1; }
    .stat-card.purple { border-left: 4px solid #9f7aea; }
    .stat-card.gray { border-left: 4px solid #718096; }
    .stat-card .label {
        font-size: 11px;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .stat-card .value {
        font-size: 24px;
        font-weight: 800;
        color: #f1f5f9;
        font-variant-numeric: tabular-nums;
    }
    .stat-card .sub {
        font-size: 12px;
        color: #64748b;
        margin-top: 4px;
    }

    /* TABLE */
    table { width: 100%; border-collapse: collapse; font-size: 14px; }
    th {
        background: rgba(15, 20, 25, 0.8);
        color: #94a3b8;
        padding: 14px 16px;
        text-align: left;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 800;
        border-bottom: 2px solid rgba(196, 69, 105, 0.3);
    }
    td {
        padding: 12px 16px;
        border-bottom: 1px solid rgba(255,255,255,0.05);
        color: #cbd5e0;
        font-variant-numeric: tabular-nums;
    }
    tr:hover td { background: rgba(255,255,255,0.02); }
    td.num {
        text-align: right;
        font-family: 'SF Mono', 'Courier New', monospace;
        font-weight: 600;
    }
    td.num.positive { color: #68d391; }
    td.num.negative { color: #fc8181; }
    tr.total-row td {
        background: rgba(196, 69, 105, 0.1);
        font-weight: 800;
        border-top: 2px solid rgba(196, 69, 105, 0.4);
        border-bottom: 2px solid rgba(196, 69, 105, 0.4);
        color: #f1f5f9;
    }
    tr.section-header td {
        background: rgba(196, 69, 105, 0.08);
        color: #fbb6ce;
        font-weight: 800;
        text-transform: uppercase;
        font-size: 11px;
        letter-spacing: 2px;
        padding: 12px 16px;
    }

    .badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 800;
        background: rgba(196, 69, 105, 0.2);
        color: #fbb6ce;
        letter-spacing: 0.5px;
    }

    .sheet-list { margin-top: 12px; }
    .sheet-list span {
        display: inline-block;
        background: rgba(255,255,255,0.06);
        color: #cbd5e0;
        padding: 5px 14px;
        border-radius: 20px;
        font-size: 12px;
        margin: 4px 4px 0 0;
        font-weight: 600;
    }

    .footer-note {
        text-align: center;
        color: #64748b;
        font-size: 12px;
        margin-top: 24px;
        padding: 20px;
    }

    @media (max-width: 768px) {
        body { padding: 12px; }
        .hero .hero-value { font-size: 48px; letter-spacing: -2px; }
        .hero .hero-icon { font-size: 40px; }
        .flow-step { padding: 16px; flex-wrap: wrap; gap: 12px; }
        .flow-step .fs-amount { font-size: 18px; width: 100%; text-align: left; padding-left: 68px; }
        .flow-step.final .fs-amount { font-size: 28px; padding-left: 0; }
        .flow-step .fs-badge { width: 40px; height: 40px; font-size: 16px; }
        .section-title h2 { font-size: 16px; }
        .card { padding: 20px; }
    }
</style>
</head>
<body>
<div class="container">
    <!-- HEADER -->
    <div class="header">
        <h1>💎 Meesho Premium Report</h1>
        <p>Every product · Every charge · Every rupee — clearly explained</p>
        <span class="meesho-badge">MEESHO SUPPLIER</span>
    </div>

    <?php if ($error): ?>
        <div class="error">⚠️ <?= htmlspecialchars($error) ?></div>
    <?php endif; ?>

    <!-- UPLOAD -->
    <div class="card">
        <form method="POST" enctype="multipart/form-data">
            <div class="upload-area" onclick="document.getElementById('fileInput').click()">
                <span class="upload-icon">📁</span>
                <div class="upload-title">Upload your Meesho payment file</div>
                <div class="upload-sub">.xlsx — no library required</div>
                <input type="file" name="meesho_file" id="fileInput"
                       accept=".xlsx,.xls,.csv" required onclick="event.stopPropagation()">
            </div>
            <div style="text-align: center; margin-top: 24px;">
                <button type="submit" class="btn">💰 Generate Premium Report</button>
            </div>
        </form>
    </div>

    <?php if ($report !== null):
        $pc = $report['profit_calc'];
        $rb = $report['revenue_breakdown'];

        // Final calculation
        $netSales = $rb['net_product_revenue'];
        $adsCost = abs($report['ads_total']);
        $finalCash = $netSales - $adsCost;
        $returnCharges = $rb['rto_courier_charges'] + $rb['customer_return_shipping'];
        $platformFees = $rb['platform_fees_total'];
        $costPrice = $rb['cost_price_total'];
        $finalProfit = $finalCash - $returnCharges - $platformFees - $costPrice;
        $isProfit = $finalProfit >= 0;
    ?>

        <!-- FILE INFO -->
        <div class="card">
            <div class="section-title">
                <div class="st-icon">📁</div>
                <h2>File Information</h2>
            </div>
            <p style="color:#94a3b8; margin-bottom:12px;"><strong style="color:#f1f5f9;">File:</strong> <?= htmlspecialchars($report['filename']) ?></p>
            <div class="sheet-list">
                <strong style="color:#94a3b8; font-size:13px;">Sheets found:</strong><br>
                <?php foreach ($report['sheet_names'] as $sn): ?>
                    <span><?= htmlspecialchars($sn) ?></span>
                <?php endforeach; ?>
            </div>
        </div>

        <!-- HERO PROFIT -->
        <div class="hero">
            <span class="hero-icon"><?= $isProfit ? '💰' : '📉' ?></span>
            <div class="hero-label">My Final <?= $isProfit ? 'Profit' : 'Loss' ?></div>
            <div class="hero-value <?= $isProfit ? 'positive' : 'negative' ?>">
                ₹<?= number_format(abs($finalProfit), 2) ?>
            </div>
            <div class="hero-sub">
                Net Sales ₹<?= number_format($netSales, 2) ?>
                − Ads ₹<?= number_format($adsCost, 2) ?>
                − Return Charges ₹<?= number_format($returnCharges, 2) ?>
                − Platform Fees ₹<?= number_format($platformFees, 2) ?>
                <?php if ($pc['per_unit_profit'] != 0): ?>
                    <br>Per delivered unit: ₹<?= number_format($pc['per_unit_profit'], 2) ?>
                <?php endif; ?>
            </div>
        </div>

        <!-- ⭐ SALES & REVENUE BREAKDOWN (STEP BY STEP) ⭐ -->
        <div class="card">
            <div class="section-title">
                <div class="st-icon">📊</div>
                <h2>Sales & Revenue Breakdown</h2>
            </div>

            <!-- 1: GROSS SALES -->
            <div class="flow-step income">
                <div class="fs-badge">1</div>
                <div class="fs-content">
                    <div class="fs-title">Gross Sales</div>
                    <div class="fs-desc">
                        <?= number_format($rb['gross_sales_pieces']) ?> pieces ordered × your selling price
                    </div>
                </div>
                <div class="fs-amount">
                    ₹<?= number_format($rb['gross_sales_value'], 2) ?>
                    <div class="fs-sub">Total ordered value</div>
                </div>
            </div>

            <!-- 2: LESS RTO -->
            <div class="flow-step expense">
                <div class="fs-badge">2</div>
                <div class="fs-content">
                    <div class="fs-title">Less: RTO (Return to Origin)</div>
                    <div class="fs-desc">
                        <?= number_format($rb['rto_pieces']) ?> pieces returned · not delivered to customer
                    </div>
                </div>
                <div class="fs-amount">
                    − ₹<?= number_format($rb['rto_value'], 2) ?>
                    <div class="fs-sub">RTO value</div>
                </div>
            </div>

            <!-- 3: NET PRODUCT REVENUE -->
            <div class="flow-step highlight">
                <div class="fs-badge">=</div>
                <div class="fs-content">
                    <div class="fs-title">Net Product Revenue</div>
                    <div class="fs-desc">
                        <?= number_format($rb['net_sold_pieces']) ?> pieces kept by customers
                    </div>
                </div>
                <div class="fs-amount">
                    ₹<?= number_format($netSales, 2) ?>
                    <div class="fs-sub">Gross − RTO</div>
                </div>
            </div>

            <!-- RETURNED STOCK -->
            <div class="flow-step neutral">
                <div class="fs-badge">📦</div>
                <div class="fs-content">
                    <div class="fs-title">Returned Stock (Back in Inventory)</div>
                    <div class="fs-desc">
                        <?= number_format($rb['returned_stock_pieces']) ?> pieces safe in inventory ·
                        <strong style="color:#68d391;">NOT a loss</strong>
                    </div>
                </div>
                <div class="fs-amount" style="color:#94a3b8;">
                    ₹<?= number_format($rb['returned_stock_value'], 2) ?>
                    <div class="fs-sub">Preserved value</div>
                </div>
            </div>

            <!-- 4: LESS ADS -->
            <div class="flow-step expense">
                <div class="fs-badge">3</div>
                <div class="fs-content">
                    <div class="fs-title">Less: Ads Cost</div>
                    <div class="fs-desc">
                        Marketing & advertising spend · <?= $report['ads_rows'] ?> entries
                    </div>
                </div>
                <div class="fs-amount">
                    − ₹<?= number_format($adsCost, 2) ?>
                    <div class="fs-sub">Ads deduction</div>
                </div>
            </div>

            <!-- 5: FINAL CASH IN HAND -->
            <div class="flow-step highlight">
                <div class="fs-badge">=</div>
                <div class="fs-content">
                    <div class="fs-title">💰 Final Cash in Hand</div>
                    <div class="fs-desc">
                        Net Sales − Ads Cost · Actual money received
                    </div>
                </div>
                <div class="fs-amount">
                    ₹<?= number_format($finalCash, 2) ?>
                    <div class="fs-sub">Money in hand</div>
                </div>
            </div>

            <!-- 6: LESS RETURN CHARGES -->
            <div class="flow-step expense">
                <div class="fs-badge">4</div>
                <div class="fs-content">
                    <div class="fs-title">Less: Return Charges</div>
                    <div class="fs-desc">
                        RTO courier ₹<?= number_format($rb['rto_courier_charges'], 2) ?> +
                        Customer return ₹<?= number_format($rb['customer_return_shipping'], 2) ?>
                    </div>
                </div>
                <div class="fs-amount">
                    − ₹<?= number_format($returnCharges, 2) ?>
                    <div class="fs-sub">Courier charges</div>
                </div>
            </div>

            <!-- 7: LESS PLATFORM FEES -->
            <div class="flow-step expense">
                <div class="fs-badge">5</div>
                <div class="fs-content">
                    <div class="fs-title">Less: Platform Fees & Commission</div>
                    <div class="fs-desc">
                        Commission + Fixed + Gold + Mall + Warehousing + Shipping + Taxes
                    </div>
                </div>
                <div class="fs-amount">
                    − ₹<?= number_format($platformFees, 2) ?>
                    <div class="fs-sub">Total platform charges</div>
                </div>
            </div>

            <!-- 8: LESS COST -->
            <?php if ($costPrice > 0): ?>
            <div class="flow-step expense">
                <div class="fs-badge">6</div>
                <div class="fs-content">
                    <div class="fs-title">Less: Cost of Sold Items</div>
                    <div class="fs-desc">Your purchase/manufacturing cost</div>
                </div>
                <div class="fs-amount">
                    − ₹<?= number_format($costPrice, 2) ?>
                    <div class="fs-sub">Your COGS</div>
                </div>
            </div>
            <?php endif; ?>

            <!-- 9: FINAL PROFIT/LOSS -->
            <div class="flow-step final <?= $isProfit ? 'profit' : 'loss' ?>">
                <div class="fs-badge">🏁</div>
                <div class="fs-content">
                    <div class="fs-title">Total <?= $isProfit ? 'Profit' : 'Loss' ?></div>
                    <div class="fs-desc">
                        <?= $isProfit ? '✅ You made a PROFIT — Great job!' : '❌ You made a LOSS — Review pricing' ?>
                    </div>
                </div>
                <div class="fs-amount">
                    ₹<?= number_format(abs($finalProfit), 2) ?>
                    <div class="fs-sub"><?= $isProfit ? 'Net Profit' : 'Net Loss' ?></div>
                </div>
            </div>
        </div>

        <!-- PRODUCT COUNTS -->
        <div class="card">
            <div class="section-title">
                <div class="st-icon">📦</div>
                <h2>Product Counts</h2>
            </div>
            <div class="stats-grid">
                <div class="stat-card green">
                    <div class="label">✅ Sold</div>
                    <div class="value"><?= number_format($report['sold']) ?></div>
                    <div class="sub"><?= number_format($report['sold_qty']) ?> units</div>
                </div>
                <div class="stat-card blue">
                    <div class="label">📦 Delivered</div>
                    <div class="value"><?= number_format($report['delivered']) ?></div>
                    <div class="sub"><?= number_format($report['delivered_qty']) ?> units</div>
                </div>
                <div class="stat-card purple">
                    <div class="label">🚚 Shipped</div>
                    <div class="value"><?= number_format($report['shipped']) ?></div>
                    <div class="sub"><?= number_format($report['shipped_qty']) ?> units</div>
                </div>
                <div class="stat-card orange">
                    <div class="label">⏳ Pending</div>
                    <div class="value"><?= number_format($report['pending']) ?></div>
                    <div class="sub"><?= number_format($report['pending_qty']) ?> units</div>
                </div>
                <div class="stat-card red">
                    <div class="label">↩️ Returns</div>
                    <div class="value"><?= number_format($report['return']) ?></div>
                    <div class="sub"><?= number_format($report['return_qty']) ?> units</div>
                </div>
                <div class="stat-card red">
                    <div class="label">🔙 RTO</div>
                    <div class="value"><?= number_format($report['rto']) ?></div>
                    <div class="sub"><?= number_format($report['rto_qty']) ?> units</div>
                </div>
                <div class="stat-card gray">
                    <div class="label">❌ Cancelled</div>
                    <div class="value"><?= number_format($report['cancelled']) ?></div>
                    <div class="sub"><?= number_format($report['cancelled_qty']) ?> units</div>
                </div>
                <div class="stat-card gray">
                    <div class="label">📊 Total Rows</div>
                    <div class="value"><?= number_format($report['total_rows']) ?></div>
                    <div class="sub"><?= number_format($report['total_orders']) ?> unique</div>
                </div>
            </div>
        </div>

        <!-- RETURNS & RTO SUMMARY -->
        <div class="card">
            <div class="section-title">
                <div class="st-icon">↩️</div>
                <h2>Returns & RTO Summary</h2>
            </div>
            <div class="stats-grid">
                <div class="stat-card red">
                    <div class="label">↩️ Customer Returns</div>
                    <div class="value"><?= number_format($pc['return_orders']) ?></div>
                    <div class="sub"><?= number_format($pc['return_qty']) ?> units · ₹<?= number_format(abs($pc['return_loss']), 2) ?></div>
                </div>
                <div class="stat-card orange">
                    <div class="label">🔙 RTO Orders</div>
                    <div class="value"><?= number_format($pc['rto_orders']) ?></div>
                    <div class="sub"><?= number_format($pc['rto_qty']) ?> units · ₹<?= number_format(abs($pc['rto_loss']), 2) ?></div>
                </div>
                <div class="stat-card orange">
                    <div class="label">🚚 Return Courier Charges</div>
                    <div class="value">₹<?= number_format($returnCharges, 2) ?></div>
                    <div class="sub">RTO + Customer return shipping</div>
                </div>
                <div class="stat-card green">
                    <div class="label">📦 Returned Stock Value</div>
                    <div class="value">₹<?= number_format($rb['returned_stock_value'], 2) ?></div>
                    <div class="sub">Safe in inventory (not a loss)</div>
                </div>
            </div>
        </div>

        <!-- FULL MONEY FLOW TABLE -->
        <div class="card">
            <div class="section-title">
                <div class="st-icon">💵</div>
                <h2>Full Money Flow — Penny by Penny</h2>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 55%;">Item</th>
                        <th style="width: 20%; text-align: right;">Amount (₹)</th>
                        <th style="width: 25%;">Notes</th>
                    </tr>
                </thead>
                <tbody>
                    <tr class="section-header"><td colspan="3">💵 Revenue</td></tr>
                    <tr><td>Your Listing Price (Total)</td><td class="num positive">₹<?= number_format($report['listing_price'], 2) ?></td><td>Your selling price</td></tr>
                    <tr><td>Meesho's Price (Total)</td><td class="num positive">₹<?= number_format($report['total_sale_amount'], 2) ?></td><td>Customer paid Meesho</td></tr>
                    <tr><td>Your Gross Margin</td><td class="num <?= $pc['gross_profit_margin'] >= 0 ? 'positive' : 'negative' ?>">₹<?= number_format($pc['gross_profit_margin'], 2) ?></td><td>Meesho − Yours</td></tr>

                    <tr class="section-header"><td colspan="3">🔻 Meesho's Cuts</td></tr>
                    <tr><td>Meesho Commission</td><td class="num negative">₹<?= number_format($report['meesho_commission'], 2) ?></td><td>Category commission</td></tr>
                    <tr><td>Meesho Gold Fee</td><td class="num negative">₹<?= number_format($report['meesho_gold_fee'], 2) ?></td><td>Gold program</td></tr>
                    <tr><td>Meesho Mall Fee</td><td class="num negative">₹<?= number_format($report['meesho_mall_fee'], 2) ?></td><td>Mall program</td></tr>
                    <tr><td>Fixed Fee</td><td class="num negative">₹<?= number_format($report['fixed_fee'], 2) ?></td><td>Per-order</td></tr>
                    <tr><td>Warehousing Fee</td><td class="num negative">₹<?= number_format($report['warehousing_fee'], 2) ?></td><td>Storage</td></tr>
                    <tr><td>Other Support</td><td class="num negative">₹<?= number_format($report['other_support'], 2) ?></td><td>Support services</td></tr>

                    <tr class="section-header"><td colspan="3">🚚 Shipping Cuts</td></tr>
                    <tr><td>Forward Shipping</td><td class="num negative">₹<?= number_format($report['shipping_charge'], 2) ?></td><td>Delivery</td></tr>
                    <tr><td>Return Shipping</td><td class="num negative">₹<?= number_format($report['return_shipping'], 2) ?></td><td>Return pickup</td></tr>
                    <tr><td>Return Premium</td><td class="num negative">₹<?= number_format($report['return_premium'], 2) ?></td><td>Return protection</td></tr>

                    <tr class="section-header"><td colspan="3">📊 Taxes</td></tr>
                    <tr><td>TCS (1%)</td><td class="num negative">₹<?= number_format($report['tcs'], 2) ?></td><td>Tax collected</td></tr>
                    <tr><td>TDS (5%)</td><td class="num negative">₹<?= number_format($report['tds'], 2) ?></td><td>Tax deducted</td></tr>
                    <tr><td>GST on Other</td><td class="num negative">₹<?= number_format($report['gst_on_other'], 2) ?></td><td>18% GST</td></tr>

                    <tr class="section-header"><td colspan="3">💸 Ads</td></tr>
                    <tr><td>Total Ad Cost</td><td class="num negative">₹<?= number_format($report['ads_cost'], 2) ?></td><td>Gross ad spend</td></tr>
                    <tr><td>GST on Ads</td><td class="num negative">₹<?= number_format($report['ads_gst'], 2) ?></td><td>18% GST</td></tr>
                    <tr class="total-row"><td>💥 TOTAL ADS</td><td class="num negative">₹<?= number_format($report['ads_total'], 2) ?></td><td><?= $report['ads_rows'] ?> entries</td></tr>

                    <tr class="section-header"><td colspan="3">🎁 Credits</td></tr>
                    <tr><td>Compensation</td><td class="num positive">₹<?= number_format($report['compensation'], 2) ?></td><td>Received</td></tr>
                    <tr><td>Claims</td><td class="num positive">₹<?= number_format($report['claims'], 2) ?></td><td>From Meesho</td></tr>
                    <tr><td>Waivers</td><td class="num positive">₹<?= number_format($report['waivers'], 2) ?></td><td>Fee waivers</td></tr>
                    <tr><td>Referral Payments</td><td class="num positive">₹<?= number_format($report['referral_net'], 2) ?></td><td><?= $report['referral_rows'] ?> entries</td></tr>

                    <tr class="section-header"><td colspan="3">🏆 FINAL RESULT</td></tr>
                    <tr class="total-row">
                        <td>✅ FINAL SETTLEMENT (Money Received)</td>
                        <td class="num positive">₹<?= number_format($report['final_settlement'], 2) ?></td>
                        <td>Actual bank payout</td>
                    </tr>
                    <tr class="total-row" style="background: <?= $isProfit ? 'rgba(72,187,120,0.15)' : 'rgba(245,101,101,0.15)' ?>;">
                        <td>💰 <strong>NET <?= $isProfit ? 'PROFIT' : 'LOSS' ?></strong></td>
                        <td class="num <?= $isProfit ? 'positive' : 'negative' ?>">₹<?= number_format(abs($finalProfit), 2) ?></td>
                        <td>After all deductions</td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- STATUS BREAKDOWN -->
        <?php if (!empty($report['status_breakdown'])): ?>
        <div class="card">
            <div class="section-title">
                <div class="st-icon">📋</div>
                <h2>Status Breakdown</h2>
            </div>
            <table>
                <thead><tr><th>Status</th><th>Count</th><th>Percentage</th></tr></thead>
                <tbody>
                    <?php $total = $report['total_rows']; foreach ($report['status_breakdown'] as $st => $c): ?>
                        <tr>
                            <td><span class="badge"><?= htmlspecialchars($st) ?></span></td>
                            <td><?= number_format($c) ?></td>
                            <td><?= $total > 0 ? round(($c / $total) * 100, 2) : 0 ?>%</td>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
        </div>
        <?php endif; ?>

        <!-- TOP PRODUCTS -->
        <?php if (!empty($report['top_products'])): ?>
        <div class="card">
            <div class="section-title">
                <div class="st-icon">🏆</div>
                <h2>Top 10 Products</h2>
            </div>
            <table>
                <thead><tr><th>#</th><th>Product</th><th>Quantity</th></tr></thead>
                <tbody>
                    <?php $rank = 1; foreach ($report['top_products'] as $p => $q): ?>
                        <tr>
                            <td><?= $rank++ ?></td>
                            <td><?= htmlspecialchars($p) ?></td>
                            <td><?= number_format($q) ?> units</td>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
        </div>
        <?php endif; ?>

    <?php endif; ?>

    <div class="footer-note">
        ✨ Meesho Premium Report · Sales → RTO → Net Sales → Ads → Cash → Fees → Profit
    </div>
</div>

<script>
const ua = document.querySelector('.upload-area');
const fi = document.getElementById('fileInput');
['dragenter','dragover'].forEach(e => ua.addEventListener(e, ev => {
    ev.preventDefault();
    ua.style.borderColor = '#ff6b6b';
    ua.style.background = 'rgba(196, 69, 105, 0.15)';
}));
['dragleave','drop'].forEach(e => ua.addEventListener(e, ev => {
    ev.preventDefault();
    ua.style.borderColor = 'rgba(255, 107, 107, 0.4)';
    ua.style.background = 'rgba(196, 69, 105, 0.05)';
}));
ua.addEventListener('drop', e => {
    if (e.dataTransfer.files.length) fi.files = e.dataTransfer.files;
});
</script>
</body>
</html>